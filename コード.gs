/**
 * For Two 家計簿システム — GAS バックエンド
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('For Two 家計簿')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const DATA_SHEET_NAME     = "シート1";
const SETTINGS_SHEET_NAME = "設定";

// ─── データ取得 ────────────────────────────────────────────
function getReportData(month) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) throw new Error("シート「" + DATA_SHEET_NAME + "」が見つかりません。");

  const data = sheet.getDataRange().getValues();
  data.shift();

  let income = 0, expense = 0, workExpense = 0, privateExpense = 0;
  const history = [];

  data.forEach((row, index) => {
    const dateVal = row[0];
    const type    = String(row[1] || "");
    const item    = String(row[2] || "");
    const amount  = Number(row[3]) || 0;
    const cat     = String(row[4] || "");
    const method  = String(row[5] || "");

    if (!dateVal || isNaN(new Date(dateVal).getTime())) return;

    const d        = new Date(dateVal);
    const rowMonth = Utilities.formatDate(d, "JST", "yyyy-MM");
    if (rowMonth !== month) return;

    const isIncome = type.includes("収入") || type.toUpperCase().includes("INCOME");
    const isForTwo = type.includes("For Two") || type.includes("仕事") || type.includes("経費");

    if (isIncome) {
      income += amount;
    } else if (isForTwo) {
      expense     += amount;
      workExpense += amount;
    } else {
      expense        += amount;
      privateExpense += amount;
    }

    history.push({
      row:      index + 2,
      date:     Utilities.formatDate(d, "JST", "MM/dd"),
      fullDate: Utilities.formatDate(d, "JST", "yyyy-MM-dd"),
      weekday:  ["日","月","火","水","木","金","土"][d.getDay()],
      type:     type,
      cat:      cat,
      method:   method,
      item:     item,
      amount:   amount
    });
  });

  return {
    monthly:  { income, expense, balance: income - expense, workExpense, privateExpense },
    history:  history,
    settings: getSettings(),
    currentMonth: month
  };
}

// ─── 入力・更新・削除 ────────────────────────────────────────
function processForm(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  if (!sheet) throw new Error("データシートが見つかりません。");
  const pureDate = Utilities.formatDate(new Date(form.date), "JST", "yyyy/MM/dd");
  sheet.appendRow([pureDate, form.type, form.item, Number(form.amount), form.category, form.method]);
  _sortSheet(sheet);
  return true;
}

function updateRow(row, form) {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const pureDate = Utilities.formatDate(new Date(form.date), "JST", "yyyy/MM/dd");
  sheet.getRange(row, 1, 1, 6).setValues([[pureDate, form.type, form.item, Number(form.amount), form.category, form.method]]);
  _sortSheet(sheet);
  return true;
}

function deleteRow(row) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME).deleteRow(row);
  return true;
}

function _sortSheet(sheet) {
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 6).sort({ column: 1, ascending: true });
}

// ─── 設定読み込み ────────────────────────────────────────────
// 設定シートのA列キー名には依存しない。
// B列の値を全行読んで、キー名に含まれるキーワードで6種類に振り分ける。
// どうしてもマッチしない行はスキップしてデフォルト値で補完する。
function getSettings() {
  const defaults = _getDefaultSettings();
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!setSheet) return defaults;

  const rows = setSheet.getDataRange().getValues();

  const settings = {
    HOME_cat:      null,
    HOME_method:   null,
    WORK_cat:      null,
    WORK_method:   null,
    INCOME_cat:    null,
    INCOME_method: null
  };

  rows.forEach(function(row) {
    var rawKey = String(row[0] || "").trim();
    var rawVal = String(row[1] || "").trim();
    if (!rawKey || !rawVal) return;

    // キー名を小文字・記号除去して比較
    var k = rawKey
      .toLowerCase()
      .replace(/[（）()「」【】\[\]\s_　]/g, "");

    // 値をパース（JSON配列 or カンマ区切り）
    var vals = _parseValues(rawVal);
    if (!vals.length) return;

    // 収入系 — 最優先
    if (k.indexOf("収入") >= 0) {
      if (k.indexOf("cat") >= 0 || k.indexOf("カテ") >= 0) {
        settings.INCOME_cat = vals;
      } else if (k.indexOf("method") >= 0 || k.indexOf("支払") >= 0) {
        settings.INCOME_method = vals;
      } else {
        // 収入だがcatもmethodもない → methodsキーかもしれない
        if (k.indexOf("methods") >= 0) settings.INCOME_method = vals;
      }
      return;
    }

    // 仕事・ForTwo系
    if (k.indexOf("仕事") >= 0 || k.indexOf("fortwo") >= 0 || k.indexOf("frotwo") >= 0) {
      if (k.indexOf("cat") >= 0 || k.indexOf("カテ") >= 0) {
        settings.WORK_cat = vals;
      } else if (k.indexOf("method") >= 0 || k.indexOf("支払") >= 0) {
        settings.WORK_method = vals;
      }
      return;
    }

    // プライベート・家計系
    if (k.indexOf("プライベート") >= 0 || k.indexOf("家計") >= 0 || k.indexOf("private") >= 0) {
      if (k.indexOf("cat") >= 0 || k.indexOf("カテ") >= 0) {
        settings.HOME_cat = vals;
      } else if (k.indexOf("method") >= 0 || k.indexOf("支払") >= 0) {
        settings.HOME_method = vals;
      }
      return;
    }

    // 汎用「支出_methods」など — 支払方法として HOME/WORK 両方に
    if (k.indexOf("支出") >= 0 && (k.indexOf("method") >= 0 || k.indexOf("支払") >= 0)) {
      if (!settings.HOME_method) settings.HOME_method = vals;
      if (!settings.WORK_method) settings.WORK_method = vals;
      return;
    }
  });

  // null のものはデフォルトで補完
  Object.keys(settings).forEach(function(k) {
    if (!settings[k] || !settings[k].length) settings[k] = defaults[k];
  });

  return settings;
}

function _parseValues(raw) {
  if (!raw) return [];
  var trimmed = raw.trim();
  // JSON配列形式
  if (trimmed.charAt(0) === "[") {
    try {
      var parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) {
        return parsed.map(function(v){ return String(v).trim(); }).filter(Boolean);
      }
    } catch(e) {}
  }
  // カンマ区切り（前後の引用符・スペースを除去）
  return trimmed.split(",").map(function(s) {
    return s.trim().replace(/^["'\s]+|["'\s]+$/g, "");
  }).filter(Boolean);
}

// ─── 設定保存 ────────────────────────────────────────────────
// アプリ上で編集した設定をスプシに書き戻す。
// 既存行のA列キー名は変えず、B列の値だけ上書きする。
// マッチしない内部キーは末尾に追記。
function saveSettings(settingsObj) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SETTINGS_SHEET_NAME);

  var rows   = sheet.getDataRange().getValues();
  var used   = {};

  // 既存行を更新
  rows.forEach(function(row, i) {
    var rawKey = String(row[0] || "").trim();
    if (!rawKey) return;
    var k = rawKey.toLowerCase().replace(/[（）()「」【】\[\]\s_　]/g, "");
    var internalKey = null;

    if      (k.indexOf("収入") >= 0 && (k.indexOf("cat") >= 0 || k.indexOf("カテ") >= 0)) internalKey = "INCOME_cat";
    else if (k.indexOf("収入") >= 0 && (k.indexOf("method") >= 0 || k.indexOf("支払") >= 0)) internalKey = "INCOME_method";
    else if ((k.indexOf("仕事") >= 0 || k.indexOf("fortwo") >= 0) && (k.indexOf("cat") >= 0 || k.indexOf("カテ") >= 0)) internalKey = "WORK_cat";
    else if ((k.indexOf("仕事") >= 0 || k.indexOf("fortwo") >= 0) && (k.indexOf("method") >= 0 || k.indexOf("支払") >= 0)) internalKey = "WORK_method";
    else if ((k.indexOf("プライベート") >= 0 || k.indexOf("家計") >= 0) && (k.indexOf("cat") >= 0 || k.indexOf("カテ") >= 0)) internalKey = "HOME_cat";
    else if ((k.indexOf("プライベート") >= 0 || k.indexOf("家計") >= 0) && (k.indexOf("method") >= 0 || k.indexOf("支払") >= 0)) internalKey = "HOME_method";
    else if (k.indexOf("支出") >= 0 && (k.indexOf("method") >= 0) && !used["HOME_method"]) { internalKey = "HOME_method"; }

    if (internalKey && settingsObj[internalKey] && !used[internalKey]) {
      rows[i][1] = settingsObj[internalKey].join(",");
      used[internalKey] = true;
    }
  });

  // 未更新の内部キーは末尾に追記
  var allKeys = ["HOME_cat","HOME_method","WORK_cat","WORK_method","INCOME_cat","INCOME_method"];
  allKeys.forEach(function(ik) {
    if (!used[ik] && settingsObj[ik]) {
      rows.push([ik, settingsObj[ik].join(",")]);
    }
  });

  sheet.clearContents();
  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, 2).setValues(
      rows.map(function(r){ return [String(r[0]||""), String(r[1]||"")]; })
    );
  }
  sheet.autoResizeColumns(1, 2);
  return true;
}

// ─── デフォルト設定値 ────────────────────────────────────────
function _getDefaultSettings() {
  return {
    HOME_cat:      ["食費","外食費","日用品","固定費","車","衣服","薬","交際費","献金","大きな出費","ガソリン","こども","美容"],
    HOME_method:   ["現金","楽天カード","楽天ETC","セゾンカード","ゆさ楽天カード","PayPay","銀行引落"],
    WORK_cat:      ["旅費交通費","雑費","消耗品費","車両費","会議費","通信費","支払手数料","その他経費"],
    WORK_method:   ["仕事用カード","楽天カード","現金","楽天ETC","銀行振込"],
    INCOME_cat:    ["メモリーツリー","株式会社クオリティオブライフ","株式会社FreeLabo","For Two"],
    INCOME_method: ["振込","楽天銀行","現金"]
  };
}

// ─── デバッグ用（Apps Scriptエディタから手動実行してログ確認） ──
function debugSettings() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) { Logger.log("設定シートなし"); return; }
  var rows  = sheet.getDataRange().getValues();
  Logger.log("=== 設定シート全行 ===");
  rows.forEach(function(row, i) {
    Logger.log("行" + (i+1) + ": KEY=[" + row[0] + "] VAL=[" + row[1] + "]");
  });
  Logger.log("=== パース結果 ===");
  Logger.log(JSON.stringify(getSettings()));
}
