/**
 * 1. ウェブアプリを表示するための窓口
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('For Two 家計簿システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const DATA_SHEET_NAME     = "シート1";
const SCHEDULE_SHEET_NAME = "スケジュール";

/**
 * 2. データ取得（列順: A日付, Bタイプ, C内容, D金額, Eカテゴリ, F方法）
 */
function getReportData(month) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) throw new Error("シート「シート1」が見つかりません。");

  const data = sheet.getDataRange().getValues();
  data.shift();

  let income = 0, expense = 0, workExpense = 0, privateExpense = 0;
  const history = [];

  data.forEach((row, index) => {
    let dateVal = row[0];
    let type    = String(row[1] || "");
    let item    = String(row[2] || "");
    let amount  = Number(row[3]) || 0;
    let cat     = String(row[4] || "");
    let method  = String(row[5] || "");

    if (!dateVal || isNaN(new Date(dateVal).getTime())) return;

    const d        = new Date(dateVal);
    const rowMonth = Utilities.formatDate(d, "JST", "yyyy-MM");

    if (rowMonth === month) {
      if (type.includes("収入") || type.toUpperCase().includes("INCOME")) {
        income += amount;
      } else if (type.includes("For Two")) {
        expense += amount;
        workExpense += amount;
      } else {
        expense += amount;
        privateExpense += amount;
      }

      history.push({
        row:      index + 2,
        date:     Utilities.formatDate(d, "JST", "MM/dd"),
        fullDate: Utilities.formatDate(d, "JST", "yyyy-MM-dd"),
        type, cat, method, item, amount
      });
    }
  });

  return {
    monthly:  { income, expense, balance: income - expense, workExpense, privateExpense },
    history,
    settings: getSettings(),
    currentMonth: month
  };
}

/**
 * 3. 保存・更新・削除
 */
function processForm(form) {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const pureDate = Utilities.formatDate(new Date(form.date), "JST", "yyyy/MM/dd");
  sheet.appendRow([pureDate, form.type, form.item, Number(form.amount), form.category, form.method]);
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).sort({ column: 1, ascending: true });
  return true;
}

function updateRow(row, form) {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const pureDate = Utilities.formatDate(new Date(form.date), "JST", "yyyy/MM/dd");
  sheet.getRange(row, 1, 1, 6).setValues([[pureDate, form.type, form.item, Number(form.amount), form.category, form.method]]);
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).sort({ column: 1, ascending: true });
  return true;
}

function deleteRow(row) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME).deleteRow(row);
  return true;
}

/**
 * 4. 設定（タグ）取得
 *    設定シートの形式：A列＝キー名、B列＝カンマ区切りの値リスト
 *    キーが重複している場合は最初の行のみ使用する
 */
function getSettings() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = ss.getSheetByName("設定") || ss.insertSheet("設定");
  const rows     = setSheet.getDataRange().getValues();
  const settings = {};

  rows.forEach(row => {
    const key = String(row[0] || "").trim();
    const val = String(row[1] || "").trim();
    if (!key) return;
    if (settings[key] !== undefined) return; // 重複行は無視
    settings[key] = val.split(",").map(v => v.trim()).filter(Boolean);
  });

  return settings;
}

/**
 * 5. タグ追加
 *    対象キーの最初の行のB列にカンマ区切りで追記する
 */
function saveTag(colName, tagValue) {
  if (!colName || !tagValue || !tagValue.trim()) throw new Error("タグ名が空です");

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = ss.getSheetByName("設定") || ss.insertSheet("設定");
  const trimmed  = tagValue.trim();
  const rows     = setSheet.getDataRange().getValues();

  let targetRow = -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === colName) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    const newRow = rows.length + 1;
    setSheet.getRange(newRow, 1).setValue(colName);
    setSheet.getRange(newRow, 2).setValue(trimmed);
    return getSettings();
  }

  const existing = String(rows[targetRow - 1][1] || "").trim();
  const list     = existing.split(",").map(v => v.trim()).filter(Boolean);
  if (list.includes(trimmed)) return getSettings();

  list.push(trimmed);
  setSheet.getRange(targetRow, 2).setValue(list.join(","));
  return getSettings();
}

/**
 * 6. タグ削除
 *    対象キーの最初の行のB列から指定タグを取り除いて書き戻す
 */
function deleteTag(colName, tagValue) {
  if (!colName || !tagValue) throw new Error("引数が不正です");

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = ss.getSheetByName("設定");
  if (!setSheet) return getSettings();

  const rows = setSheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === colName) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) return getSettings();

  const existing = String(rows[targetRow - 1][1] || "").trim();
  const newList  = existing.split(",").map(v => v.trim()).filter(v => v && v !== tagValue);
  setSheet.getRange(targetRow, 2).setValue(newList.join(","));
  return getSettings();
}

/**
 * 7. デバッグ用：設定シートの内容をログ出力
 */
function debugSettings() {
  const settings = getSettings();
  Logger.log(JSON.stringify(settings, null, 2));
}
