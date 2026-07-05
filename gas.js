const HEADERS = ["id", "roomId", "date", "item", "peopleQty", "inQty", "usedQty", "balanceQty", "wasteQty", "note", "updatedAt"];

function doGet(e) {
  const params = e.parameter || {};
  const sheet = getPlainSheet(params.sheet);
  const rows = readRowsDynamic(sheet);
  return json({ ok: true, rows });
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || "{}");

  if (payload.action === "appendReport") {
    const sheet = getSheetWithHeaders(payload.sheet, payload.headers || []);
    appendReportRows(sheet, payload.headers || [], payload.rows || []);
    return json({ ok: true, mode: "appendReport", count: (payload.rows || []).length });
  }

  if (payload.action === "upsertDaily") {
    const sheet = getDailySheet(payload.sheet);
    const rowNumber = upsertDailyByDate(sheet, payload.fields || {});
    return json({ ok: true, mode: "upsertDaily", rowNumber: rowNumber });
  }

  const sheet = getSheet(payload.sheet);

  if (payload.action === "overwrite") {
    writeAll(sheet, payload.rows || []);
    return json({ ok: true, mode: "overwrite", count: (payload.rows || []).length });
  }

  if (payload.action === "delete") {
    deleteById(sheet, payload.id);
    return json({ ok: true, mode: "delete" });
  }

  upsertRow(sheet, payload.row);
  return json({ ok: true, mode: "upsert", row: payload.row });
}

function getDailySheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = name || "Sheet A";
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  ensureDailyHeader(sheet, ["日期", "備註"]);
  return sheet;
}

function ensureDailyHeader(sheet, requiredHeaders) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (headers.join("") === "") {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }
  const existing = getHeaderMap(sheet);
  requiredHeaders.forEach((header) => {
    if (!existing[header]) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
    }
  });
}

function upsertDailyByDate(sheet, fields) {
  if (!fields["日期"]) throw new Error("日期欄位是必要欄位");
  fields["日期"] = normalizeDateKey(fields["日期"]);
  ensureDailyHeader(sheet, Object.keys(fields));
  const headerMap = getHeaderMap(sheet);
  const dateColumn = headerMap["日期"];
  const matchedRows = findRowsByDate(sheet, dateColumn, fields["日期"]);
  const rowNumber = matchedRows.length
    ? mergeDuplicateDateRows(sheet, matchedRows, headerMap)
    : sheet.getLastRow() + 1;

  Object.entries(fields).forEach(([header, value]) => {
    const column = headerMap[header];
    if (!column) return;
    const cell = sheet.getRange(rowNumber, column);
    if (header === "日期") cell.setNumberFormat("@");
    cell.setValue(value == null ? "" : value);
  });
  return rowNumber;
}

// 建立欄位名稱 Mapping，同時對 Key 做去空白處理以避免前後空格導致的比對失敗
function getHeaderMap(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    if (header) map[String(header).trim()] = index + 1;
  });
  return map;
}

function findRowsByDate(sheet, dateColumn, targetDate) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = sheet.getRange(2, dateColumn, lastRow - 1, 1).getValues();
  const target = normalizeDateKey(targetDate);
  return values.reduce((matched, row, index) => {
    if (normalizeDateKey(row[0]) === target) matched.push(index + 2);
    return matched;
  }, []);
}

function mergeDuplicateDateRows(sheet, rowNumbers, headerMap) {
  const keepRow = Math.min.apply(null, rowNumbers);
  const duplicateRows = rowNumbers.filter((row) => row !== keepRow).sort((a, b) => b - a);
  if (!duplicateRows.length) return keepRow;

  const lastColumn = sheet.getLastColumn();
  const dateColumn = headerMap["日期"];
  const mergedValues = sheet.getRange(keepRow, 1, 1, lastColumn).getValues()[0];

  duplicateRows.slice().reverse().forEach((rowNumber) => {
    const values = sheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
    values.forEach((value, index) => {
      const column = index + 1;
      if (column === dateColumn || isBlank(value)) return;
      mergedValues[index] = value;
    });
  });

  mergedValues[dateColumn - 1] = normalizeDateKey(mergedValues[dateColumn - 1]);
  sheet.getRange(keepRow, 1, 1, lastColumn).setValues([mergedValues]);
  duplicateRows.forEach((rowNumber) => sheet.deleteRow(rowNumber));
  return keepRow;
}

function normalizeDateKey(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const text = String(value || "").trim();
  const match = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (match) return formatDateKey(match[1], match[2], match[3]);
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return text.replace(/\//g, "-").slice(0, 10);
}

function formatDateKey(year, month, day) {
  return [
    String(year).padStart(4, "0"),
    String(month).padStart(2, "0"),
    String(day).padStart(2, "0")
  ].join("-");
}

function isBlank(value) {
  return value === "" || value == null;
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = name || "Sheet A";
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  ensureHeader(sheet);
  return sheet;
}

function getPlainSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = name || "Sheet A";
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function getSheetWithHeaders(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = name || "Summary";
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const safeHeaders = headers && headers.length ? headers : ["recordedAt", "month", "name", "value"];
  const range = sheet.getRange(1, 1, 1, safeHeaders.length);
  const values = range.getValues()[0];
  if (values.join("") === "") range.setValues([safeHeaders]);
  else if (!hasHeaderRow(sheet, safeHeaders)) {
    sheet.appendRow(safeHeaders);
  }
  return sheet;
}

function appendReportRows(sheet, headers, rows) {
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length)
    .setValues(rows.map((row) => headers.map((key) => row[key] == null ? "" : row[key])));
}

function hasHeaderRow(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (!lastRow) return false;
  const width = headers.length;
  return sheet.getRange(1, 1, lastRow, width).getValues()
    .some((row) => row.join("|") === headers.join("|"));
}

function ensureHeader(sheet) {
  const range = sheet.getRange(1, 1, 1, HEADERS.length);
  const values = range.getValues()[0];
  if (values.join("") !== HEADERS.join("")) range.setValues([HEADERS]);
}

function readRows(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues().map((row) => {
    const record = {};
    HEADERS.forEach((key, index) => record[key] = row[index]);
    return record;
  });
}

// 重構重點：將 getValues() 改為 getDisplayValues()，取得試算表中的字串而非原始 Date 物件。
// 這可以完美防止經由 JSON.stringify 轉換為 UTC-0 時所造成的 1 天日期差 Bug。
function readRowsDynamic(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn < 1) return [];
  
  // 取得經去空白處理後的 Headers 陣列
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => String(h || "").trim());
  
  // 使用 getDisplayValues() 保持與畫面顯示的一致性字串
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  
  return values.map((row) => {
    const record = {};
    headers.forEach((key, index) => {
      if (key) record[key] = row[index];
    });
    return record;
  }).filter((record) => Object.values(record).some((value) => !isBlank(value)));
}

function upsertRow(sheet, row) {
  if (!row || !row.id) throw new Error("row.id is required");
  const rows = readRows(sheet);
  const index = rows.findIndex((record) => String(record.id).trim() === String(row.id).trim());
  const values = [HEADERS.map((key) => row[key] || "")];
  if (index >= 0) sheet.getRange(index + 2, 1, 1, HEADERS.length).setValues(values);
  else sheet.appendRow(values[0]);
}

function writeAll(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (!rows.length) return;
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows.map((row) => HEADERS.map((key) => row[key] || "")));
}

function deleteById(sheet, id) {
  const rows = readRows(sheet);
  const index = rows.findIndex((record) => String(record.id).trim() === String(id).trim());
  if (index >= 0) sheet.deleteRow(index + 2);
}

function json(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}