const SPREADSHEET_ID = "19bcl9KOZa2trDGTW3pG9DTcIghh5Y4AYVIEcxJ2bz3s";
const SHEET_NAME = "意願登記";

const SLOT_LABELS = {
  "週一-am": "週一上午班",
  "週一-pm": "週一下午班",
  "週二-am": "週二上午班",
  "週二-pm": "週二下午班",
  "週三-am": "週三上午班",
  "週三-pm": "週三下午班",
  "週四-am": "週四上午班",
  "週四-pm": "週四下午班",
  "週五-am": "週五上午班",
  "週五-pm": "週五下午班",
  "週六-special": "週六輪值班"
};

function doGet(e) {
  const params = e.parameter || {};
  const callback = sanitizeCallback(params.callback || "callback");

  try {
    const action = params.action || "list";
    let payload;

    if (action === "save") {
      payload = saveRecord(params);
    } else if (action === "clear") {
      payload = clearRecords();
    } else {
      payload = listRecords();
    }

    return jsonp(callback, Object.assign({ ok: true }, payload));
  } catch (error) {
    return jsonp(callback, { ok: false, error: error.message });
  }
}

function saveRecord(params) {
  const name = String(params.name || "").trim();
  const max = Number(params.max || 1);
  const preferences = JSON.parse(params.preferences || "[]");

  if (!name) throw new Error("缺少姓名。");
  if (!Array.isArray(preferences) || preferences.length === 0) throw new Error("缺少可排時段。");

  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[0]).trim() === name);
  const row = [
    name,
    max,
    preferences.join(","),
    preferences.map(id => SLOT_LABELS[id] || id).join("、"),
    preferences.includes("週六-special") ? "是" : "否",
    new Date(),
    "GitHub Pages",
    ""
  ];

  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return { record: normalizeRecord(row) };
}

function listRecords() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues().slice(1);
  const records = values
    .filter(row => String(row[0] || "").trim())
    .map(normalizeRecord);

  return { records };
}

function clearRecords() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  return { records: [] };
}

function normalizeRecord(row) {
  return {
    name: String(row[0] || "").trim(),
    max: Number(row[1] || 1),
    preferences: String(row[2] || "")
      .split(",")
      .map(value => value.trim())
      .filter(Boolean),
    preferenceText: String(row[3] || ""),
    saturday: String(row[4] || ""),
    updatedAt: row[5] instanceof Date ? row[5].toISOString() : String(row[5] || "")
  };
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`找不到工作表：${SHEET_NAME}`);
  return sheet;
}

function jsonp(callback, data) {
  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify(data)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function sanitizeCallback(callback) {
  return String(callback).replace(/[^\w.$]/g, "") || "callback";
}
