// ============================================================
//  PROJECT CALENDAR — Code.gs
//  Apps Script Backend
// ============================================================

const SHEET_NAME = "Events";
const COLUMNS   = ["id","title","date","endDate","time","endTime","color","description","category"];

// ── Entry Point ──────────────────────────────────────────────
function doGet() {
  // Coba nama file dengan dan tanpa .html
  let output;
  try {
    output = HtmlService.createHtmlOutputFromFile("index");
  } catch(e1) {
    try {
      output = HtmlService.createHtmlOutputFromFile("index.html");
    } catch(e2) {
      // Tampilkan daftar file yang tersedia untuk debug
      return HtmlService.createHtmlOutput(
        "<h3>Debug: Nama file yang tersedia di project ini:</h3>" +
        "<p>Error 1 (index): " + e1.message + "</p>" +
        "<p>Error 2 (index.html): " + e2.message + "</p>" +
        "<p>Silakan rename file HTML kamu sesuai nama yang dikenali.</p>"
      );
    }
  }
  return output
    .setTitle("Project Calendar")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Sheet Helper ─────────────────────────────────────────────
function getOrCreateSheet_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, COLUMNS.length)
         .setBackground("#1A1D2E")
         .setFontColor("#FFFFFF")
         .setFontWeight("bold");
  }
  return sheet;
}

// ── CRUD Operations ──────────────────────────────────────────

/** Fetch all events — returns array of plain objects */
function getEvents() {
  const sheet = getOrCreateSheet_();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

/** Add a new event */
function addEvent(event) {
  const sheet = getOrCreateSheet_();
  const id    = Utilities.getUuid();
  sheet.appendRow([
    id,
    event.title       || "",
    event.date        || "",
    event.endDate     || event.date || "",
    event.time        || "",
    event.endTime     || "",
    event.color       || "#4F6EF7",
    event.description || "",
    event.category    || "General"
  ]);
  return { success: true, id };
}

/** Update an existing event by id */
function updateEvent(event) {
  const sheet  = getOrCreateSheet_();
  const data   = sheet.getDataRange().getValues();
  const colMap = {};
  data[0].forEach((h, i) => colMap[h] = i);

  for (let r = 1; r < data.length; r++) {
    if (data[r][colMap["id"]] === event.id) {
      const row = r + 1;
      sheet.getRange(row, colMap["title"]       + 1).setValue(event.title       || "");
      sheet.getRange(row, colMap["date"]        + 1).setValue(event.date        || "");
      sheet.getRange(row, colMap["endDate"]     + 1).setValue(event.endDate     || event.date || "");
      sheet.getRange(row, colMap["time"]        + 1).setValue(event.time        || "");
      sheet.getRange(row, colMap["endTime"]     + 1).setValue(event.endTime     || "");
      sheet.getRange(row, colMap["color"]       + 1).setValue(event.color       || "#4F6EF7");
      sheet.getRange(row, colMap["description"] + 1).setValue(event.description || "");
      sheet.getRange(row, colMap["category"]    + 1).setValue(event.category    || "General");
      return { success: true };
    }
  }
  return { success: false, error: "Event not found" };
}

/** Delete an event by id */
function deleteEvent(eventId) {
  const sheet = getOrCreateSheet_();
  const data  = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf("id");

  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol] === eventId) {
      sheet.deleteRow(r + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Event not found" };
}