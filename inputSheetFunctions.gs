// ============================================================
// Expense logger. Edits to the "Input" sheet get appended to a
// monthly log sheet. A time-driven trigger runs createMonthlySheet
// on the 1st of each month (midnight–1am) to pre-create that
// month's sheet from the prior month's template.
// ============================================================

// --- Helpers ---------------------------------------------------

// Today at noon in the spreadsheet's timezone. Noon keeps day-math
// safely away from DST transitions and date boundaries.
function getSpreadsheetTodayAtNoon() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  const dateString = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const offsetString = Utilities.formatDate(now, tz, "XXX"); // e.g. "-05:00"
  return new Date(`${dateString}T12:00:00${offsetString}`);
}

// Returns { currentKey, priorKey } like "October 2026" / "September 2026",
// computed in the spreadsheet's timezone.
function getMonthKeys() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const today = getSpreadsheetTodayAtNoon();

  const year = parseInt(Utilities.formatDate(today, tz, "yyyy"), 10);
  const month = parseInt(Utilities.formatDate(today, tz, "MM"), 10); // 1–12

  const priorYear = month === 1 ? year - 1 : year;
  const priorMonthIdx = month === 1 ? 11 : month - 2; // 0-indexed
  const monthNames = ['January','February','March','April','May','June',
                      'July','August','September','October','November','December'];

  return {
    currentKey: Utilities.formatDate(today, tz, "MMMM yyyy"),
    priorKey: `${monthNames[priorMonthIdx]} ${priorYear}`
  };
}

// Hide every sheet except "Input" and the given sheet to keep.
// Apps Script won't let you hide the active sheet, so activate
// the keeper first.
function hideOldMonthSheets(keepSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(keepSheet);

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name !== "Input" && name !== keepSheet.getName()) {
      sheet.hideSheet();
    }
  });
}

// Ensures a sheet exists for the current month. Copies the prior
// month's sheet as a template if available; otherwise creates a
// bare sheet with just the header row.
function ensureCurrentMonthSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { currentKey, priorKey } = getMonthKeys();

  const existing = ss.getSheetByName(currentKey);
  if (existing) return existing;

  const template = ss.getSheetByName(priorKey);
  if (template) {
    const sheet = template.copyTo(ss);
    sheet.setName(currentKey);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(ss.getSheets().length);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 1 && lastCol > 0) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }
    
    hideOldMonthSheets(sheet);
    return sheet;
  }

  const sheet = ss.insertSheet(currentKey);
  sheet.appendRow(['Date', 'Amount', 'Category', 'Card Type']);
  hideOldMonthSheets(sheet);
  return sheet;
}

// --- Entry points ----------------------------------------------

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "Input") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row !== 2 || col < 1 || col > 3) return;

  const amount = sheet.getRange("A2").getValue();
  const category = sheet.getRange("B2").getValue();
  const cardType = sheet.getRange("C2").getValue();
  if (!amount || !category || !cardType) return;

  cutAndPaste();
}

function cutAndPaste() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");

  const amount = inputSheet.getRange("A2").getValue();
  const category = inputSheet.getRange("B2").getValue();
  const cardType = inputSheet.getRange("C2").getValue();
  if (!amount || !category || !cardType) return;

  const monthlySheet = ensureCurrentMonthSheet();
  const today = getSpreadsheetTodayAtNoon();

  monthlySheet.appendRow([today, amount, category, cardType]);

  // Store date as a real Date; format the new cell as date-only.
  const lastRow = monthlySheet.getLastRow();
  monthlySheet.getRange(lastRow, 1).setNumberFormat('MM/dd/yyyy');

  inputSheet.getRange("A2:C2").clearContent();
}

// Time-driven trigger: 1st of month, midnight–1am.
function createMonthlySheet() {
  ensureCurrentMonthSheet();
}