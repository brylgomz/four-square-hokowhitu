/**
 * Four Square Hokowhitu — hoko_kitchen backend
 *
 * One-time setup:
 *   1. Go to script.google.com, create a new project, paste this whole file in as Code.gs.
 *   2. Run the `setup` function once (Run > setup). Approve the permissions prompt.
 *      This creates (or finds) the "hoko_kitchen" spreadsheet inside the target
 *      Drive folder, with 4 tabs: Food_List, Start_day, Remaining_Foods, Leftover_Foods.
 *   3. Deploy > New deployment > type "Web app". Execute as "Me", access "Anyone".
 *   4. Copy the /exec URL it gives you and paste it into APPS_SCRIPT_URL at the
 *      top of the <script> block in four_square_hokowhitu.html.
 *   5. Open the /exec URL directly in a browser to confirm — it should return
 *      {"ok":true,"message":"hoko_kitchen backend is running"}.
 *      (Do not rely on automated fetch checks here — Google's bot detection
 *      makes /exec URLs unreliable to verify from scripts; a real browser load
 *      is the only trustworthy check.)
 */

var DRIVE_FOLDER_ID = "1s1UxVhOK8pZvUmAwdMQY9Q8q3VEAKPhA"; // Foursquare_Hokowhitu
var SPREADSHEET_NAME = "hoko_kitchen";
var DATE_FORMAT = "dd/MM/yyyy"; // must match the date format the app sends
var REPORT_EMAIL = "ryll_bullets@yahoo.com"; // recipient for the end-of-day leftover report

var FOOD_LIST = [
  "Bacon & Egg Pies","Cheeseburger Pies","Chicken Cranberry & Brie Pies",
  "Creamy Chicken Pies","Steak & Mushrooms Pies","Mince & Cheese Pies",
  "BS Mince & Cheese Pies","Steak & Cheese Pies","BS Steak & Cheese Pies",
  "BS Quiche Pies","BS Steak & Onion Pies","Sausage Roll Pies",
  "Lamb & Kumara Pies","Potato Top Pies","Pepper Steak Pies",
  "Battered Hot Dog","Cordon Bleu","Lasagne Toppa",
  "Beef & Cabbage (Springroll)","Butter Chicken (Springroll)",
  "Kransky","Kebab","DimSim","Mac&Cheese","Satay Kebab","Chicken Popcorn",
  "Chipotle Chicken Burrito","Nacho Beef & Cheese Burrito"
];

var DATA_TABS = ["Start_day", "Remaining_Foods", "Leftover_Foods"];

var ACTION_TO_TAB = {
  start: "Start_day",
  remaining: "Remaining_Foods",
  leftover: "Leftover_Foods"
};

// ---------- one-time setup ----------

function setup() {
  var ss = getOrCreateSpreadsheet();
  setupFoodListTab(ss);
  DATA_TABS.forEach(function(tabName) {
    setupDataTab(ss, tabName);
  });
  Logger.log("Setup complete: " + ss.getUrl());
}

function getOrCreateSpreadsheet() {
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var files = folder.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  var ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  var defaultSheet = ss.getSheets()[0];
  defaultSheet.setName("Food_List");
  return ss;
}

function setupFoodListTab(ss) {
  var sheet = ss.getSheetByName("Food_List") || ss.insertSheet("Food_List");
  sheet.clear();
  sheet.getRange(1, 1).setValue("Food Item").setFontWeight("bold");
  for (var i = 0; i < FOOD_LIST.length; i++) {
    sheet.getRange(i + 2, 1).setValue(FOOD_LIST[i]);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumn(1);
}

function setupDataTab(ss, tabName) {
  var sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sheet.clear();
  sheet.getRange(1, 1).setValue("Food Item").setFontWeight("bold");
  for (var i = 0; i < FOOD_LIST.length; i++) {
    sheet.getRange(i + 2, 1).setValue(FOOD_LIST[i]);
  }

  var dateStrings = remainingDatesOf2026();
  for (var c = 0; c < dateStrings.length; c++) {
    var col = c + 2;
    var cell = sheet.getRange(1, col);
    // setNumberFormat BEFORE setValue: otherwise Sheets auto-converts a
    // "2026-09-10"-shaped string into a real Date, which then fails a
    // strict string match in getOrCreateDateColumn below and causes a
    // brand new duplicate column to be created on every sync instead of
    // reusing this one.
    cell.setNumberFormat("@").setValue(dateStrings[c]).setFontWeight("bold");
  }

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

function remainingDatesOf2026() {
  var tz = Session.getScriptTimeZone();
  var start = new Date(2026, 8, 6); // fixed tracking start date: 6 Sep 2026
  var end = new Date(2026, 11, 31);
  var dates = [];
  var d = new Date(start);
  while (d <= end) {
    dates.push(Utilities.formatDate(d, tz, DATE_FORMAT));
    d.setDate(d.getDate() + 1);
  }
  return dates;
}

// Non-destructive: inserts any date from remainingDatesOf2026() that isn't
// already a header column yet (e.g. earlier dates than what the sheet
// currently starts at, like backfilling from 06/09/2026). New columns are
// inserted in chronological order at the front — existing columns and all
// their data just shift right, nothing is cleared or removed.
function insertMissingEarlyDates() {
  var ss = getOrCreateSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var allDates = remainingDatesOf2026();
  DATA_TABS.forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) { return; }
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var existingIso = {};
    for (var c = 1; c < headerRow.length; c++) {
      var d = parseHeaderDate(headerRow[c]);
      if (d) { existingIso[Utilities.formatDate(d, tz, "yyyy-MM-dd")] = true; }
    }
    var missing = allDates.filter(function(dateStr) {
      var d = parseHeaderDate(dateStr);
      return !existingIso[Utilities.formatDate(d, tz, "yyyy-MM-dd")];
    });
    for (var m = missing.length - 1; m >= 0; m--) {
      sheet.insertColumnBefore(2);
      sheet.getRange(1, 2).setNumberFormat("@").setValue(missing[m]).setFontWeight("bold");
    }
  });
  Logger.log("Inserted any missing date columns (backfilled from 06/09/2026).");
}

// One-off, non-destructive migration: relabels existing date-header cells
// in place (old "yyyy-MM-dd" strings, or any cell Sheets auto-converted
// to a real Date) to DATE_FORMAT. Never clears or removes rows/columns —
// only rewrites the header text, so all existing food rows and saved
// quantities are left exactly where they are.
function reformatDateHeaders() {
  var ss = getOrCreateSpreadsheet();
  var tz = Session.getScriptTimeZone();
  DATA_TABS.forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) { return; }
    var lastCol = sheet.getLastColumn();
    if (lastCol < 2) { return; }
    var headerRange = sheet.getRange(1, 2, 1, lastCol - 1);
    var headerRow = headerRange.getValues()[0];
    for (var c = 0; c < headerRow.length; c++) {
      var d = parseHeaderDate(headerRow[c]);
      if (!d) { continue; }
      var newStr = Utilities.formatDate(d, tz, DATE_FORMAT);
      sheet.getRange(1, c + 2).setNumberFormat("@").setValue(newStr).setFontWeight("bold");
    }
  });
  Logger.log("Date headers reformatted to " + DATE_FORMAT + ".");
}

// Recognizes a header cell as a date regardless of which format it was
// written in (old "yyyy-MM-dd", new DATE_FORMAT, or a real Date object)
// and returns it as a JS Date, or null if it isn't a date at all.
function parseHeaderDate(cellValue) {
  if (cellValue instanceof Date) { return cellValue; }
  if (typeof cellValue !== "string") { return null; }
  var iso = cellValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) { return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])); }
  var dmy = cellValue.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (dmy) { return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1])); }
  return null;
}

// ---------- runtime sync from the app ----------

function doGet(e) {
  return jsonResponse({ ok: true, message: "hoko_kitchen backend is running" });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === "addFoodItem") {
      return handleAddFoodItem(body);
    }

    if (body.action === "soldOut") {
      return handleSoldOut(body);
    }

    var tabName = ACTION_TO_TAB[body.action];
    if (!tabName) {
      return jsonResponse({ ok: false, error: "Unknown action: " + body.action });
    }

    var ss = getOrCreateSpreadsheet();
    var sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
    var dateCol = getOrCreateDateColumn(sheet, body.date);
    var items = body.items || [];

    // Plain items overwrite the cell for the day; TopUp items are extra
    // quantity cooked on top of the base dish, so they add onto whatever
    // is already there (e.g. "Bacon & Egg Pies TopUp" qty 1 adds 1 onto
    // the "Bacon & Egg Pies" cell). TopUps are applied in a second pass
    // (regardless of the order the app sent them in) so a base-item
    // overwrite in the same request can never clobber a TopUp that
    // landed first.
    items.forEach(function(item) {
      if (baseFoodName(item.food)) { return; }
      var row = getOrCreateFoodRow(sheet, item.food);
      sheet.getRange(row, dateCol).setValue(item.qty || 0);
    });
    items.forEach(function(item) {
      var base = baseFoodName(item.food);
      if (!base) { return; }
      var row = getOrCreateFoodRow(sheet, base);
      var cell = sheet.getRange(row, dateCol);
      var existing = cell.getValue();
      var current = (typeof existing === "number") ? existing : 0;
      cell.setValue(current + (item.qty || 0));
    });

    // "leftover" is the action behind both the 6:30 PM Summary (Mon-Sat)
    // and 6:00 PM Summary (Sunday) buttons — email the spreadsheet on
    // that checkpoint. Wrapped in its own try/catch so a mail failure
    // (quota, permissions) never turns an already-successful sheet write
    // into an error response for the app.
    if (body.action === "leftover") {
      try {
        emailReport(ss);
      } catch (mailErr) {
        Logger.log("Failed to email report: " + mailErr);
      }
    }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// Emails the current hoko_kitchen spreadsheet (exported as .xlsx) to
// REPORT_EMAIL. Runs as whoever the Web App is deployed to execute as.
//
// Note: DriveApp's File.getAs() does not support converting a native
// Google Sheet straight to xlsx ("Converting from
// application/vnd.google-apps.spreadsheet to
// application/vnd.openxmlformats-officedocument.spreadsheetml.sheet is
// not supported"). The reliable workaround is fetching the spreadsheet's
// own export URL with the script's own auth token.
function emailReport(ss) {
  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  });
  var blob = response.getBlob().setName(SPREADSHEET_NAME + ".xlsx");
  MailApp.sendEmail({
    to: REPORT_EMAIL,
    subject: SPREADSHEET_NAME + " Report",
    body: "Hi, Please see attached file for the report",
    attachments: [blob]
  });
}

// Run this one manually (Run > testEmailReport) to both trigger the
// Gmail-send permission prompt the first time, and confirm a real test
// email actually reaches REPORT_EMAIL. emailReport() itself can't be run
// standalone from the editor — it needs the spreadsheet object doPost
// normally passes in.
function testEmailReport() {
  emailReport(getOrCreateSpreadsheet());
  Logger.log("Test email sent to " + REPORT_EMAIL);
}

// Adds a brand-new food item as a row in Food_List and all three data
// tabs (Start_day, Remaining_Foods, Leftover_Foods). Uses the same
// getOrCreateFoodRow as the sync path, so re-adding an item that already
// exists is a harmless no-op rather than a duplicate row.
function handleAddFoodItem(body) {
  var food = (body.food || "").toString().replace(/^\s+|\s+$/g, "");
  if (!food) {
    return jsonResponse({ ok: false, error: "Food name is required" });
  }

  var ss = getOrCreateSpreadsheet();
  var foodListSheet = ss.getSheetByName("Food_List") || ss.insertSheet("Food_List");
  getOrCreateFoodRow(foodListSheet, food);

  DATA_TABS.forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
    getOrCreateFoodRow(sheet, food);
  });

  return jsonResponse({ ok: true });
}

// "Sold Out" button: first checks Start_day for the given date — if every
// quantity there is blank/0 (nothing was ever started today), there's
// nothing to sell out, so this bails with the NO_START_DAY code before
// touching anything else. Otherwise it looks at Remaining_Foods for the
// same date: if every quantity there is blank/0 (nothing was logged as
// remaining), that same Remaining_Foods column is also (re)confirmed as
// all zero. Either way — a sold-out end-of-day means nothing was left
// over — Leftover_Foods for that date is zeroed out too. Finishes by
// emailing the report, same as the leftover checkpoint.
function handleSoldOut(body) {
  var ss = getOrCreateSpreadsheet();
  var dateStr = body.date;

  var startSheet = ss.getSheetByName(ACTION_TO_TAB.start) || ss.insertSheet(ACTION_TO_TAB.start);
  var startCol = getOrCreateDateColumn(startSheet, dateStr);
  var startLastRow = Math.max(startSheet.getLastRow(), 1);

  var startAllZero = true;
  if (startLastRow >= 2) {
    var startValues = startSheet.getRange(2, startCol, startLastRow - 1, 1).getValues();
    for (var s = 0; s < startValues.length; s++) {
      var sv = startValues[s][0];
      if (sv !== "" && sv !== null && Number(sv) !== 0) { startAllZero = false; break; }
    }
  }
  if (startAllZero) {
    return jsonResponse({ ok: false, code: "NO_START_DAY" });
  }

  var remainingSheet = ss.getSheetByName(ACTION_TO_TAB.remaining) || ss.insertSheet(ACTION_TO_TAB.remaining);
  var remainingCol = getOrCreateDateColumn(remainingSheet, dateStr);
  var remainingLastRow = Math.max(remainingSheet.getLastRow(), 1);

  var allZero = true;
  if (remainingLastRow >= 2) {
    var values = remainingSheet.getRange(2, remainingCol, remainingLastRow - 1, 1).getValues();
    for (var r = 0; r < values.length; r++) {
      var v = values[r][0];
      if (v !== "" && v !== null && Number(v) !== 0) { allZero = false; break; }
    }
  }

  if (allZero && remainingLastRow >= 2) {
    remainingSheet.getRange(2, remainingCol, remainingLastRow - 1, 1).setValue(0);
  }

  var leftoverSheet = ss.getSheetByName(ACTION_TO_TAB.leftover) || ss.insertSheet(ACTION_TO_TAB.leftover);
  var leftoverCol = getOrCreateDateColumn(leftoverSheet, dateStr);
  var leftoverLastRow = Math.max(leftoverSheet.getLastRow(), 1);
  if (leftoverLastRow >= 2) {
    leftoverSheet.getRange(2, leftoverCol, leftoverLastRow - 1, 1).setValue(0);
  }

  try {
    emailReport(ss);
  } catch (mailErr) {
    Logger.log("Failed to email report: " + mailErr);
  }

  return jsonResponse({ ok: true });
}

// ---------- reporting ----------

// Builds (or rebuilds) a "<source>_Report(s)" tab with a Date / Total
// Quantity table driven entirely by live formulas that reference the
// source tab, plus a line chart over that table. Because the cells are
// formulas (not copied values), every quantity saved into the source tab
// — today or any future date already in range — recalculates the totals
// and the chart automatically; there is nothing else to sync. The SUM
// range is padded well past the current last row so food items added
// later via "Add Item" are picked up without re-running this.
var REPORT_ROW_PADDING = 200; // generous headroom for future food rows

function buildDailyTotalReport(sourceTabName, reportTabName, chartTitle) {
  var ss = getOrCreateSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceTabName);
  if (!sourceSheet) {
    throw new Error(sourceTabName + " sheet not found — run setup() first.");
  }

  var reportSheet = ss.getSheetByName(reportTabName) || ss.insertSheet(reportTabName);
  reportSheet.getCharts().forEach(function(chart) { reportSheet.removeChart(chart); });
  reportSheet.clear();

  var lastCol = sourceSheet.getLastColumn(); // date columns run B..lastCol
  var sumToRow = Math.max(sourceSheet.getLastRow(), 2) + REPORT_ROW_PADDING;
  var numDates = lastCol - 1;

  reportSheet.getRange(1, 1).setValue("Date").setFontWeight("bold");
  reportSheet.getRange(1, 2).setValue("Total Quantity").setFontWeight("bold");

  for (var i = 0; i < numDates; i++) {
    var sourceCol = i + 2;   // source column (B, C, ...)
    var reportRow = i + 2;   // report row (2, 3, ...)
    var colLetter = columnToLetter(sourceCol);
    reportSheet.getRange(reportRow, 1).setFormula("=" + sourceTabName + "!" + colLetter + "1");
    reportSheet.getRange(reportRow, 2).setFormula(
      "=SUM(" + sourceTabName + "!" + colLetter + "2:" + colLetter + sumToRow + ")"
    );
  }
  reportSheet.getRange(2, 1, numDates, 1).setNumberFormat("@");
  reportSheet.setFrozenRows(1);
  reportSheet.autoResizeColumns(1, 2);

  var dataRange = reportSheet.getRange(1, 1, numDates + 1, 2);
  var chart = reportSheet.newChart()
    .asLineChart()
    .addRange(dataRange)
    .setPosition(2, 4, 0, 0)
    .setOption("title", chartTitle)
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { title: "Date", slantedText: true, textStyle: { fontSize: 9 } })
    .setOption("vAxis", { title: "Total Quantity", minValue: 0 })
    .setOption("width", 900)
    .setOption("height", 420)
    .build();
  reportSheet.insertChart(chart);

  Logger.log(reportTabName + " created/updated: " + ss.getUrl());
}

function setupStartDayReport() {
  buildDailyTotalReport("Start_day", "Start_day_Report", "Start Day — Total Quantity Cooked Per Day");
}

function setupRemainingFoodsReport() {
  buildDailyTotalReport("Remaining_Foods", "Remaining_Foods_Reports", "Remaining Foods — Total Quantity Per Day");
}

function setupLeftoverFoodsReport() {
  buildDailyTotalReport("Leftover_Foods", "Leftover_Foods_Reports", "Leftover Foods — Total Quantity Per Day");
}

// Convenience: builds/rebuilds all three report tabs in one run.
function setupAllReports() {
  setupStartDayReport();
  setupRemainingFoodsReport();
  setupLeftoverFoodsReport();
}

function columnToLetter(column) {
  var temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function getOrCreateDateColumn(sheet, dateStr) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var tz = Session.getScriptTimeZone();
  for (var c = 0; c < headerRow.length; c++) {
    if (headerCellMatches(headerRow[c], dateStr, tz)) { return c + 1; }
  }
  var newCol = lastCol + 1;
  sheet.getRange(1, newCol).setNumberFormat("@").setValue(dateStr).setFontWeight("bold");
  return newCol;
}

// Header cells can be a plain DATE_FORMAT string (correct, once written
// with setNumberFormat("@") first), an old "yyyy-MM-dd" string from
// before the format switch, or a real Date object (any cell Sheets
// managed to auto-convert) — recognize all three so old and new columns
// still match instead of silently duplicating.
function headerCellMatches(cellValue, dateStr, tz) {
  if (cellValue === dateStr) { return true; }
  var cellDate = parseHeaderDate(cellValue);
  var targetDate = parseHeaderDate(dateStr);
  if (!cellDate || !targetDate) { return false; }
  return Utilities.formatDate(cellDate, tz, "yyyy-MM-dd") ===
         Utilities.formatDate(targetDate, tz, "yyyy-MM-dd");
}

// "Bacon & Egg Pies TopUp" -> "Bacon & Egg Pies"; returns null for a
// plain (non-TopUp) item.
var TOPUP_SUFFIX = " TopUp";
function baseFoodName(foodName) {
  if (foodName.length > TOPUP_SUFFIX.length &&
      foodName.slice(-TOPUP_SUFFIX.length) === TOPUP_SUFFIX) {
    return foodName.slice(0, -TOPUP_SUFFIX.length);
  }
  return null;
}

function getOrCreateFoodRow(sheet, foodName) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var col = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var r = 0; r < col.length; r++) {
    if (col[r][0] === foodName) { return r + 1; }
  }
  var newRow = lastRow + 1;
  sheet.getRange(newRow, 1).setValue(foodName);
  return newRow;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
