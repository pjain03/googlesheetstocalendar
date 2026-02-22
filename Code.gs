/**
 * Syncs birthdays from a Google Sheet to a Google Calendar.
 * Automatically triggered by Google Apps Script onEdit (if set up as an installable trigger)
 * or can be run manually from the menu.
 */

// Global settings
const SHEET_URL = ''; // Optional: Paste your full Google Sheets URL here (e.g., https://docs.google.com/spreadsheets/d/1BxiMVs0XRYFgCE_/edit#gid=12345). Leave blank if script is bound to the sheet and using the default first tab.
const ROW_START = 2; // Row 1 is usually headers
const CALENDAR_ID = ''; // REQUIRED: Paste a specific Calendar ID here (e.g. 'c_123456@group.calendar.google.com'). Do not leave blank.

/**
 * Helper to parse the Spreadsheet ID and Worksheet ID (gid) from a URL
 */
function parseSheetUrl() {
  if (!SHEET_URL) return { spreadsheetId: null, worksheetId: 0 };
  
  var result = { spreadsheetId: null, worksheetId: 0 };
  
  // Extract Spreadsheet ID
  var idMatch = SHEET_URL.match(/\/d\/(.*?)(\/|$)/);
  if (idMatch && idMatch[1]) {
    result.spreadsheetId = idMatch[1];
  }
  
  // Extract Worksheet ID (gid)
  var gidMatch = SHEET_URL.match(/[#&]gid=([0-9]+)/);
  if (gidMatch && gidMatch[1]) {
    result.worksheetId = parseInt(gidMatch[1], 10);
  }
  
  return result;
}



/**
 * Isolates date parsing from the sheet.
 * @param {string} dateStr The string representation of the date.
 * @returns {Object|null} An object { year, month, day } or null if invalid.
 */
function getDate(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return null;
  dateStr = dateStr.trim();
  if (dateStr === "") return null;

  var dateParts = dateStr.includes('/') ? dateStr.split('/') : dateStr.split('-');
  if (dateParts.length < 2) return null;

  var day = parseInt(dateParts[0], 10);
  var month = parseInt(dateParts[1], 10);
  var year = dateParts.length >= 3 ? parseInt(dateParts[2], 10) : new Date().getFullYear();

  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
  if (month < 1 || month > 12) return null;

  // Validate day bounds, strictly checking max days for that specific month & year (leap year aware)
  var daysInMonth = new Date(year, month, 0).getDate();
  if (day < 1 || day > daysInMonth) return null;

  return { day: day, month: month, year: year };
}

function onInstallableEdit(e) {
  // We only run sync if the edit happened on the correct sheet
  if (e && e.source) {
    var spreadsheet = e.source;
    var sheet = spreadsheet.getActiveSheet();
    var parsed = parseSheetUrl();
    
    // Verify Spreadsheet ID if provided
    if (parsed.spreadsheetId && spreadsheet.getId() !== parsed.spreadsheetId) {
      return;
    }
    
    // Verify Worksheet ID
    if (sheet.getSheetId() !== parsed.worksheetId) {
      return; 
    }
  }

  // Polling Debouncer Logic
  var props = PropertiesService.getScriptProperties();
  var myTime = new Date().getTime().toString();
  props.setProperty('LAST_EDIT_TIME', myTime);
  
  // Enter the polling loop
  while (true) {
    Utilities.sleep(2000); // Wait 2 seconds
    var currentTimeInProps = props.getProperty('LAST_EDIT_TIME');
    
    if (currentTimeInProps !== myTime) {
       // A newer edit has arrived since we went to sleep!
       // This execution is no longer the final one. We can confidently terminate it.
       Logger.log("A newer edit was detected. Terminating this execution block early.");
       return;
    } else {
       // Our time is still the latest time across all executions!
       // No new edits have happened in the last 2 seconds. Break the loop and execute the sync.
       break;
    }
  }

  // If we reach here, we are the undisputed final edit in a continuous block of changes.
  syncBirthdays();
}

/**
 * Creates a custom menu in Google Sheets
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Birthday Sync')
    .addItem('Sync Now', 'syncBirthdays')
    .addItem('Setup Trigger', 'setupTrigger')
    .addToUi();
}

/**
 * Syncs the entire "Birthdays" sheet to the User's Default Calendar
 * Handles DD/MM and DD/MM/YYYY formats
 * Ensures full valid state (removes deleted rows from calendar)
 */
function syncBirthdays() {
  if (!CALENDAR_ID) {
    Logger.log("CRITICAL ERROR: CALENDAR_ID is required. Execution aborted to prevent accidental primary calendar wipes.");
    var ui = SpreadsheetApp.getUi();
    if (ui) ui.alert("CRITICAL ERROR: CALENDAR_ID cannot be blank. You must specify a dedicated Calendar ID to prevent wiping your primary calendar.");
    return;
  }

  var targetCalendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!targetCalendar) {
    Logger.log("Calendar not found. Verify your CALENDAR_ID is correct and this account has edit access to it.");
    var ui = SpreadsheetApp.getUi();
    if (ui) ui.alert("Error: Calendar not found. Verify your CALENDAR_ID.");
    return;
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(2000); // Wait up to 2 seconds for other executions to finish
  } catch (e) {
    Logger.log('Could not obtain lock after 2 seconds. Skipping this concurrent execution.');
    return;
  }

  try {
    var parsed = parseSheetUrl();
    var ss = parsed.spreadsheetId ? SpreadsheetApp.openById(parsed.spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log("Spreadsheet not found.");
      return;
    }

  // Find sheet by its ID (gid)
  var sheets = ss.getSheets();
  var sheet = null;
  for (var k = 0; k < sheets.length; k++) {
    if (sheets[k].getSheetId() === parsed.worksheetId) {
      sheet = sheets[k];
      break;
    }
  }

  if (!sheet) {
    Logger.log("Worksheet with ID " + parsed.worksheetId + " not found.");
    return;
  }

  var dataRange = sheet.getDataRange();
  // Use getDisplayValues() to force all cells to return their string representations exactly as seen.
  var values = dataRange.getDisplayValues();
  
  // ============================================
  // FULL STATELESS RESET: Delete EVERY event currently on this calendar!
  // ============================================
  
  var now = new Date();
  var searchStart = new Date(now.getFullYear() - 1, 0, 1); // Jan 1st of the previous year
  var hundredYearsFromNow = new Date();
  hundredYearsFromNow.setFullYear(now.getFullYear() + 100);

  // Fetch all events currently on the calendar and delete them (wipes the entire calendar clean)
  var existingEvents = targetCalendar.getEvents(searchStart, hundredYearsFromNow);
  var deletedSeriesIds = {}; // Track series we've already wiped to avoid duplicate, error-throwing deletion attempts

  for (var j = 0; j < existingEvents.length; j++) {
      var evt = existingEvents[j];
      try {
        // If it's part of a series (which our script makes), deleting the series from this event completely removes it
        var series = evt.getEventSeries();
        if (series) {
            var seriesId = series.getId();
            if (!deletedSeriesIds[seriesId]) {
                series.deleteEventSeries();
                deletedSeriesIds[seriesId] = true;
                Logger.log("Deleted old event series: " + evt.getTitle());
            }
        } else {
            evt.deleteEvent();
        }
      } catch (err) {
         Logger.log("Could not delete event: " + err.message);
      }
  }

  // Iterate over all rows starting from ROW_START
  for (var i = ROW_START - 1; i < values.length; i++) {
    try {
      var row = values[i];
    
    // Columns: A=Name(0), B=Date(1)
    var name = row[0] ? row[0].toString().trim() : "";
    var dateStr = row[1] ? row[1].toString().trim() : "";

    if (!name || !dateStr) {
      continue; // Skip empty rows
    }

    // Attempt to extract the date from the string
    var parsedDate = getDate(dateStr);
    
    if (!parsedDate) {
      Logger.log("Skipping row " + (i+1) + " due to invalid date format: " + dateStr);
      continue;
    }

    var day = parsedDate.day;
    var month = parsedDate.month;
    var year = parsedDate.year;

    // We will create the start date for the recurring event
    // Create it at noon to avoid timezone shift issues ending up on previous day
    var startDate = new Date(year, month - 1, day, 12, 0, 0, 0); 
    var eventTitle = "🎈 " + name + "'s Birthday";

    // Create new annually recurring all-day event
    var recurrence = CalendarApp.newRecurrence().addYearlyRule();
    targetCalendar.createAllDayEventSeries(eventTitle, startDate, recurrence);

    Logger.log("Created Event for " + name);
    } catch(rowErr) {
       Logger.log("Error processing row " + (i+1) + ": " + rowErr.message);
    }
  }

  } finally {
    lock.releaseLock();
  }
}

/**
 * Utility function to set up the Installable onEdit Trigger.
 * Needs to be run manually once to grant permissions.
 */
function setupTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  
  // Prevent duplicate triggers
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onInstallableEdit') {
      SpreadsheetApp.getUi().alert("Trigger is already set up!");
      return;
    }
  }
  
  ScriptApp.newTrigger('onInstallableEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  SpreadsheetApp.getUi().alert("Trigger successfully set up! Birthdays will now sync automatically on edit.");
}
