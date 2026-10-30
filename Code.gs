/**
 * Google Apps Script Backend for CSE (Sections 6, 7, 8) Student Academic Details Portal
 * Target Spreadsheet: "CSE_Master - 29.10.2026.xlsx"
 * Target Worksheet: "CSE_6,7,8"
 */

var TARGET_SPREADSHEET_ID = "1Djj2Lbq14TmJwLRcDzuXi7QsUgnaAunw";
var TARGET_SHEET_GID = 214353412;
var TARGET_FILE_NAME = "CSE_Master - 29.10.2026.xlsx";
var TARGET_SHEET_NAME = "CSE_6,7,8";

/**
 * Serves the HTML Web App UI or handles REST JSON actions.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var res = {};
    try {
      if (action === 'list') {
        res = getStudentList();
      } else if (action === 'student') {
        res = getStudentRecord({ adminNo: e.parameter.adminNo, name: e.parameter.name });
      } else if (action === 'dashboard') {
        res = getDashboardData();
      } else {
        res = { success: false, message: 'Invalid action: ' + action };
      }
    } catch (err) {
      res = { success: false, message: err.toString() };
    }
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Student Data Form – CSE 6, 7, 8')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles POST requests for external or fallback submissions.
 */
function doPost(e) {
  var data = {};
  try {
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    var res = updateStudentRecord(data);
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Resolves the Google Spreadsheet in Google Drive matching CSE_Master.
 */
function getSpreadsheet() {
  // 1. Primary Direct ID Match
  if (TARGET_SPREADSHEET_ID) {
    try {
      var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      if (ss) return ss;
    } catch (e) {
      console.warn("Direct openById failed, checking Drive search: " + e.message);
    }
  }

  var props = PropertiesService.getScriptProperties();
  var cachedId = props.getProperty('SPREADSHEET_ID');

  if (cachedId) {
    try {
      var cachedSs = SpreadsheetApp.openById(cachedId);
      if (cachedSs) return cachedSs;
    } catch (err) {
      // Cached ID is stale or inaccessible, proceed to locate file
    }
  }

  var file = null;

  // 2. Exact match with extension
  var files = DriveApp.getFilesByName(TARGET_FILE_NAME);
  if (files.hasNext()) {
    file = files.next();
  }

  // 3. Match without extension
  if (!file) {
    var nameWithoutExt = TARGET_FILE_NAME.replace(/\.xlsx$/i, '');
    files = DriveApp.getFilesByName(nameWithoutExt);
    if (files.hasNext()) {
      file = files.next();
    }
  }

  // 4. Search query containing CSE_Master
  if (!file) {
    files = DriveApp.searchFiles('title contains "CSE_Master" and trashed = false');
    if (files.hasNext()) {
      file = files.next();
    }
  }

  if (!file) {
    throw new Error(
      'Target spreadsheet "' + TARGET_FILE_NAME + '" not found in Google Drive. ' +
      'Please verify access to Spreadsheet ID: ' + TARGET_SPREADSHEET_ID
    );
  }

  var fileId = file.getId();
  var spreadsheet = null;

  try {
    spreadsheet = SpreadsheetApp.openById(fileId);
  } catch (openErr) {
    // If direct open of .xlsx fails, check if a converted Google Sheet copy exists
    var convertedName = file.getName().replace(/\.xlsx$/i, '') + ' (Google Sheet)';
    var convFiles = DriveApp.getFilesByName(convertedName);
    if (convFiles.hasNext()) {
      spreadsheet = SpreadsheetApp.openById(convFiles.next().getId());
    } else {
      throw new Error(
        'Could not directly open "' + file.getName() + '": ' + openErr.message +
        '. Please open the Excel file once in Google Drive to initialize it as a Google Sheet.'
      );
    }
  }

  if (spreadsheet) {
    props.setProperty('SPREADSHEET_ID', spreadsheet.getId());
  }

  return spreadsheet;
}

/**
 * Resolves the target worksheet CSE_6,7,8.
 */
function getWorksheet(ss) {
  // 1. Direct match by worksheet GID
  if (TARGET_SHEET_GID !== undefined && TARGET_SHEET_GID !== null) {
    var allSheets = ss.getSheets();
    for (var k = 0; k < allSheets.length; k++) {
      if (allSheets[k].getSheetId() === TARGET_SHEET_GID) {
        return allSheets[k];
      }
    }
  }

  // 2. Match by exact worksheet name
  var sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (sheet) return sheet;

  // 3. Fuzzy match case/whitespace
  var sheets = ss.getSheets();
  var normalizedTarget = TARGET_SHEET_NAME.replace(/\s+/g, '').toLowerCase();

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().replace(/\s+/g, '').toLowerCase();
    if (sName === normalizedTarget) {
      return sheets[i];
    }
  }

  // 4. Check if name contains 6, 7, 8
  for (var j = 0; j < sheets.length; j++) {
    var nameLower = sheets[j].getName().toLowerCase();
    if (nameLower.indexOf('6') !== -1 && (nameLower.indexOf('7') !== -1 || nameLower.indexOf('8') !== -1)) {
      return sheets[j];
    }
  }

  // Fallback to active/first sheet
  return sheets[0];
}

/**
 * Dynamically resolves column positions and data rows.
 */
function getSheetContext() {
  var ss = getSpreadsheet();
  var sheet = getWorksheet(ss);
  var values = sheet.getDataRange().getValues();

  if (!values || values.length === 0) {
    throw new Error('Worksheet "' + sheet.getName() + '" has no data.');
  }

  // Detect header row (first 5 rows)
  var headerRowIdx = 0;
  for (var r = 0; r < Math.min(5, values.length); r++) {
    var rowStr = values[r].join(' ').toLowerCase();
    if (
      rowStr.indexOf('name') !== -1 ||
      rowStr.indexOf('admin') !== -1 ||
      rowStr.indexOf('student') !== -1 ||
      rowStr.indexOf('s.no') !== -1
    ) {
      headerRowIdx = r;
      break;
    }
  }

  var headers = values[headerRowIdx];

  // Default standard 1-based columns (Cols A to H)
  var cols = {
    sn: 1,
    name: 2,
    adminNo: 3,
    category: 4,
    ssc: 5,
    inter: 6,
    eamcet: 7,
    ruralUrban: 8
  };

  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim().toLowerCase();
    if (h === 's.no' || h === 'sno' || h === 's. no' || h === 'sl.no' || h === 'sl no') {
      cols.sn = c + 1;
    } else if (h.indexOf('name') !== -1 && (h.indexOf('student') !== -1 || h === 'name')) {
      cols.name = c + 1;
    } else if (
      h.indexOf('admin') !== -1 ||
      h.indexOf('admn') !== -1 ||
      h.indexOf('admission') !== -1 ||
      h.indexOf('roll') !== -1
    ) {
      cols.adminNo = c + 1;
    } else if (h.indexOf('categ') !== -1) {
      cols.category = c + 1;
    } else if (h.indexOf('ssc') !== -1 || h.indexOf('10th') !== -1) {
      cols.ssc = c + 1;
    } else if (h.indexOf('inter') !== -1 || h.indexOf('12th') !== -1) {
      cols.inter = c + 1;
    } else if (h.indexOf('eamcet') !== -1 || h.indexOf('rank') !== -1) {
      cols.eamcet = c + 1;
    } else if (h.indexOf('rural') !== -1 || h.indexOf('urban') !== -1) {
      cols.ruralUrban = c + 1;
    }
  }

  return {
    ss: ss,
    sheet: sheet,
    values: values,
    headerRowIdx: headerRowIdx,
    dataStartRowIdx: headerRowIdx + 1, // 0-based index for values array
    cols: cols
  };
}

/**
 * Returns full list of students for the search autocomplete dropdown.
 */
function getStudentList() {
  try {
    var ctx = getSheetContext();
    var values = ctx.values;
    var cols = ctx.cols;
    var students = [];

    for (var i = ctx.dataStartRowIdx; i < values.length; i++) {
      var row = values[i];
      var name = String(row[cols.name - 1] || '').trim();
      var adminNo = String(row[cols.adminNo - 1] || '').trim();
      var category = String(row[cols.category - 1] || '').trim();

      if (name || adminNo) {
        students.push({
          name: name,
          adminNo: adminNo,
          category: category
        });
      }
    }

    return {
      success: true,
      students: students
    };
  } catch (err) {
    return {
      success: false,
      message: err.message
    };
  }
}

/**
 * Fetches an individual student's existing record by Admission Number or Name.
 */
function getStudentRecord(query) {
  try {
    if (!query || (!query.adminNo && !query.name)) {
      return { success: false, message: 'Student Admission Number or Name is required.' };
    }

    var ctx = getSheetContext();
    var values = ctx.values;
    var cols = ctx.cols;
    var qAdmin = String(query.adminNo || '').trim().toLowerCase();
    var qName = String(query.name || '').trim().toLowerCase();

    for (var i = ctx.dataStartRowIdx; i < values.length; i++) {
      var row = values[i];
      var rAdmin = String(row[cols.adminNo - 1] || '').trim().toLowerCase();
      var rName = String(row[cols.name - 1] || '').trim().toLowerCase();

      var match = false;
      if (qAdmin && rAdmin && qAdmin === rAdmin) {
        match = true;
      } else if (qName && rName && qName === rName) {
        match = true;
      }

      if (match) {
        return {
          success: true,
          student: {
            sn: String(row[cols.sn - 1] || ''),
            name: String(row[cols.name - 1] || ''),
            adminNo: String(row[cols.adminNo - 1] || ''),
            category: String(row[cols.category - 1] || ''),
            sscMarks: String(row[cols.ssc - 1] || ''),
            interMarks: String(row[cols.inter - 1] || ''),
            eamcetRank: String(row[cols.eamcet - 1] || ''),
            ruralUrban: String(row[cols.ruralUrban - 1] || ''),
            rowNumber: i + 1
          }
        };
      }
    }

    return {
      success: false,
      message: 'Student record could not be found.'
    };
  } catch (err) {
    return {
      success: false,
      message: err.message
    };
  }
}

/**
 * Updates Columns E, F, G, H for a student with concurrency locking.
 */
function updateStudentRecord(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return {
      success: false,
      message: 'Server is currently busy with other updates. Please try again in a few seconds.'
    };
  }

  try {
    if (!payload || (!payload.adminNo && !payload.name)) {
      return { success: false, message: 'Invalid payload: Admission Number or Name required.' };
    }

    var ctx = getSheetContext();
    var sheet = ctx.sheet;
    var values = ctx.values;
    var cols = ctx.cols;

    var qAdmin = String(payload.adminNo || '').trim().toLowerCase();
    var qName = String(payload.name || '').trim().toLowerCase();
    var targetRow = -1;
    var matchedRowData = null;

    for (var i = ctx.dataStartRowIdx; i < values.length; i++) {
      var row = values[i];
      var rAdmin = String(row[cols.adminNo - 1] || '').trim().toLowerCase();
      var rName = String(row[cols.name - 1] || '').trim().toLowerCase();

      var match = false;
      if (qAdmin && rAdmin && qAdmin === rAdmin) {
        match = true;
      } else if (qName && rName && qName === rName) {
        match = true;
      }

      if (match) {
        targetRow = i + 1; // 1-based row index for sheet
        matchedRowData = row;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'Unable to locate student record to update.' };
    }

    // Update Columns E, F, G, H
    sheet.getRange(targetRow, cols.ssc).setValue(payload.sscMarks);
    sheet.getRange(targetRow, cols.inter).setValue(payload.interMarks);
    sheet.getRange(targetRow, cols.eamcet).setValue(payload.eamcetRank);
    sheet.getRange(targetRow, cols.ruralUrban).setValue(payload.ruralUrban);

    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Student details recorded successfully in CSE_Master.',
      updatedRecord: {
        sn: String(matchedRowData[cols.sn - 1] || ''),
        name: String(matchedRowData[cols.name - 1] || payload.name),
        adminNo: String(matchedRowData[cols.adminNo - 1] || payload.adminNo),
        category: String(matchedRowData[cols.category - 1] || ''),
        sscMarks: payload.sscMarks,
        interMarks: payload.interMarks,
        eamcetRank: payload.eamcetRank,
        ruralUrban: payload.ruralUrban
      }
    };
  } catch (err) {
    return {
      success: false,
      message: 'Failed to update record: ' + err.message
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns all student records and progress statistics for the Admin Dashboard.
 */
function getDashboardData() {
  try {
    var ctx = getSheetContext();
    var values = ctx.values;
    var cols = ctx.cols;
    var records = [];
    var completedCount = 0;

    for (var i = ctx.dataStartRowIdx; i < values.length; i++) {
      var row = values[i];
      var name = String(row[cols.name - 1] || '').trim();
      var adminNo = String(row[cols.adminNo - 1] || '').trim();

      if (!name && !adminNo) continue;

      var ssc = String(row[cols.ssc - 1] || '').trim();
      var inter = String(row[cols.inter - 1] || '').trim();
      var eamcet = String(row[cols.eamcet - 1] || '').trim();
      var ruralUrban = String(row[cols.ruralUrban - 1] || '').trim();

      var isComplete = Boolean(ssc && inter && eamcet && ruralUrban);
      if (isComplete) completedCount++;

      records.push({
        sn: String(row[cols.sn - 1] || (records.length + 1)),
        name: name,
        adminNo: adminNo,
        category: String(row[cols.category - 1] || ''),
        sscMarks: ssc,
        interMarks: inter,
        eamcetRank: eamcet,
        ruralUrban: ruralUrban,
        isComplete: isComplete
      });
    }

    var total = records.length;
    var pending = total - completedCount;
    var percentage = total > 0 ? Math.round((completedCount / total) * 100) : 0;

    return {
      success: true,
      records: records,
      stats: {
        total: total,
        completed: completedCount,
        pending: pending,
        percentage: percentage
      }
    };
  } catch (err) {
    return {
      success: false,
      message: err.message
    };
  }
}
