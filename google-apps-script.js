// ============================================================
// VZ BUFF AND POLISH — Google Apps Script
// Paste this into your Google Sheet: Extensions → Apps Script
// Deploy → New Deployment → Web App → Execute as: Me → Anyone
// ============================================================
// Current Apps Script URL:
// https://script.google.com/macros/s/AKfycbz2kzfjE3_-mb4mcXKuzpA4EDBYrD6oBfxaaTGVFbwhUu5N9_jssIqW6Iq87AqS3GDu/exec
// ============================================================
// SETUP:
// 1. Google Sheet: "VZ - Buff, Polish, and Repair"
//    - "Inbound Unit Info" headers (row 1):
//      A: IMEI | B: OEM | C: SKU | D: MODEL | E: Order Type | F: DISPO IN
//    - "Transaction Log" (auto-created) headers (row 1):
//      A: IMEI | B: Page | C: Parameter | D: Result | E: User | F: Timestamp | G: Note
// 2. Device data should already be in the "Inbound Unit Info" tab
// ============================================================

function _respond(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'read';

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---- READ ALL DEVICES ----
    if (action === 'readalldevices') {
      var diSheet = ss.getSheetByName('Inbound Unit Info');
      if (!diSheet) return _respond({ success: false, error: 'Sheet "Inbound Unit Info" not found' }, callback);

      var lastRow = diSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: true, data: [] }, callback);

      var data = diSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      var headers = ['IMEI', 'OEM', 'SKU', 'MODEL', 'Order Type', 'DISPO IN'];
      var rows = data.map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { obj[h] = row[i] ? row[i].toString() : ''; });
        return obj;
      });

      return _respond({ success: true, data: rows }, callback);
    }

    // ---- READ ALL STATUSES (latest result per IMEI+Parameter) ----
    if (action === 'readallstatuses') {
      var tlSheet = ss.getSheetByName('Transaction Log');
      if (!tlSheet) return _respond({ success: true, data: {} }, callback);

      var lastRow = tlSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: true, data: {} }, callback);

      var data = tlSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      var allStatuses = {}; // { imei: { param: { result, user, timestamp, note } } }

      for (var i = 0; i < data.length; i++) {
        var imei = data[i][0].toString().trim();
        if (!imei) continue;
        if (!allStatuses[imei]) allStatuses[imei] = {};
        var param = data[i][2].toString().trim();
        allStatuses[imei][param] = {
          page: data[i][1].toString(),
          result: data[i][3].toString(),
          user: data[i][4].toString(),
          timestamp: data[i][5].toString(),
          note: data[i][6] ? data[i][6].toString() : ''
        };
      }

      return _respond({ success: true, data: allStatuses }, callback);
    }

    // ---- LOOKUP DEVICE BY IMEI ----
    if (action === 'lookupdevice') {
      var imei = (e.parameter.imei || '').toString().trim();
      if (!imei) return _respond({ success: false, error: 'IMEI required' }, callback);

      var diSheet = ss.getSheetByName('Inbound Unit Info');
      if (!diSheet) return _respond({ success: false, error: 'Sheet "Inbound Unit Info" not found' }, callback);

      var lastRow = diSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: false, error: 'No device data found' }, callback);

      var imeiCol = diSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var rowIndex = -1;
      for (var i = 0; i < imeiCol.length; i++) {
        if (imeiCol[i][0].toString().trim() === imei) {
          rowIndex = i;
          break;
        }
      }

      if (rowIndex === -1) return _respond({ success: false, error: 'IMEI not found: ' + imei }, callback);

      var rowData = diSheet.getRange(rowIndex + 2, 1, 1, 6).getValues()[0];
      var device = {
        'IMEI': rowData[0] ? rowData[0].toString() : '',
        'OEM': rowData[1] ? rowData[1].toString() : '',
        'SKU': rowData[2] ? rowData[2].toString() : '',
        'MODEL': rowData[3] ? rowData[3].toString() : '',
        'Order Type': rowData[4] ? rowData[4].toString() : '',
        'DISPO IN': rowData[5] ? rowData[5].toString() : ''
      };

      return _respond({ success: true, data: device }, callback);
    }

    // ---- LOG TEST RESULT (single parameter) ----
    if (action === 'logtest') {
      var tlSheet = ss.getSheetByName('Transaction Log');
      if (!tlSheet) {
        tlSheet = ss.insertSheet('Transaction Log');
        tlSheet.getRange(1, 1, 1, 7).setValues([['IMEI', 'Page', 'Parameter', 'Result', 'User', 'Timestamp', 'Note']]);
      } else {
        // Ensure Note header exists (in case tab was created before note support)
        var headerG = tlSheet.getRange(1, 7).getValue();
        if (!headerG) tlSheet.getRange(1, 7).setValue('Note');
      }

      var imei = (e.parameter.imei || '').toString().trim();
      var page = (e.parameter.page || '').toString().trim();
      var param = (e.parameter.parameter || '').toString().trim();
      var result = (e.parameter.result || '').toString().trim();
      var user = (e.parameter.user || '').toString().trim();
      var note = (e.parameter.note || '').toString().trim();

      if (!imei) return _respond({ success: false, error: 'IMEI required' }, callback);
      if (!param) return _respond({ success: false, error: 'Parameter required' }, callback);
      if (!result) return _respond({ success: false, error: 'Result required' }, callback);
      if (!user) return _respond({ success: false, error: 'User required' }, callback);

      var now = new Date();
      var ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d/yyyy HH:mm:ss');

      tlSheet.appendRow([imei, page, param, result, user, ts, note]);

      // Force IMEI column to text
      var lastRow = tlSheet.getLastRow();
      tlSheet.getRange(lastRow, 1).setNumberFormat('@');
      tlSheet.getRange(lastRow, 1).setValue(imei);

      return _respond({ success: true, message: 'Test logged' }, callback);
    }

    // ---- LOG BATCH TEST RESULTS (multiple parameters at once) ----
    if (action === 'logbatch') {
      var tlSheet = ss.getSheetByName('Transaction Log');
      if (!tlSheet) {
        tlSheet = ss.insertSheet('Transaction Log');
        tlSheet.getRange(1, 1, 1, 6).setValues([['IMEI', 'Page', 'Parameter', 'Result', 'User', 'Timestamp']]);
      }

      var imei = (e.parameter.imei || '').toString().trim();
      var page = (e.parameter.page || '').toString().trim();
      var user = (e.parameter.user || '').toString().trim();
      var testsJson = (e.parameter.tests || '[]').toString();

      if (!imei) return _respond({ success: false, error: 'IMEI required' }, callback);
      if (!user) return _respond({ success: false, error: 'User required' }, callback);

      var tests = JSON.parse(testsJson);
      if (!tests.length) return _respond({ success: false, error: 'No tests provided' }, callback);

      var now = new Date();
      var ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d/yyyy HH:mm:ss');

      var rows = tests.map(function(t) {
        return [imei, page || t.page || '', t.parameter || '', t.result || '', user, ts];
      });

      tlSheet.getRange(tlSheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);

      // Format IMEI column as text for all new rows
      var startRow = tlSheet.getLastRow() - rows.length + 1;
      var imeiRange = tlSheet.getRange(startRow, 1, rows.length, 1);
      imeiRange.setNumberFormat('@');

      return _respond({ success: true, message: rows.length + ' tests logged' }, callback);
    }

    // ---- GET CURRENT STATUS FOR AN IMEI (latest result per parameter) ----
    if (action === 'getstatus') {
      var imei = (e.parameter.imei || '').toString().trim();
      if (!imei) return _respond({ success: false, error: 'IMEI required' }, callback);

      var tlSheet = ss.getSheetByName('Transaction Log');
      if (!tlSheet) return _respond({ success: true, data: {} }, callback);

      var lastRow = tlSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: true, data: {} }, callback);

      var data = tlSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      var current = {};

      // Walk through all rows; later entries overwrite earlier ones (most recent wins)
      for (var i = 0; i < data.length; i++) {
        if (data[i][0].toString().trim() === imei) {
          var param = data[i][2].toString().trim();
          current[param] = {
            page: data[i][1].toString(),
            result: data[i][3].toString(),
            user: data[i][4].toString(),
            timestamp: data[i][5].toString(),
            note: data[i][6] ? data[i][6].toString() : ''
          };
        }
      }

      return _respond({ success: true, data: current }, callback);
    }

    // ---- GET FULL HISTORY FOR AN IMEI ----
    if (action === 'gethistory') {
      var imei = (e.parameter.imei || '').toString().trim();
      if (!imei) return _respond({ success: false, error: 'IMEI required' }, callback);

      var tlSheet = ss.getSheetByName('Transaction Log');
      if (!tlSheet) return _respond({ success: true, data: [] }, callback);

      var lastRow = tlSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: true, data: [] }, callback);

      var data = tlSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      var headers = ['IMEI', 'Page', 'Parameter', 'Result', 'User', 'Timestamp', 'Note'];
      var rows = [];

      for (var i = 0; i < data.length; i++) {
        if (data[i][0].toString().trim() === imei) {
          var obj = {};
          headers.forEach(function(h, idx) { obj[h] = data[i][idx] ? data[i][idx].toString() : ''; });
          rows.push(obj);
        }
      }

      return _respond({ success: true, data: rows }, callback);
    }

    // ---- READ ALL TRANSACTION LOG (for reporting) ----
    if (action === 'readlog') {
      var tlSheet = ss.getSheetByName('Transaction Log');
      if (!tlSheet) return _respond({ success: true, data: [] }, callback);

      var lastRow = tlSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: true, data: [] }, callback);

      var data = tlSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      var headers = ['IMEI', 'Page', 'Parameter', 'Result', 'User', 'Timestamp', 'Note'];
      var rows = data.map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { obj[h] = row[i] ? row[i].toString() : ''; });
        return obj;
      });

      return _respond({ success: true, data: rows }, callback);
    }

    // ---- UNKNOWN ACTION ----
    return _respond({ success: false, error: 'Unknown action: ' + action }, callback);

  } catch (err) {
    return _respond({ success: false, error: err.message || err.toString() }, callback);
  }
}
