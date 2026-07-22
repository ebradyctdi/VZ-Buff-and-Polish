// ============================================================
// VZ BUFF AND POLISH — Google Apps Script
// Paste this into your Google Sheet: Extensions → Apps Script
// Deploy → New Deployment → Web App → Execute as: Me → Anyone
// ============================================================
// Current Apps Script URL:
// https://script.google.com/macros/s/AKfycbyWkDLEA_EGz3OhWNHWzllRfOnKqn8wuJQB2GSAeyvuG2vjJvBCuTaW-E7MTozqEidS/exec
// ============================================================
// SETUP:
// 1. Google Sheet: "VZ - Buff, Polish, and Repair"
//    - "Inbound Unit Info" headers (row 1):
//      A: IMEI | B: OEM | C: SKU | D: MODEL | E: Order Type | F: DISPO IN
//    - "Transaction Log" (auto-created) headers (row 1):
//      A: IMEI | B: Page | C: Task ID | D: Result | E: User | F: Timestamp | G: Note
//    - "Task Configuration" (auto-created) headers (row 1):
//      A: Task ID | B: Task Name | C: Label | D: Submenu | E: Page | F: Result Type
//      G: Options | H: CPO | I: CLNR | J: Sort Order | K: Active
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

    // ---- READ ALL STATUSES (latest result per IMEI+Task ID) ----
    if (action === 'readallstatuses') {
      var tlSheet = ss.getSheetByName('Transaction Log');
      if (!tlSheet) return _respond({ success: true, data: {} }, callback);

      var lastRow = tlSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: true, data: {} }, callback);

      // Build Task Name → Task ID map for backward compatibility
      var taskNameToId = {};
      var tcSheet = ss.getSheetByName('Task Configuration');
      if (tcSheet && tcSheet.getLastRow() >= 2) {
        var tcData = tcSheet.getRange(2, 1, tcSheet.getLastRow() - 1, 2).getValues();
        tcData.forEach(function(row) {
          if (row[0] && row[1]) taskNameToId[row[1].toString().trim()] = row[0].toString().trim();
        });
      }

      var data = tlSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      var allStatuses = {}; // { imei: { taskid: { result, user, timestamp, note } } }

      for (var i = 0; i < data.length; i++) {
        var imei = data[i][0].toString().trim();
        if (!imei) continue;
        if (!allStatuses[imei]) allStatuses[imei] = {};
        var rawId = data[i][2].toString().trim();
        // If rawId matches a Task Name, map it to the Task ID
        var taskid = taskNameToId[rawId] || rawId;
        allStatuses[imei][taskid] = {
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
        tlSheet.getRange(1, 1, 1, 7).setValues([['IMEI', 'Page', 'Task ID', 'Result', 'User', 'Timestamp', 'Note']]);
      } else {
        // Ensure Note header exists (in case tab was created before note support)
        var headerG = tlSheet.getRange(1, 7).getValue();
        if (!headerG) tlSheet.getRange(1, 7).setValue('Note');
        // Update column C header if it says Parameter
        var headerC = tlSheet.getRange(1, 3).getValue();
        if (headerC === 'Parameter') tlSheet.getRange(1, 3).setValue('Task ID');
      }

      var imei = (e.parameter.imei || '').toString().trim();
      var page = (e.parameter.page || '').toString().trim();
      var taskid = (e.parameter.taskid || e.parameter.parameter || '').toString().trim();
      var result = (e.parameter.result || '').toString().trim();
      var user = (e.parameter.user || '').toString().trim();
      var note = (e.parameter.note || '').toString().trim();

      if (!imei) return _respond({ success: false, error: 'IMEI required' }, callback);
      if (!taskid) return _respond({ success: false, error: 'Task ID required' }, callback);
      if (!result) return _respond({ success: false, error: 'Result required' }, callback);
      if (!user) return _respond({ success: false, error: 'User required' }, callback);

      var now = new Date();
      var ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d/yyyy HH:mm:ss');

      tlSheet.appendRow([imei, page, taskid, result, user, ts, note]);

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

    // ---- GET CURRENT STATUS FOR AN IMEI (latest result per Task ID) ----
    if (action === 'getstatus') {
      var imei = (e.parameter.imei || '').toString().trim();
      if (!imei) return _respond({ success: false, error: 'IMEI required' }, callback);

      var tlSheet = ss.getSheetByName('Transaction Log');
      if (!tlSheet) return _respond({ success: true, data: {} }, callback);

      var lastRow = tlSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: true, data: {} }, callback);

      // Build Task Name → Task ID map for backward compatibility
      var taskNameToId = {};
      var tcSheet = ss.getSheetByName('Task Configuration');
      if (tcSheet && tcSheet.getLastRow() >= 2) {
        var tcData = tcSheet.getRange(2, 1, tcSheet.getLastRow() - 1, 2).getValues();
        tcData.forEach(function(row) {
          if (row[0] && row[1]) taskNameToId[row[1].toString().trim()] = row[0].toString().trim();
        });
      }

      var data = tlSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      var current = {};

      // Walk through all rows; later entries overwrite earlier ones (most recent wins)
      for (var i = 0; i < data.length; i++) {
        if (data[i][0].toString().trim() === imei) {
          var rawId = data[i][2].toString().trim();
          var taskid = taskNameToId[rawId] || rawId;
          current[taskid] = {
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

      // Build Task Name → Task ID map for backward compatibility
      var taskNameToId = {};
      var tcSheet = ss.getSheetByName('Task Configuration');
      if (tcSheet && tcSheet.getLastRow() >= 2) {
        var tcData = tcSheet.getRange(2, 1, tcSheet.getLastRow() - 1, 2).getValues();
        tcData.forEach(function(row) {
          if (row[0] && row[1]) taskNameToId[row[1].toString().trim()] = row[0].toString().trim();
        });
      }

      var data = tlSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      var rows = data.map(function(row) {
        var rawId = row[2] ? row[2].toString().trim() : '';
        var taskId = taskNameToId[rawId] || rawId;
        return {
          'IMEI': row[0] ? row[0].toString() : '',
          'Page': row[1] ? row[1].toString() : '',
          'Task ID': taskId,
          'Result': row[3] ? row[3].toString() : '',
          'User': row[4] ? row[4].toString() : '',
          'Timestamp': row[5] ? row[5].toString() : '',
          'Note': row[6] ? row[6].toString() : ''
        };
      });

      return _respond({ success: true, data: rows }, callback);
    }

    // ---- READ TASK CONFIGURATION ----
    if (action === 'readtasks') {
      var tcSheet = ss.getSheetByName('Task Configuration');
      if (!tcSheet) return _respond({ success: true, data: [] }, callback);
      var lastRow = tcSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: true, data: [] }, callback);
      var data = tcSheet.getRange(2, 1, lastRow - 1, 11).getValues();
      var headers = ['Task ID', 'Task Name', 'Label', 'Submenu', 'Page', 'Result Type', 'Options', 'CPO', 'CLNR', 'Sort Order', 'Active'];
      var rows = data.filter(function(row) { return row[0]; }).map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { obj[h] = row[i] !== null && row[i] !== undefined ? row[i].toString() : ''; });
        return obj;
      });
      return _respond({ success: true, data: rows }, callback);
    }

    // ---- READ TASKS FOR A SPECIFIC PAGE ----
    if (action === 'readpagetasks') {
      var page = (e.parameter.page || '').toString().trim();
      if (!page) return _respond({ success: false, error: 'Page required' }, callback);
      var tcSheet = ss.getSheetByName('Task Configuration');
      if (!tcSheet) return _respond({ success: true, data: [] }, callback);
      var lastRow = tcSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: true, data: [] }, callback);
      var data = tcSheet.getRange(2, 1, lastRow - 1, 11).getValues();
      var headers = ['Task ID', 'Task Name', 'Label', 'Submenu', 'Page', 'Result Type', 'Options', 'CPO', 'CLNR', 'Sort Order', 'Active'];
      var rows = data.filter(function(row) {
        return row[0] && row[4].toString().trim() === page && row[10].toString().trim() === 'Yes';
      }).map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { obj[h] = row[i] !== null && row[i] !== undefined ? row[i].toString() : ''; });
        return obj;
      });
      // Sort by Sort Order
      rows.sort(function(a, b) { return (parseInt(a['Sort Order']) || 0) - (parseInt(b['Sort Order']) || 0); });
      return _respond({ success: true, data: rows }, callback);
    }

    // ---- SAVE TASK (add or update) ----
    if (action === 'savetask') {
      var tcSheet = ss.getSheetByName('Task Configuration');
      if (!tcSheet) {
        tcSheet = ss.insertSheet('Task Configuration');
        tcSheet.getRange(1, 1, 1, 11).setValues([['Task ID', 'Task Name', 'Label', 'Submenu', 'Page', 'Result Type', 'Options', 'CPO', 'CLNR', 'Sort Order', 'Active']]);
      }

      var taskId = (e.parameter.taskid || e.parameter.taskId || '').toString().trim();
      var taskName = (e.parameter.taskname || e.parameter.taskName || '').toString().trim();
      var label = (e.parameter.label || '').toString().trim();
      var submenu = (e.parameter.submenu || '').toString().trim();
      var page = (e.parameter.page || '').toString().trim();
      var resultType = (e.parameter.resulttype || e.parameter.resultType || '').toString().trim();
      var options = (e.parameter.options || '').toString().trim();
      var cpo = (e.parameter.cpo || 'Yes').toString().trim();
      var clnr = (e.parameter.clnr || 'Yes').toString().trim();
      var sortOrder = (e.parameter.sortorder || e.parameter.sortOrder || '0').toString().trim();
      var active = (e.parameter.active || 'Yes').toString().trim();

      if (!taskId) return _respond({ success: false, error: 'Task ID required' }, callback);
      if (!taskName) return _respond({ success: false, error: 'Task Name required' }, callback);
      if (!page) return _respond({ success: false, error: 'Page required' }, callback);
      if (!resultType) return _respond({ success: false, error: 'Result Type required' }, callback);

      // Check if task ID already exists
      var lastRow = tcSheet.getLastRow();
      var existingRow = -1;
      if (lastRow >= 2) {
        var ids = tcSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (ids[i][0].toString().trim() === taskId) { existingRow = i + 2; break; }
        }
      }

      var rowData = [taskId, taskName, label, submenu, page, resultType, options, cpo, clnr, sortOrder, active];

      if (existingRow > 0) {
        tcSheet.getRange(existingRow, 1, 1, 11).setValues([rowData]);
      } else {
        tcSheet.appendRow(rowData);
      }

      return _respond({ success: true, message: 'Task saved: ' + taskId }, callback);
    }

    // ---- DELETE TASK ----
    if (action === 'deletetask') {
      var tcSheet = ss.getSheetByName('Task Configuration');
      if (!tcSheet) return _respond({ success: false, error: 'No tasks configured' }, callback);
      var taskId = (e.parameter.taskid || e.parameter.taskId || '').toString().trim();
      if (!taskId) return _respond({ success: false, error: 'Task ID required' }, callback);

      var lastRow = tcSheet.getLastRow();
      if (lastRow < 2) return _respond({ success: false, error: 'Task not found' }, callback);
      var ids = tcSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (ids[i][0].toString().trim() === taskId) {
          tcSheet.deleteRow(i + 2);
          return _respond({ success: true, message: 'Task deleted' }, callback);
        }
      }
      return _respond({ success: false, error: 'Task not found' }, callback);
    }

    // ---- SEED ALL TASKS (one-time bulk insert) ----
    if (action === 'seedtasks') {
      var tcSheet = ss.getSheetByName('Task Configuration');
      if (!tcSheet) {
        tcSheet = ss.insertSheet('Task Configuration');
        tcSheet.getRange(1, 1, 1, 11).setValues([['Task ID', 'Task Name', 'Label', 'Submenu', 'Page', 'Result Type', 'Options', 'CPO', 'CLNR', 'Sort Order', 'Active']]);
      }

      var tasks = [
        // Unit Validation
        ['1', 'IMEI', 'IMEI', 'Unit Intake & Grading', 'Unit Validation', 'PASS/FAIL', '', 'Yes', 'Yes', '1', 'Yes'],
        ['2', 'OEM', 'OEM', 'Unit Intake & Grading', 'Unit Validation', 'PASS/FAIL', '', 'Yes', 'Yes', '2', 'Yes'],
        ['3', 'SKU', 'SKU', 'Unit Intake & Grading', 'Unit Validation', 'PASS/FAIL', '', 'Yes', 'Yes', '3', 'Yes'],
        ['4', 'MODEL', 'MODEL', 'Unit Intake & Grading', 'Unit Validation', 'PASS/FAIL', '', 'Yes', 'Yes', '4', 'Yes'],
        ['5', 'DISPO IN', 'DISPO IN', 'Unit Intake & Grading', 'Unit Validation', 'PASS/FAIL', '', 'Yes', 'Yes', '5', 'Yes'],

        // Initial Cosmetic Grading
        ['6', 'Wholesale COSMETIC GRADE IN (Aligned to Verizon Wholesale Cosmetic Definitions)', 'Wholesale Cosmetic Grade IN', 'Unit Intake & Grading', 'Initial Cosmetic Grading', 'DROPDOWN', 'AA,A,B,C,D,E', 'Yes', 'Yes', '1', 'Yes'],
        ['7', 'D2C GRADE IN (Aligned to Verizon D2C Cosmetic Definitions)', 'D2C Grade IN', 'Unit Intake & Grading', 'Initial Cosmetic Grading', 'DROPDOWN', 'A+,A,B,C,C-,D+,D', 'Yes', 'Yes', '2', 'Yes'],

        // Customer Information
        ['8', 'CI CLEAR STATUS (Pass / Fail)', 'CI Clear Status', 'Unit Intake & Grading', 'Customer Information', 'PASS/FAIL/NA', '', 'Yes', 'Yes', '1', 'Yes'],

        // Initial Functional Testing
        ['9', 'Is Device Beyond Economical Repair (BER)? (Y/N)', 'BER Status', 'Unit Intake & Grading', 'Initial Functional Testing', 'YES/NO', '', 'Yes', 'Yes', '1', 'Yes'],
        ['10', 'Was Functional Light Repair Attempted (Y/N)', 'Light Repair Attempted', 'Unit Intake & Grading', 'Initial Functional Testing', 'YES/NO', '', 'Yes', 'Yes', '2', 'Yes'],
        ['11', 'PRE-SERVICE FUNCTION TEST (Pass/Fail)', 'Pre-Service Function Test', 'Unit Intake & Grading', 'Initial Functional Testing', 'PASS/FAIL/NA', '', 'Yes', 'Yes', '3', 'Yes'],
        ['12', 'PRE-SERVICE RF TEST (Pass / Fail)', 'Pre-Service RF Test', 'Unit Intake & Grading', 'Initial Functional Testing', 'PASS/FAIL/NA', '', 'No', 'Yes', '4', 'Yes'],
        ['13', 'PRE-SERVICE AIR LEAK (PRESSURE CHAMBER) TEST (Pass / Fail)', 'Pre-Service Air Leak Test', 'Unit Intake & Grading', 'Initial Functional Testing', 'PASS/FAIL/NA', '', 'No', 'Yes', '5', 'Yes'],
        ['14', 'BSOH IN BATTERY STATE OF HEALTH', 'BSOH IN (Battery State of Health)', 'Unit Intake & Grading', 'Initial Functional Testing', 'PERCENTAGE', '', 'Yes', 'Yes', '6', 'Yes'],

        // Repair
        ['15', 'If device failed Pre-Service Function Test, what parts failed? Please list out all components which failed CPO Standards', 'Parts That Failed Pre-Service', 'Unit Services', 'Repair', 'FREE TEXT', '', 'Yes', 'Yes', '1', 'Yes'],
        ['16', 'Earpiece', 'Earpiece', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '2', 'Yes'],
        ['17', 'Microphone', 'Microphone', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '3', 'Yes'],
        ['18', 'Face ID', 'Face ID', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '4', 'Yes'],
        ['19', 'Speaker', 'Speaker', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '5', 'Yes'],
        ['20', 'Rear Camera', 'Rear Camera', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '6', 'Yes'],
        ['21', 'Front Camera', 'Front Camera', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '7', 'Yes'],
        ['22', 'Touch Screen', 'Touch Screen', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '8', 'Yes'],
        ['23', 'Ring Mute Switch', 'Ring Mute Switch', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '9', 'Yes'],
        ['24', 'Light Sensor', 'Light Sensor', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '10', 'Yes'],
        ['25', 'Camera Flash', 'Camera Flash', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '11', 'Yes'],
        ['26', 'Device Vibrate', 'Device Vibrate', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '12', 'Yes'],
        ['27', 'Battery', 'Battery', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '13', 'Yes'],
        ['28', 'RF', 'RF Repaired', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '14', 'Yes'],
        ['29', 'Comments on RF (e.g., Antenna, Proximity Sensor, etc.)', 'RF Comments', 'Unit Services', 'Repair', 'FREE TEXT', '', 'Yes', 'Yes', '15', 'Yes'],
        ['30', 'Other', 'Other Repaired', 'Unit Services', 'Repair', 'YES/NO', '', 'Yes', 'Yes', '16', 'Yes'],
        ['31', 'Comments on Other (e.g., charging port, etc.)', 'Other Comments', 'Unit Services', 'Repair', 'FREE TEXT', '', 'Yes', 'Yes', '17', 'Yes'],

        // Buff & Polish
        ['32', 'Was Unit Buff & Polished (Y/N)', 'Was Unit Buff & Polished?', 'Unit Services', 'Buff & Polish', 'YES/NO', '', 'Yes', 'Yes', '1', 'Yes'],
        ['33', '1 SIDE B&P (Indicate Y if Attempted)', '1 Side B&P Attempted', 'Unit Services', 'Buff & Polish', 'YES/NO/NA', '', 'Yes', 'Yes', '2', 'Yes'],
        ['34', '2 SIDE B&P (Indicate Y if Attempted)', '2 Side B&P Attempted', 'Unit Services', 'Buff & Polish', 'YES/NO/NA', '', 'Yes', 'Yes', '3', 'Yes'],
        ['35', 'B&P NOT ATTEMPTED', 'B&P Not Attempted', 'Unit Services', 'Buff & Polish', 'DROPDOWN', 'NOT ATTEMPTED,N/A', 'Yes', 'Yes', '4', 'Yes'],
        ['36', 'Reason for B&P No Attempt', 'Reason for B&P No Attempt', 'Unit Services', 'Buff & Polish', 'FREE TEXT', '', 'Yes', 'Yes', '5', 'Yes'],

        // Paint
        ['37', 'Paint Result', 'Paint Result', 'Unit Services', 'Paint', 'PASS/FAIL/NA', '', 'Yes', 'Yes', '1', 'Yes'],

        // Final Functional Testing
        ['38', 'POST-SERVICE FUNCTION TEST (Pass / Fail)', 'Post-Service Function Test', 'Final Results', 'Final Functional Testing', 'PASS/FAIL/NA', '', 'Yes', 'Yes', '1', 'Yes'],
        ['39', 'POST-SERVICE RF TEST (Pass / Fail)', 'Post-Service RF Test', 'Final Results', 'Final Functional Testing', 'PASS/FAIL/NA', '', 'No', 'Yes', '2', 'Yes'],
        ['40', 'POST-SERVICE AIR LEAK (PRESSURE CHAMBER) TEST (Pass / Fail)', 'Post-Service Air Leak Test', 'Final Results', 'Final Functional Testing', 'PASS/FAIL/NA', '', 'No', 'Yes', '3', 'Yes'],
        ['41', 'BSOH OUT BATTERY STATE OF HEALTH', 'BSOH OUT (Battery State of Health)', 'Final Results', 'Final Functional Testing', 'PERCENTAGE', '', 'Yes', 'Yes', '4', 'Yes'],
        ['42', 'VENDOR NOTES', 'Vendor Notes', 'Final Results', 'Final Functional Testing', 'FREE TEXT', '', 'Yes', 'Yes', '5', 'Yes'],

        // Final Cosmetic Grading
        ['43', 'Wholesale COSMETIC GRADE OUT (Aligned to Verizon Wholesale Cosmetic Definitions)', 'Wholesale Cosmetic Grade OUT', 'Final Results', 'Final Cosmetic Grading', 'DROPDOWN', 'AA,A,B,C,D,E', 'Yes', 'Yes', '1', 'Yes'],
        ['44', 'D2C COSMETIC GRADE OUT (Aligned to Verizon D2C Cosmetic Definitions)', 'D2C Cosmetic Grade OUT', 'Final Results', 'Final Cosmetic Grading', 'DROPDOWN', 'A+,A,B,C,C-,D+,D', 'Yes', 'Yes', '2', 'Yes'],

        // QC
        ['45', 'QC Result', 'QC Result', 'Final Results', 'QC', 'PASS/FAIL/NA', '', 'Yes', 'Yes', '1', 'Yes']
      ];

      tcSheet.getRange(2, 1, tasks.length, 11).setValues(tasks);
      return _respond({ success: true, message: tasks.length + ' tasks seeded' }, callback);
    }

    // ---- UNKNOWN ACTION ----
    return _respond({ success: false, error: 'Unknown action: ' + action }, callback);

  } catch (err) {
    return _respond({ success: false, error: err.message || err.toString() }, callback);
  }
}
