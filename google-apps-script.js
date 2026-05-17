// ============================================================
// GOOGLE APPS SCRIPT — paste this into Extensions > Apps Script
// in your Google Sheet, then deploy as a NEW version of the web app.
// ============================================================

// WRITE SECRET — must match the one generated from your PIN on the client.
// This is a SHA-256 hash of your PIN + a salt. Only someone who knows the PIN
// can generate this token. Change this if you change your PIN.
// Default: SHA-256("1234:soccer-eval-write") 
var WRITE_TOKEN = '2b42e19e7ba32b70ad55b64896f15d2880af066fa6193844cb38c4c7e52ea2e0';

// Returns all assessments as JSON (public read — no auth needed)
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessments');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Assessments');
    sheet.appendRow(['id', 'team', 'playerName', 'position', 'date', 'ratings', 'notes']);
  }

  var action = e.parameter.action || 'getAll';
  var team = e.parameter.team || '';

  if (action === 'getAll') {
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      if (!team || row.team === team) {
        try { row.ratings = JSON.parse(row.ratings); } catch(err) { row.ratings = {}; }
        try { row.notes = JSON.parse(row.notes); } catch(err) { row.notes = []; }
        rows.push(row);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getTeams') {
    var data = sheet.getDataRange().getValues();
    var teams = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][1]) teams[data[i][1]] = true;
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true, teams: Object.keys(teams) }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Saves or deletes an assessment (requires write token)
function doPost(e) {
  var body = JSON.parse(e.postData.contents);

  // ===== AUTH CHECK =====
  if (!body.token || body.token !== WRITE_TOKEN) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assessments');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Assessments');
    sheet.appendRow(['id', 'team', 'playerName', 'position', 'date', 'ratings', 'notes']);
  }

  var action = body.action || 'save';

  if (action === 'save') {
    var id = body.id || Utilities.getUuid();
    var row = [
      id,
      body.team || '',
      body.playerName || '',
      body.position || '',
      body.date || new Date().toISOString(),
      JSON.stringify(body.ratings || {}),
      JSON.stringify(body.notes || [])
    ];

    var data = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow(row);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, id: id }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'delete') {
    var id = body.id;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}
