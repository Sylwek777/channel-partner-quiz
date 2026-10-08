/**
 * Backend for the Channel Partner Playbook Quiz.
 * Deploy as a Google Apps Script Web App ("Anyone" access) bound to a
 * Google Sheet. Respondents never see a login screen: this script runs
 * under the deployer's identity, not the visitor's.
 */

var SHEET_NAME = 'Responses';
var QUESTION_COUNT = 6;

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var header = ['timestamp', 'name', 'score'];
    for (var i = 1; i <= QUESTION_COUNT; i++) header.push('q' + i);
    sheet.appendRow(header);
  }
  return sheet;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = JSON.parse(e.postData.contents);
    var answers = Array.isArray(data.answers) ? data.answers : [];
    var sheet = getSheet();
    var row = [new Date(), String(data.name || '').slice(0, 40), Number(data.score) || 0];
    for (var i = 0; i < QUESTION_COUNT; i++) {
      row.push(answers[i] !== undefined ? answers[i] : '');
    }
    sheet.appendRow(row);
    return jsonOutput({ ok: true });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var action = (e.parameter && e.parameter.action) || 'ping';
  var sheet = getSheet();
  var range = sheet.getDataRange().getValues();
  var rows = range.slice(1); // drop header row

  if (action === 'count') {
    return jsonOutput({ total: rows.length });
  }

  if (action === 'results') {
    var perQuestion = [];
    for (var i = 0; i < QUESTION_COUNT; i++) perQuestion.push([0, 0, 0, 0]);
    var scoreHist = new Array(QUESTION_COUNT + 1).fill(0);
    var sum = 0, top = 0;

    rows.forEach(function (r) {
      var score = Number(r[2]) || 0;
      sum += score;
      if (score > top) top = score;
      if (scoreHist[score] !== undefined) scoreHist[score]++;
      for (var i = 0; i < QUESTION_COUNT; i++) {
        var a = r[3 + i];
        if (a === 0 || a === 1 || a === 2 || a === 3) {
          perQuestion[i][a]++;
        }
      }
    });

    return jsonOutput({
      total: rows.length,
      avg: rows.length ? sum / rows.length : 0,
      top: top,
      perQuestion: perQuestion,
      scoreHist: scoreHist
    });
  }

  return jsonOutput({ ok: true, message: 'Channel Partner Quiz API' });
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
