/**
 * Google Apps Script — data sink for the Design Process Evaluation study.
 *
 * SETUP
 * 1. Open your Google Sheet → Extensions → Apps Script.
 * 2. Paste this whole file (replace anything there).
 * 3. Deploy → New deployment → type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Copy the /exec URL → that is your webhookUrl in index.html.
 * 4. Re-deploy (New version) whenever you change this file.
 *
 * LAYOUT: one ROW PER PHASE. Participant + project columns repeat on every
 * phase row so you can see who evaluated what, and how.
 */

var SHEET_NAME = 'responses';

// Column order for each phase row. Change here if you add survey questions.
var HEADERS = [
  // --- submission / participant ---
  'receivedAt', 'participantId', 'studyId', 'sessionId',
  'experience', 'domain', 'completionCode',
  'startTime', 'endTime', 'durationSeconds',
  // --- project ---
  'projectId', 'projectName', 'designer',
  // --- phase ---
  'phaseNumber', 'phaseName', 'actualCognitiveState',
  // --- A: diagnose the block ---
  'identifiedStates', 'primaryState', 'awareness', 'intervention',
  // --- B: believability ---
  'authenticity', 'believability', 'experienced', 'awkward',
  // --- C: process quality ---
  'plausibility', 'faithfulness',
  // --- D: confidence ---
  'confidence', 'infoSufficiency',
  // --- E: whole-project (only on last phase) ---
  'trajFlow', 'trajConsistency', 'trajText'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents);

    var sheet = getSheet_();

    // health-check ping
    if (body.ping) {
      sheet.appendRow(['PING', new Date().toISOString()]);
      return json_({ ok: true, ping: true });
    }

    var receivedAt = new Date().toISOString();
    var rows = [];
    var projects = body.projects || [];

    for (var pi = 0; pi < projects.length; pi++) {
      var proj = projects[pi];
      var phases = proj.phases || [];
      for (var hi = 0; hi < phases.length; hi++) {
        var ph = phases[hi];
        rows.push([
          receivedAt, body.participantId, body.studyId, body.sessionId,
          body.experience, body.domain, body.completionCode,
          body.startTime, body.endTime, body.durationSeconds,
          proj.projectId, proj.projectName, proj.designer,
          ph.phaseNumber, ph.phaseName, ph.actualCognitiveState,
          arr_(ph.identifiedStates), ph.primaryState, ph.awareness, ph.intervention,
          ph.authenticity, ph.believability, ph.experienced, ph.awkward,
          ph.plausibility, ph.faithfulness,
          ph.confidence, ph.infoSufficiency,
          ph.trajFlow, ph.trajConsistency, ph.trajText
        ]);
      }
    }

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    }

    return json_({ ok: true, rowsWritten: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// simple GET so you can open the /exec URL in a browser to check it's live
function doGet() {
  return json_({ ok: true, service: 'design-eval sink', sheet: SHEET_NAME });
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function arr_(v) {
  if (Array.isArray(v)) return v.join(', ');
  return v == null ? '' : v;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
