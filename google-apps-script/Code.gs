/**
 * World Cup 2026 Bracket Pool — Google Sheets backend
 *
 * SETUP (one time):
 * 1. Create a new Google Sheet
 * 2. Extensions → Apps Script → paste this file → Save
 * 3. Run setup() from the editor (authorize when prompted)
 * 4. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL into index.html → SHEETS_WEB_APP_URL
 *
 * Optional: set ADMIN_KEY below and the same value in index.html → SHEETS_ADMIN_KEY
 * so only the host can post match results.
 */

const SHEET_PLAYERS = 'Players';
const SHEET_PICKS = 'Picks';
const SHEET_RESULTS = 'Results';

// Change this and mirror it in index.html SHEETS_ADMIN_KEY (leave '' to allow anyone to enter results)
const ADMIN_KEY = '';

const MATCH_IDS = [
  'm49', 'm50', 'm51', 'm52', 'm53', 'm54', 'm55', 'm56',
  'm57', 'm58', 'm59', 'm60', 'm61', 'm62', 'm63', 'm64',
  'r16-1', 'r16-2', 'r16-3', 'r16-4', 'r16-5', 'r16-6', 'r16-7', 'r16-8',
  'qf-1', 'qf-2', 'qf-3', 'qf-4',
  'sf-1', 'sf-2',
  'final'
];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_PLAYERS, ['id', 'name']);
  ensureSheet_(ss, SHEET_PICKS, ['player_id', 'match_id', 'pick']);
  const results = ensureSheet_(ss, SHEET_RESULTS, ['match_id', 'winner']);

  const existing = readSheet_(results).map(function (r) { return r[0]; });
  MATCH_IDS.forEach(function (id) {
    if (existing.indexOf(id) === -1) {
      results.appendRow([id, '']);
    }
  });

  SpreadsheetApp.flush();
  Logger.log('Setup complete. Deploy as web app next.');
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'getAll';
    if (action === 'getAll') {
      return jsonOut_(getAll_());
    }
    return jsonOut_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    switch (action) {
      case 'addPlayer':
        return jsonOut_(addPlayer_(body.name));
      case 'removePlayer':
        return jsonOut_(removePlayer_(body.playerId));
      case 'savePick':
        return jsonOut_(savePick_(body.playerId, body.matchId, body.pick, body.clearMatchIds || []));
      case 'saveResult':
        requireAdmin_(body.adminKey);
        return jsonOut_(saveResult_(body.matchId, body.winner, body.clearMatchIds || []));
      case 'clearResults':
        requireAdmin_(body.adminKey);
        return jsonOut_(clearResults_());
      default:
        return jsonOut_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err.message || err) });
  }
}

function getAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var playersRows = readSheet_(ss.getSheetByName(SHEET_PLAYERS));
  var picksRows = readSheet_(ss.getSheetByName(SHEET_PICKS));
  var resultsRows = readSheet_(ss.getSheetByName(SHEET_RESULTS));

  var players = playersRows.map(function (r) {
    return { id: String(r[0]), name: String(r[1]) };
  });

  var picks = {};
  picksRows.forEach(function (r) {
    var pid = String(r[0]);
    var mid = String(r[1]);
    var pick = String(r[2]);
    if (!pick) return;
    if (!picks[pid]) picks[pid] = {};
    picks[pid][mid] = pick;
  });

  var results = {};
  resultsRows.forEach(function (r) {
    var mid = String(r[0]);
    var winner = r[1] ? String(r[1]) : '';
    if (winner) results[mid] = winner;
  });

  return {
    ok: true,
    players: players,
    picks: picks,
    results: results,
    sheetUrl: ss.getUrl(),
    updatedAt: Date.now()
  };
}

function addPlayer_(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('Name required');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PLAYERS);
  var id = 'p_' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  sheet.appendRow([id, name]);
  return { ok: true, player: { id: id, name: name }, updatedAt: Date.now() };
}

function removePlayer_(playerId) {
  playerId = String(playerId);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  deleteRowsWhere_(ss.getSheetByName(SHEET_PLAYERS), 0, playerId);
  deleteRowsWhere_(ss.getSheetByName(SHEET_PICKS), 0, playerId);
  return { ok: true, updatedAt: Date.now() };
}

function savePick_(playerId, matchId, pick, clearMatchIds) {
  playerId = String(playerId);
  matchId = String(matchId);
  pick = pick ? String(pick) : '';

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PICKS);
  upsertPick_(sheet, playerId, matchId, pick);

  (clearMatchIds || []).forEach(function (mid) {
    upsertPick_(sheet, playerId, String(mid), '');
  });

  return { ok: true, updatedAt: Date.now() };
}

function saveResult_(matchId, winner, clearMatchIds) {
  matchId = String(matchId);
  winner = winner ? String(winner) : '';

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
  upsertResult_(sheet, matchId, winner);

  (clearMatchIds || []).forEach(function (mid) {
    upsertResult_(sheet, String(mid), '');
  });

  return { ok: true, updatedAt: Date.now() };
}

function clearResults_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    sheet.getRange(i + 1, 2).setValue('');
  }
  return { ok: true, updatedAt: Date.now() };
}

function requireAdmin_(key) {
  if (!ADMIN_KEY) return;
  if (String(key || '') !== ADMIN_KEY) {
    throw new Error('Admin key required to change results');
  }
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function readSheet_(sheet) {
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).filter(function (r) { return r[0] !== '' && r[0] != null; });
}

function upsertPick_(sheet, playerId, matchId, pick) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === playerId && String(data[i][1]) === matchId) {
      if (pick) {
        sheet.getRange(i + 1, 3).setValue(pick);
      } else {
        sheet.deleteRow(i + 1);
      }
      return;
    }
  }
  if (pick) sheet.appendRow([playerId, matchId, pick]);
}

function upsertResult_(sheet, matchId, winner) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === matchId) {
      sheet.getRange(i + 1, 2).setValue(winner);
      return;
    }
  }
  sheet.appendRow([matchId, winner]);
}

function deleteRowsWhere_(sheet, colIndex, value) {
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIndex]) === String(value)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
