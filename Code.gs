/* ============================================================
   TÌM VỀ GÓC BÀN XƯA v2 — Google Apps Script Backend
   ============================================================
   HƯỚNG DẪN:
   1. Tạo Google Sheet mới (trống)
   2. Mở Extensions > Apps Script
   3. Xóa code mặc định, paste toàn bộ file này vào
   4. Bấm Deploy > New deployment
      - Type: Web app
      - Execute as: Me
      - Who has access: Anyone
   5. Copy URL deployment, paste vào game
   ============================================================ */

const EMPTY = '__E__';
const SKIP = '__SKIP__';

/* --- Routing --- */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  let result;
  try {
    switch (action) {
      case 'consensus':   result = getConsensus(); break;
      case 'leaderboard': result = getLeaderboard(); break;
      case 'names':       result = getNames(); break;
      default: result = { ok: true, msg: 'Tìm Về Góc Bàn Xưa API v2' };
    }
  } catch (err) { result = { error: err.toString() }; }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case 'submit': result = submitVotes(data); break;
      default: result = { error: 'Unknown action' };
    }
  } catch (err) { result = { error: err.toString() }; }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* --- Sheet helpers --- */

function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#2d6a4f').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

const VOTE_H   = ['Timestamp','SessionID','PlayerName','PlayerSeat','TargetSeat','Direction','VotedName'];
const PLAYER_H = ['Timestamp','SessionID','PlayerName','ClaimedSeat','Score','Bonus','Total'];
const SEAT_H   = ['SeatNumber','ConsensusName','VoteCount','TotalVotes','Confidence','AllVotes'];

/* --- Submit --- */

function submitVotes(data) {
  const vSh = getSheet('Votes', VOTE_H);
  const pSh = getSheet('Players', PLAYER_H);
  const now = new Date();
  const sid = data.sessionId || Utilities.getUuid();

  // Self-vote: player claims their own seat
  vSh.appendRow([now, sid, data.playerName, data.playerSeat,
                  data.playerSeat, 'SELF', data.playerName]);

  // Neighbor votes
  const votes = data.votes || [];
  votes.forEach(function(v) {
    if (v.votedName && v.votedName !== SKIP) {
      vSh.appendRow([now, sid, data.playerName, data.playerSeat,
                      v.targetSeat, v.direction, v.votedName]);
    }
  });

  // Rebuild consensus & score
  // Phương án B: chỉ tính ghế hợp lệ, quy đổi tỷ lệ × 18
  var consensus = rebuildConsensus();
  var rawScore = 0, rawMax = 0, bonus = 0;
  var details = [];
  var SCALE_MAX = 18;

  votes.forEach(function(v) {
    // Boundary positions (targetSeat is null): auto-correct, not scored
    if (!v.targetSeat) {
      details.push({ seat: null, dir: v.direction, voted: v.votedName,
                      matched: true, pts: 0, consensus: v.votedName,
                      confidence: 100, auto: true });
      return;
    }
    if (!v.votedName || v.votedName === SKIP) {
      details.push({ seat: v.targetSeat, dir: v.direction, voted: null,
                      matched: false, pts: 0, consensus: null, confidence: 0, auto: false });
      return;
    }
    var pts = v.points || 1;
    rawMax += pts;
    var c = consensus[v.targetSeat];
    var matched = c && c.name === v.votedName;
    var earned = matched ? pts : 0;
    rawScore += earned;
    if (matched && c.count === 2) bonus += 1;
    details.push({ seat: v.targetSeat, dir: v.direction, voted: v.votedName,
                    matched: matched, pts: earned,
                    consensus: c ? c.name : null,
                    confidence: c ? c.confidence : 0, auto: false });
  });

  // Scale: (rawScore / rawMax) × 18, rounded
  var scaledScore = rawMax > 0 ? Math.round(rawScore / rawMax * SCALE_MAX) : 0;
  var total = scaledScore + bonus;
  pSh.appendRow([now, sid, data.playerName, data.playerSeat, scaledScore, bonus, total]);

  return { success: true, score: scaledScore, bonus: bonus, total: total,
           maxScore: SCALE_MAX, rawScore: rawScore, rawMax: rawMax,
           details: details, consensus: consensus, sessionId: sid };
}

/* --- Consensus --- */

function rebuildConsensus() {
  var vSh = getSheet('Votes', VOTE_H);
  var cSh = getSheet('SeatMap', SEAT_H);
  var data = vSh.getDataRange().getValues();
  var seatVotes = {};

  for (var i = 1; i < data.length; i++) {
    var seat = data[i][4], name = data[i][6];
    if (!seat || !name) continue;
    if (!seatVotes[seat]) seatVotes[seat] = {};
    seatVotes[seat][name] = (seatVotes[seat][name] || 0) + 1;
  }

  cSh.clear();
  cSh.appendRow(SEAT_H);
  cSh.getRange(1, 1, 1, SEAT_H.length)
    .setFontWeight('bold').setBackground('#2d6a4f').setFontColor('#fff');
  cSh.setFrozenRows(1);

  var result = {};
  for (var s = 1; s <= 48; s++) {
    var votes = seatVotes[s] || {};
    var entries = Object.keys(votes).map(function(k){ return [k, votes[k]]; })
                  .sort(function(a,b){ return b[1]-a[1]; });
    var totalV = entries.reduce(function(sum, e){ return sum + e[1]; }, 0);
    var top = entries[0] || [null, 0];
    var conf = totalV > 0 ? Math.round(top[1] / totalV * 100) : 0;
    cSh.appendRow([s, top[0], top[1], totalV, conf, JSON.stringify(votes)]);
    result[s] = { name: top[0], count: top[1], total: totalV, confidence: conf, votes: votes };
  }
  return result;
}

/* --- Read endpoints --- */

function getConsensus() {
  var cSh = getSheet('SeatMap', SEAT_H);
  var data = cSh.getDataRange().getValues();
  var seats = {};
  for (var i = 1; i < data.length; i++) {
    seats[data[i][0]] = {
      name: data[i][1], count: data[i][2], total: data[i][3],
      confidence: data[i][4],
      votes: data[i][5] ? JSON.parse(data[i][5]) : {}
    };
  }
  var pSh = getSheet('Players', PLAYER_H);
  var pData = pSh.getDataRange().getValues();
  var players = {};
  for (var i = 1; i < pData.length; i++) players[pData[i][2]] = true;
  var filled = 0;
  for (var k in seats) if (seats[k].total > 0) filled++;

  return { seats: seats, totalPlayers: Object.keys(players).length,
           totalSessions: pData.length - 1, filledSeats: filled };
}

function getLeaderboard() {
  var pSh = getSheet('Players', PLAYER_H);
  var data = pSh.getDataRange().getValues();
  var entries = [];
  for (var i = 1; i < data.length; i++) {
    entries.push({
      time: Utilities.formatDate(new Date(data[i][0]),
            Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      name: data[i][2], seat: data[i][3],
      score: data[i][4], bonus: data[i][5], total: data[i][6]
    });
  }
  entries.sort(function(a,b){ return b.total - a.total; });
  return { entries: entries };
}

function getNames() {
  var vSh = getSheet('Votes', VOTE_H);
  var data = vSh.getDataRange().getValues();
  var names = {};
  for (var i = 1; i < data.length; i++) {
    var n = data[i][6];
    if (n && n !== EMPTY && n !== SKIP) names[n] = true;
  }
  var arr = Object.keys(names);
  arr.sort(function(a,b){ return a.localeCompare(b, 'vi'); });
  return { names: arr };
}

/* --- Init on open --- */
function onOpen() {
  getSheet('Votes', VOTE_H);
  getSheet('Players', PLAYER_H);
  getSheet('SeatMap', SEAT_H);
}
