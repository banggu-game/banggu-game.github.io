// 방구 게임 공유 순위표 — 구글 시트 Apps Script
// 시트의 [확장 프로그램 → Apps Script]에 이 코드를 통째로 붙여넣고 웹 앱으로 배포하세요.

const SHEET_NAME = 'scores';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['game', 'id', 'name', 'score', 'time', 'date']);
  }
  return sh;
}

function clean_(s) { return String(s || '').replace(/[^\w-]/g, '').slice(0, 40); }

function name_(s) {
  s = String(s || '').replace(/[\r\n\t]/g, ' ').trim().slice(0, 8);
  if (!s) s = '진경';
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // 시트 수식으로 해석되지 않게
  return s;
}

function top_(game) {
  const sh = sheet_();
  if (sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter(r => r[0] === game && r[1] !== '' && Number(r[3]) > 0);
  rows.sort((a, b) => Number(b[3]) - Number(a[3]) || new Date(a[5]) - new Date(b[5]));
  return rows.slice(0, 3).map(r => ({
    id: String(r[1]),
    name: String(r[2]).replace(/^'/, ''),
    score: Number(r[3]),
    time: r[4] === '' ? null : Number(r[4])
  }));
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return out_({ ok: true, list: top_(clean_(e.parameter.game)) });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const game = clean_(d.game);
    const id = clean_(d.id);
    if (!game || !id) return out_({ ok: false });
    const sh = sheet_();
    const n = Math.max(sh.getLastRow() - 1, 0);
    const ids = n ? sh.getRange(2, 2, n, 1).getValues().map(r => String(r[0])) : [];
    const idx = ids.indexOf(id);

    if (d.action === 'add') {
      const score = Math.floor(Number(d.score));
      if (!(score > 0 && score < 100000)) return out_({ ok: false });
      const time = d.time == null || d.time === '' ? '' : Math.floor(Number(d.time));
      if (idx < 0) sh.appendRow([game, id, name_(d.name), score, time, new Date()]);
    } else if (d.action === 'rename') {
      if (idx >= 0) sh.getRange(idx + 2, 3).setValue(name_(d.name));
    }
    return out_({ ok: true, list: top_(game) });
  } finally {
    lock.releaseLock();
  }
}
