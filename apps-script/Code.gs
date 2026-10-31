/**
 * 궁금해 검색 — 노래 목록 저장 + 유튜브 검색
 *
 * 설치 (한 번만)
 * 1. 새 구글 시트를 만들고  확장 프로그램 → Apps Script
 * 2. 이 파일 내용을 그대로 붙여 넣기
 * 3. 왼쪽 '서비스' 옆 + → YouTube Data API v3 → 추가
 * 4. 아래 PIN에 어른용 비밀번호를 적기 (비우면 누구나 넣고 뺄 수 있어요)
 * 5. 배포 → 새 배포 → 유형: 웹 앱
 *    실행 계정: 나 / 액세스 권한: 모든 사용자 → 배포 → 권한 허용
 * 6. 나온 웹 앱 주소(https://script.google.com/macros/s/.../exec)를 알려 주기
 */
const PIN = '';
const SHEET = '노래';

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange('A:F').setNumberFormat('@');
    sh.appendRow(['key', 'id', 'title', 'keys', 'start', 'end', 'added']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function list_() {
  const v = sheet_().getDataRange().getValues();
  v.shift();
  return v.filter(r => r[0]).map(r => ({
    key: String(r[0]), id: String(r[1]), title: String(r[2]), keys: String(r[3]),
    start: Number(r[4]) || 0, end: Number(r[5]) || 0,
  }));
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 시트에서 수식으로 읽히지 않게
const safe_ = (s, n) => {
  s = String(s || '').slice(0, n);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'search') return out_({ ok: true, items: search_(p.q || '') });
    if (p.action === 'check') return out_({ ok: true, needPin: !!PIN, valid: !PIN || p.pin === PIN });
    return out_({ ok: true, songs: list_() });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (_) {}
  if (PIN && b.pin !== PIN) return out_({ ok: false, error: 'PIN' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const have = new Set(list_().map(x => x.key));
    const add = s => {
      s = s || {};
      if (!/^[\w-]{11}$/.test(s.id)) return 'BAD_ID';
      const start = Math.max(0, Math.floor(Number(s.start) || 0));
      const end = Math.max(0, Math.floor(Number(s.end) || 0));
      const key = s.id + '-' + start + '-' + end;
      if (have.has(key)) return 'DUP';
      sh.appendRow([key, s.id, safe_(s.title, 80), safe_(s.keys, 200), String(start), String(end), new Date()]);
      have.add(key);
      return '';
    };
    if (b.action === 'add') {
      const err = add(b.song);
      if (err) return out_({ ok: false, error: err, songs: list_() });
    } else if (b.action === 'addMany') {
      (b.songs || []).slice(0, 200).forEach(add);
    } else if (b.action === 'delete') {
      const vals = sh.getDataRange().getValues();
      for (let i = vals.length - 1; i >= 1; i--) {
        if (String(vals[i][0]) === String(b.key)) sh.deleteRow(i + 1);
      }
    } else {
      return out_({ ok: false, error: 'BAD_ACTION' });
    }
    return out_({ ok: true, songs: list_() });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function search_(q) {
  q = String(q).trim();
  if (!q) return [];
  const r = YouTube.Search.list('snippet', {
    q: q, type: 'video', maxResults: 12, safeSearch: 'strict',
    regionCode: 'KR', relevanceLanguage: 'ko', videoEmbeddable: 'true',
  });
  const items = r.items || [];
  const ids = items.map(i => i.id.videoId);
  const dur = {};
  if (ids.length) {
    (YouTube.Videos.list('contentDetails', { id: ids.join(',') }).items || [])
      .forEach(v => { dur[v.id] = iso_(v.contentDetails.duration); });
  }
  return items.map(i => ({
    id: i.id.videoId, title: i.snippet.title, channel: i.snippet.channelTitle, duration: dur[i.id.videoId] || 0,
  }));
}

function iso_(s) {
  const m = String(s).match(/P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  return m ? (+m[1] || 0) * 86400 + (+m[2] || 0) * 3600 + (+m[3] || 0) * 60 + (+m[4] || 0) : 0;
}
