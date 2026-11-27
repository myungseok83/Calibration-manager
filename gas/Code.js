/**
 * ===============================================================
 * 측정기 교정관리 시스템 - Google Apps Script 백엔드
 * VERSION: V1.0
 * - [최초] 교정대상 가져오기(엑셀 → 웹앱 → 시트), 입고/교정/출하 단계 처리(일괄 포함), 단계 되돌리기
 * - [최초] 로그인(비밀번호 SHA-256 저장) · 권한 3단계(관리자/담당자/뷰어) · 첫 가입자는 관리자
 * - [최초] 🔑 접속 비밀번호: 스크립트 속성 ACCESS_CODE를 넣으면 비밀번호를 모르는 요청은 모두 거절
 * - [최초] 처리 이력('교정이력' 탭), 새 회차 가져오기 시 이전 회차 기록을 '교정완료기록' 탭에 보관
 * - [최초] 저장 동시 실행 잠금(LockService) + 같은 요청번호(_reqId) 10분간 중복 실행 방지
 * ===============================================================
 * 연결 시트: 이 스크립트가 붙어 있는 스프레드시트(확장 프로그램 → Apps Script)
 * 탭은 처음 실행할 때 자동으로 만들어집니다.
 */
const GAS_VERSION = 'V1.0';

const SHEET_EQ = '교정대상';
const SHEET_HIST = '교정이력';
const SHEET_ARCH = '교정완료기록';
const SHEET_USERS = '사용자';
const SHEET_META = '설정';

// 교정대상 탭 열 (순서 = 시트 열 순서)
const EQ_FIELDS = [
  'id', 'name', 'unit', 'cycle', 'lastCal', 'nextDue', 'agency', 'serial', 'location', 'owner', 'dept', 'grr',
  'method', 'round',
  'inDate', 'inGiver', 'inTaker', 'appearance', 'accessory',
  'calDate', 'result', 'certNo', 'newDue', 'label',
  'outDate', 'outReceiver',
  'note', 'active', 'updatedAt', 'updatedBy'
];
const EQ_HEADERS = [
  '측정기 번호', '측정기 이름', '측정단위', '교정주기(개월)', '교정일자(직전)', '차기교정일자', '교정기관', 'SERIAL NO.', '사용처', '담당자', '사용부서', 'GR&R',
  '교정방식', '회차',
  '입고일', '인계자', '인수자', '외관·작동', '부속품',
  '실시교정일', '교정결과', '성적서번호', '신규 차기교정일', '라벨부착',
  '출하일', '수령자',
  '비고', '사용', '수정일시', '수정자'
];
const MASTER_FIELDS = ['name', 'unit', 'cycle', 'lastCal', 'nextDue', 'agency', 'serial', 'location', 'owner', 'dept', 'grr'];
const STEP_FIELDS = {
  in: ['inDate', 'inGiver', 'inTaker', 'appearance', 'accessory'],
  cal: ['calDate', 'result', 'certNo', 'newDue', 'label'],
  out: ['outDate', 'outReceiver']
};
const EDITABLE = [].concat(STEP_FIELDS.in, STEP_FIELDS.cal, STEP_FIELDS.out, ['method', 'note']);
const ROLE_LEVEL = { viewer: 1, editor: 2, admin: 3 };
const FIELD_LABEL = {}; EQ_FIELDS.forEach(function (f, i) { FIELD_LABEL[f] = EQ_HEADERS[i]; });

/* ---------------- 공통 ---------------- */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function getSS() { return SpreadsheetApp.getActiveSpreadsheet(); }
function props() { return PropertiesService.getScriptProperties(); }
function nowStr() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }

function getOrCreateSheet(name, headers) {
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#dbe4f3');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@'); // 날짜가 자동 변환되지 않게 글자로 저장
  }
  return sh;
}
function eqSheet() { return getOrCreateSheet(SHEET_EQ, EQ_HEADERS); }
function histSheet() { return getOrCreateSheet(SHEET_HIST, ['일시', '처리자', '측정기 번호', '구분', '내용', '회차']); }
function archSheet() {
  return getOrCreateSheet(SHEET_ARCH, ['보관일시', '회차', '측정기 번호', '측정기 이름', 'SERIAL NO.', '사용부서', '담당자',
    '입고일', '실시교정일', '교정결과', '성적서번호', '신규 차기교정일', '출하일', '비고']);
}
function usersSheet() { return getOrCreateSheet(SHEET_USERS, ['아이디', '비밀번호(암호화)', '이름', '권한', '가입일시']); }
function metaSheet() { return getOrCreateSheet(SHEET_META, ['항목', '값']); }

function getMeta(key) {
  const v = metaSheet().getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === key) return String(v[i][1]);
  return '';
}
function setMeta(key, value) {
  const sh = metaSheet(); const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  sh.appendRow([key, value]);
}

function cellStr(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  }
  return String(v);
}

function readEq() {
  const sh = eqSheet();
  const last = sh.getLastRow();
  if (last < 2) return { sh: sh, rows: [] };
  const vals = sh.getRange(2, 1, last - 1, EQ_FIELDS.length).getValues();
  const rows = vals.map(function (r) {
    const o = {}; EQ_FIELDS.forEach(function (f, i) { o[f] = cellStr(r[i]); }); return o;
  }).filter(function (o) { return o.id; });
  return { sh: sh, rows: rows };
}
function writeEq(sh, rows) {
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, EQ_FIELDS.length).clearContent();
  if (!rows.length) return;
  const out = rows.map(function (o) { return EQ_FIELDS.map(function (f) { return o[f] === undefined ? '' : String(o[f]); }); });
  const rng = sh.getRange(2, 1, out.length, EQ_FIELDS.length);
  rng.setNumberFormat('@');
  rng.setValues(out);
}

function addMonths(ymd, months) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || ''));
  const n = parseInt(months, 10);
  if (!m || !n) return '';
  const y = +m[1], mo = +m[2] - 1 + n, d = +m[3];
  const last = new Date(y, mo + 1, 0).getDate(); // 말일 보정(1/31 + 1개월 = 2/28)
  const dt = new Date(y, mo, Math.min(d, last));
  return Utilities.formatDate(dt, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
}
function isPass(result) { return result === '합격' || result === '조건부합격'; }

function addHistory(rows) { // rows: [[user, id, kind, detail, round]]
  if (!rows.length) return;
  const t = nowStr();
  const sh = histSheet();
  const out = rows.map(function (r) { return [t].concat(r); });
  sh.getRange(sh.getLastRow() + 1, 1, out.length, 6).setNumberFormat('@').setValues(out);
}

/* ---------------- 접속 비밀번호 / 로그인 ---------------- */
function accessCodeValue() { try { return String(props().getProperty('ACCESS_CODE') || '').trim(); } catch (e) { return ''; } }
function accessOk(k) { const c = accessCodeValue(); return !c || String(k || '').trim() === c; }

function sha256(s) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 0xff).toString(16)).slice(-2); }).join('');
}
function tokenSecret() {
  let s = props().getProperty('TOKEN_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props().setProperty('TOKEN_SECRET', s); }
  return s;
}
function makeToken(username) {
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(username), tokenSecret()));
  return username + '|' + sig;
}
function findUser(username) {
  const v = usersSheet().getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() === String(username).toLowerCase()) {
      return { row: i + 1, username: String(v[i][0]), pw: String(v[i][1]), name: String(v[i][2]), role: String(v[i][3]) || 'viewer' };
    }
  }
  return null;
}
// 토큰 확인 → 사용자(권한은 매번 시트에서 다시 읽음: 권한 변경 즉시 반영)
function authUser(token) {
  const t = String(token || ''); const p = t.lastIndexOf('|');
  if (p < 1) return null;
  const username = t.slice(0, p);
  if (makeToken(username) !== t) return null;
  return findUser(username);
}
function need(body, role) {
  const u = authUser(body._token);
  if (!u) return { err: { status: 'fail', authFail: true, message: '다시 로그인해주세요.' } };
  if ((ROLE_LEVEL[u.role] || 0) < ROLE_LEVEL[role]) return { err: { status: 'fail', message: '권한이 없습니다. (' + role + ' 이상 필요)' } };
  return { user: u };
}
function userView(u) { return { username: u.username, name: u.name, role: u.role }; }

function handleLogin(body) {
  const u = findUser(body.username);
  if (!u || u.pw !== sha256(body.password)) return { status: 'fail', message: '아이디 또는 비밀번호가 맞지 않습니다.' };
  return { status: 'success', user: userView(u), token: makeToken(u.username) };
}
function handleRegister(body) {
  const username = String(body.username || '').trim(), name = String(body.name || '').trim();
  if (!/^[A-Za-z0-9_.-]{2,30}$/.test(username)) return { status: 'fail', message: '아이디는 영문/숫자 2~30자로 입력해주세요.' };
  if (String(body.password || '').length < 4) return { status: 'fail', message: '비밀번호는 4자 이상으로 입력해주세요.' };
  if (!name) return { status: 'fail', message: '이름을 입력해주세요.' };
  if (findUser(username)) return { status: 'fail', message: '이미 있는 아이디입니다.' };
  const sh = usersSheet();
  const role = sh.getLastRow() < 2 ? 'admin' : 'viewer'; // 첫 가입자 = 관리자
  sh.appendRow([username, sha256(body.password), name, role, nowStr()]);
  const u = findUser(username);
  return { status: 'success', user: userView(u), token: makeToken(username), firstAdmin: role === 'admin' };
}
function handleChangePassword(body) {
  const a = need(body, 'viewer'); if (a.err) return a.err;
  if (a.user.pw !== sha256(body.oldPassword)) return { status: 'fail', message: '현재 비밀번호가 맞지 않습니다.' };
  if (String(body.newPassword || '').length < 4) return { status: 'fail', message: '새 비밀번호는 4자 이상' };
  usersSheet().getRange(a.user.row, 2).setValue(sha256(body.newPassword));
  return { status: 'success' };
}
function handleListUsers(body) {
  const a = need(body, 'admin'); if (a.err) return a.err;
  const v = usersSheet().getDataRange().getValues();
  const list = [];
  for (let i = 1; i < v.length; i++) list.push({ username: String(v[i][0]), name: String(v[i][2]), role: String(v[i][3]), createdAt: cellStr(v[i][4]) });
  return { status: 'success', data: list };
}
function handleUpdateUser(body) {
  const a = need(body, 'admin'); if (a.err) return a.err;
  const u = findUser(body.username); if (!u) return { status: 'fail', message: '사용자 없음' };
  if (body.remove) {
    if (u.username === a.user.username) return { status: 'fail', message: '본인 계정은 삭제할 수 없습니다.' };
    usersSheet().deleteRow(u.row); return { status: 'success' };
  }
  if (body.resetPassword) { usersSheet().getRange(u.row, 2).setValue(sha256(body.resetPassword)); return { status: 'success' }; }
  if (!ROLE_LEVEL[body.role]) return { status: 'fail', message: '권한 값 오류' };
  if (u.username === a.user.username && body.role !== 'admin') return { status: 'fail', message: '본인의 관리자 권한은 해제할 수 없습니다.' };
  usersSheet().getRange(u.row, 4).setValue(body.role);
  return { status: 'success' };
}

/* ---------------- 교정대상 가져오기 ---------------- */
// body.round: '2026-11' 같은 회차 이름, body.rows: [{id, name, ... , method}]
function handleImportPlan(body) {
  const a = need(body, 'admin'); if (a.err) return a.err;
  const round = String(body.round || '').trim();
  if (!round) return { status: 'fail', message: '회차 이름이 필요합니다.' };
  const incoming = body.rows || [];
  const r = readEq(); const map = {};
  r.rows.forEach(function (o, i) { map[o.id] = i; });
  const t = nowStr(); const who = a.user.name;
  let added = 0, updated = 0, archived = 0;
  const arch = [];
  incoming.forEach(function (src) {
    const id = String(src.id || '').trim(); if (!id) return;
    let o;
    if (map[id] !== undefined) {
      o = r.rows[map[id]];
      const hasWork = o.inDate || o.calDate || o.outDate;
      if (hasWork && o.round !== round) { // 이전 회차 기록 보관 후 진행 칸 비우기
        arch.push([t, o.round, o.id, o.name, o.serial, o.dept, o.owner, o.inDate, o.calDate, o.result, o.certNo, o.newDue, o.outDate, o.note]);
        [].concat(STEP_FIELDS.in, STEP_FIELDS.cal, STEP_FIELDS.out).forEach(function (f) { o[f] = ''; });
        archived++;
      }
      updated++;
    } else {
      o = {}; EQ_FIELDS.forEach(function (f) { o[f] = ''; }); o.id = id;
      r.rows.push(o); map[id] = r.rows.length - 1; added++;
    }
    MASTER_FIELDS.forEach(function (f) { if (src[f] !== undefined && src[f] !== null) o[f] = String(src[f]); });
    if (!o.method) o.method = src.method === '현장' ? '현장' : '입고';
    o.round = round; o.active = 'Y'; o.updatedAt = t; o.updatedBy = who;
  });
  if (body.deactivateMissing) {
    const keep = {}; incoming.forEach(function (s) { keep[String(s.id || '').trim()] = 1; });
    r.rows.forEach(function (o) { if (!keep[o.id]) o.active = 'N'; });
  }
  writeEq(r.sh, r.rows);
  if (arch.length) { const sh = archSheet(); sh.getRange(sh.getLastRow() + 1, 1, arch.length, arch[0].length).setNumberFormat('@').setValues(arch); }
  setMeta('currentRound', round);
  addHistory([[who, '-', '가져오기', round + ' 회차: 신규 ' + added + ', 갱신 ' + updated + ', 이전기록 보관 ' + archived, round]]);
  return { status: 'success', added: added, updated: updated, archived: archived, data: readAllData() };
}

/* ---------------- 단계 처리 ---------------- */
// body: { ids:[], kind:'입고'|'교정'|'출하'|'수정', patch:{필드:값} }  (한 번에 여러 대 처리 가능)
function handleUpdateItems(body) {
  const a = need(body, 'editor'); if (a.err) return a.err;
  const ids = (body.ids || []).map(String); const patch = body.patch || {};
  const keys = Object.keys(patch).filter(function (k) { return EDITABLE.indexOf(k) >= 0; });
  if (!ids.length || !keys.length) return { status: 'fail', message: '변경할 내용이 없습니다.' };
  if (keys.indexOf('method') >= 0 && a.user.role !== 'admin') return { status: 'fail', message: '교정방식 변경은 관리자만 가능합니다.' };
  const r = readEq(); const t = nowStr(); const who = a.user.name; const hist = []; const changed = [];
  const idset = {}; ids.forEach(function (x) { idset[x] = 1; });
  r.rows.forEach(function (o) {
    if (!idset[o.id]) return;
    keys.forEach(function (k) { o[k] = patch[k] === null || patch[k] === undefined ? '' : String(patch[k]); });
    if (keys.indexOf('calDate') >= 0 || keys.indexOf('result') >= 0) o.newDue = isPass(o.result) ? addMonths(o.calDate, o.cycle) : '';
    o.updatedAt = t; o.updatedBy = who;
    hist.push([who, o.id, body.kind || '수정', keys.map(function (k) { return FIELD_LABEL[k] + ' ' + (o[k] || '-'); }).join(', '), o.round]);
    changed.push(o);
  });
  if (!changed.length) return { status: 'fail', message: '대상 측정기를 찾지 못했습니다.' };
  writeEq(r.sh, r.rows); addHistory(hist);
  return { status: 'success', items: changed };
}
// 단계 되돌리기(관리자): step = 'in' | 'cal' | 'out'
function handleUndoStep(body) {
  const a = need(body, 'admin'); if (a.err) return a.err;
  const fields = STEP_FIELDS[body.step]; if (!fields) return { status: 'fail', message: '단계 값 오류' };
  const r = readEq(); const o = r.rows.filter(function (x) { return x.id === String(body.id); })[0];
  if (!o) return { status: 'fail', message: '측정기 없음' };
  fields.forEach(function (f) { o[f] = ''; }); o.updatedAt = nowStr(); o.updatedBy = a.user.name;
  writeEq(r.sh, r.rows);
  addHistory([[a.user.name, o.id, '되돌리기', ({ in: '입고', cal: '교정', out: '출하' })[body.step] + ' 단계 취소' + (body.reason ? ' — ' + body.reason : ''), o.round]]);
  return { status: 'success', items: [o] };
}
function handleDeleteItem(body) {
  const a = need(body, 'admin'); if (a.err) return a.err;
  const r = readEq(); const before = r.rows.length;
  r.rows = r.rows.filter(function (x) { return x.id !== String(body.id); });
  if (r.rows.length === before) return { status: 'fail', message: '측정기 없음' };
  writeEq(r.sh, r.rows);
  addHistory([[a.user.name, String(body.id), '삭제', body.reason || '', '']]);
  return { status: 'success' };
}

/* ---------------- 조회 ---------------- */
function readAllData() {
  return { items: readEq().rows, round: getMeta('currentRound'), version: GAS_VERSION };
}
function getHistory(id) {
  const v = histSheet().getDataRange().getValues(); const out = [];
  for (let i = v.length - 1; i >= 1 && out.length < 200; i--) {
    if (!id || String(v[i][2]) === String(id)) out.push({ at: cellStr(v[i][0]), by: String(v[i][1]), id: String(v[i][2]), kind: String(v[i][3]), detail: String(v[i][4]), round: String(v[i][5]) });
  }
  return out;
}
function getArchive(id) {
  const v = archSheet().getDataRange().getValues(); const out = [];
  for (let i = 1; i < v.length; i++) if (String(v[i][2]) === String(id)) {
    out.push({ round: String(v[i][1]), inDate: cellStr(v[i][7]), calDate: cellStr(v[i][8]), result: String(v[i][9]), certNo: String(v[i][10]), newDue: cellStr(v[i][11]), outDate: cellStr(v[i][12]) });
  }
  return out;
}

/* ---------------- 진입점 ---------------- */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = p.action || 'read';
    if (action === 'version') return jsonResponse({ status: 'success', data: GAS_VERSION });
    if (action === 'checkAccess') return jsonResponse({ status: 'success', required: !!accessCodeValue(), ok: accessOk(p.ak), version: GAS_VERSION });
    if (!accessOk(p.ak)) return jsonResponse({ status: 'fail', accessDenied: true, message: '접속 비밀번호가 필요합니다.' });
    if (action === 'read') return jsonResponse({ status: 'success', data: readAllData() });
    if (action === 'getHistory') return jsonResponse({ status: 'success', data: getHistory(p.id), archive: p.id ? getArchive(p.id) : [] });
    return jsonResponse({ status: 'fail', message: 'Unknown GET action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  let lock = null;
  try {
    const action = e.parameter.action;
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (!accessOk(e.parameter.ak || body._ak)) return jsonResponse({ status: 'fail', accessDenied: true, message: '접속 비밀번호가 필요합니다.' });
    lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) return jsonResponse({ status: 'error', busy: true, message: '다른 저장을 처리 중입니다. 잠시 후 다시 시도해주세요.' });
    const reqId = body._reqId ? String(body._reqId) : '';
    const cache = CacheService.getScriptCache();
    if (reqId) { const c = cache.get('req_' + reqId); if (c) return ContentService.createTextOutput(c).setMimeType(ContentService.MimeType.JSON); }
    const result = routePost(action, body);
    const text = JSON.stringify(result);
    if (reqId && text.length < 90000) cache.put('req_' + reqId, text, 600);
    return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e2) {} }
  }
}

function routePost(action, body) {
  switch (action) {
    case 'login': return handleLogin(body);
    case 'register': return handleRegister(body);
    case 'changePassword': return handleChangePassword(body);
    case 'listUsers': return handleListUsers(body);
    case 'updateUser': return handleUpdateUser(body);
    case 'importPlan': return handleImportPlan(body);
    case 'updateItems': return handleUpdateItems(body);
    case 'undoStep': return handleUndoStep(body);
    case 'deleteItem': return handleDeleteItem(body);
    default: return { status: 'fail', message: 'Unknown POST action: ' + action };
  }
}

// 편집기에서 한 번 실행하면 탭을 미리 만들어 둠(선택)
function setup() { eqSheet(); histSheet(); archSheet(); usersSheet(); metaSheet(); tokenSecret(); }
