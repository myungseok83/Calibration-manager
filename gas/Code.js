/**
 * ===============================================================
 * 측정기 교정관리 솔루션 - Google Apps Script 백엔드
 * VERSION: V2.0
 * - [구조전환] 장비마스터(전체 대장) + 교정실적(이벤트 누적) + 교정차수 + 부적합조치
 *   · 교정결과(합격/조건부합격) 입력 시 장비마스터 최종교정일·차기교정일 자동 갱신(되돌리면 이전 값 복원)
 *   · 불합격/교정불가 입력 시 장비 상태 '사용중지' + 부적합조치 자동 생성
 * - [이관] V1 '교정대상' 탭이 있고 장비마스터가 비어 있으면 처음 실행 때 자동 이관(진행기록은 'V1 이관' 차수로)
 * - 로그인/권한(관리자·담당자·뷰어), 접속 비밀번호(ACCESS_CODE), 저장 잠금, 요청번호 중복 방지는 V1과 동일
 * [V1.0] 최초 배포 — 회차 리스트 방식
 * ===============================================================
 */
const GAS_VERSION = 'V2.0';

/* ---------------- 테이블 정의 ---------------- */
const T = {
  eq: { name: '장비마스터', cols: [
    ['id', '관리번호'], ['name', '장비명'], ['model', '모델'], ['maker', '제조사'], ['serial', 'SERIAL NO.'], ['range', '측정범위'], ['unit', '측정단위'],
    ['dept', '사용부서'], ['owner', '담당자'], ['location', '사용처'], ['cycle', '교정주기(개월)'], ['method', '교정방식'], ['agency', '교정기관'],
    ['status', '장비상태'], ['introDate', '도입일'], ['grr', 'GR&R'], ['lastCal', '최종교정일'], ['nextDue', '차기교정일'], ['note', '비고'],
    ['createdAt', '등록일시'], ['updatedAt', '수정일시'], ['updatedBy', '수정자']] },
  ev: { name: '교정실적', cols: [
    ['recId', '실적ID'], ['batchId', '차수ID'], ['eqId', '관리번호'], ['eqName', '장비명'], ['dept', '사용부서'], ['owner', '담당자'], ['dueDate', '교정기한(계획)'],
    ['inDate', '입고일'], ['inGiver', '인계자'], ['inTaker', '인수자'], ['appearance', '외관·작동'], ['accessory', '부속품'],
    ['calDate', '교정일'], ['result', '교정결과'], ['certNo', '성적서번호'], ['cost', '비용(원)'], ['newDue', '신규 차기교정일'], ['label', '라벨부착'],
    ['outDate', '출하일'], ['outReceiver', '수령자'], ['certUrl', '성적서 링크'], ['note', '비고'], ['state', '상태'],
    ['prevLastCal', '(이전)최종교정일'], ['prevNextDue', '(이전)차기교정일'], ['prevStatus', '(이전)장비상태'],
    ['createdAt', '생성일시'], ['updatedAt', '수정일시'], ['updatedBy', '수정자']] },
  batch: { name: '교정차수', cols: [
    ['batchId', '차수ID'], ['title', '차수명'], ['agency', '교정기관'], ['planDate', '방문/발송 예정일'], ['memo', '메모'], ['closed', '종료'],
    ['createdAt', '생성일시'], ['createdBy', '생성자']] },
  nc: { name: '부적합조치', cols: [
    ['ncId', '부적합ID'], ['eqId', '관리번호'], ['recId', '실적ID'], ['foundDate', '발생일'], ['content', '부적합 내용'], ['action', '조치'],
    ['impact', '영향성 검토'], ['state', '상태'], ['closedDate', '완료일'], ['updatedAt', '수정일시'], ['updatedBy', '수정자']] },
  log: { name: '변경이력', cols: [['at', '일시'], ['by', '처리자'], ['eqId', '관리번호'], ['kind', '구분'], ['detail', '내용']] },
  users: { name: '사용자', cols: [['username', '아이디'], ['pw', '비밀번호(암호화)'], ['name', '이름'], ['role', '권한'], ['createdAt', '가입일시']] },
  meta: { name: '설정', cols: [['key', '항목'], ['value', '값']] }
};
const EQ_MASTER_IMPORT = ['name', 'unit', 'cycle', 'agency', 'serial', 'location', 'owner', 'dept', 'grr'];
const EQ_EDITABLE = ['name', 'model', 'maker', 'serial', 'range', 'unit', 'dept', 'owner', 'location', 'cycle', 'method', 'agency', 'introDate', 'grr', 'lastCal', 'nextDue', 'note'];
const STEP_FIELDS = {
  in: ['inDate', 'inGiver', 'inTaker', 'appearance', 'accessory'],
  cal: ['calDate', 'result', 'certNo', 'cost', 'label'],
  out: ['outDate', 'outReceiver']
};
const EV_EDITABLE = [].concat(STEP_FIELDS.in, STEP_FIELDS.cal, STEP_FIELDS.out, ['certUrl', 'note']);
const EQ_STATUSES = ['사용', '보관', '수리', '사용중지', '폐기'];
const ROLE_LEVEL = { viewer: 1, editor: 2, admin: 3 };
const TZ = 'Asia/Seoul';

/* ---------------- 공통 ---------------- */
function jsonResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function getSS() { return SpreadsheetApp.getActiveSpreadsheet(); }
function props() { return PropertiesService.getScriptProperties(); }
function nowStr() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function cellStr(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}
function sheetOf(t) {
  const ss = getSS(); let sh = ss.getSheetByName(t.name);
  if (!sh) {
    sh = ss.insertSheet(t.name);
    const hd = t.cols.map(function (c) { return c[1]; });
    sh.getRange(1, 1, 1, hd.length).setValues([hd]).setFontWeight('bold').setBackground('#dbe4f3');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), hd.length).setNumberFormat('@');
  }
  return sh;
}
function readT(t) {
  const sh = sheetOf(t); const last = sh.getLastRow(); const n = t.cols.length;
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, n).getValues().map(function (r) {
    const o = {}; t.cols.forEach(function (c, i) { o[c[0]] = cellStr(r[i]); }); return o;
  }).filter(function (o) { return o[t.cols[0][0]]; });
}
function writeT(t, rows) {
  const sh = sheetOf(t); const n = t.cols.length; const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, n).clearContent();
  if (!rows.length) return;
  const out = rows.map(function (o) { return t.cols.map(function (c) { return o[c[0]] === undefined || o[c[0]] === null ? '' : String(o[c[0]]); }); });
  sh.getRange(2, 1, out.length, n).setNumberFormat('@').setValues(out);
}
function appendT(t, rows) {
  if (!rows.length) return;
  const sh = sheetOf(t); const n = t.cols.length;
  const out = rows.map(function (o) { return t.cols.map(function (c) { return o[c[0]] === undefined || o[c[0]] === null ? '' : String(o[c[0]]); }); });
  sh.getRange(sh.getLastRow() + 1, 1, out.length, n).setNumberFormat('@').setValues(out);
}
function blank(t) { const o = {}; t.cols.forEach(function (c) { o[c[0]] = ''; }); return o; }
function label(t, f) { for (let i = 0; i < t.cols.length; i++) if (t.cols[i][0] === f) return t.cols[i][1]; return f; }
function log(rows) { const at = nowStr(); appendT(T.log, rows.map(function (r) { return { at: at, by: r[0], eqId: r[1], kind: r[2], detail: r[3] }; })); }
function newId(prefix) { return prefix + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 900 + 100); }

function getMeta(key) { const r = readT(T.meta).filter(function (o) { return o.key === key; })[0]; return r ? r.value : ''; }
function setMeta(key, value) {
  const rows = readT(T.meta); const r = rows.filter(function (o) { return o.key === key; })[0];
  if (r) r.value = value; else rows.push({ key: key, value: value });
  writeT(T.meta, rows);
}
function addMonths(ymd, months) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || '')); const n = parseInt(months, 10);
  if (!m || !n) return '';
  const y = +m[1], mo = +m[2] - 1 + n, d = +m[3];
  const last = new Date(y, mo + 1, 0).getDate();
  return Utilities.formatDate(new Date(y, mo, Math.min(d, last)), TZ, 'yyyy-MM-dd');
}
function isPass(r) { return r === '합격' || r === '조건부합격'; }
function isFail(r) { return r === '불합격' || r === '교정불가'; }

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
function makeToken(u) { return u + '|' + Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(u), tokenSecret())); }
function findUser(username) {
  const rows = readT(T.users);
  for (let i = 0; i < rows.length; i++) if (rows[i].username.toLowerCase() === String(username || '').toLowerCase()) { rows[i].role = rows[i].role || 'viewer'; return rows[i]; }
  return null;
}
function authUser(token) {
  const t = String(token || ''); const p = t.lastIndexOf('|'); if (p < 1) return null;
  const u = t.slice(0, p); if (makeToken(u) !== t) return null;
  return findUser(u);
}
function need(body, role) {
  const u = authUser(body._token);
  if (!u) return { err: { status: 'fail', authFail: true, message: '다시 로그인해주세요.' } };
  if ((ROLE_LEVEL[u.role] || 0) < ROLE_LEVEL[role]) return { err: { status: 'fail', message: '권한이 없습니다.' } };
  return { user: u };
}
function userView(u) { return { username: u.username, name: u.name, role: u.role }; }
function handleLogin(b) {
  const u = findUser(b.username);
  if (!u || u.pw !== sha256(b.password)) return { status: 'fail', message: '아이디 또는 비밀번호가 맞지 않습니다.' };
  return { status: 'success', user: userView(u), token: makeToken(u.username) };
}
function handleRegister(b) {
  const username = String(b.username || '').trim(), name = String(b.name || '').trim();
  if (!/^[A-Za-z0-9_.-]{2,30}$/.test(username)) return { status: 'fail', message: '아이디는 영문/숫자 2~30자로 입력해주세요.' };
  if (String(b.password || '').length < 4) return { status: 'fail', message: '비밀번호는 4자 이상으로 입력해주세요.' };
  if (!name) return { status: 'fail', message: '이름을 입력해주세요.' };
  if (findUser(username)) return { status: 'fail', message: '이미 있는 아이디입니다.' };
  const first = readT(T.users).length === 0;
  appendT(T.users, [{ username: username, pw: sha256(b.password), name: name, role: first ? 'admin' : 'viewer', createdAt: nowStr() }]);
  return { status: 'success', user: userView(findUser(username)), token: makeToken(username), firstAdmin: first };
}
function handleChangePassword(b) {
  const a = need(b, 'viewer'); if (a.err) return a.err;
  if (a.user.pw !== sha256(b.oldPassword)) return { status: 'fail', message: '현재 비밀번호가 맞지 않습니다.' };
  if (String(b.newPassword || '').length < 4) return { status: 'fail', message: '새 비밀번호는 4자 이상' };
  const rows = readT(T.users); rows.forEach(function (o) { if (o.username === a.user.username) o.pw = sha256(b.newPassword); }); writeT(T.users, rows);
  return { status: 'success' };
}
function handleListUsers(b) {
  const a = need(b, 'admin'); if (a.err) return a.err;
  return { status: 'success', data: readT(T.users).map(function (u) { return { username: u.username, name: u.name, role: u.role, createdAt: u.createdAt }; }) };
}
function handleUpdateUser(b) {
  const a = need(b, 'admin'); if (a.err) return a.err;
  let rows = readT(T.users); const u = rows.filter(function (o) { return o.username === b.username; })[0];
  if (!u) return { status: 'fail', message: '사용자 없음' };
  if (b.remove) {
    if (u.username === a.user.username) return { status: 'fail', message: '본인 계정은 삭제할 수 없습니다.' };
    rows = rows.filter(function (o) { return o !== u; });
  } else if (b.resetPassword) u.pw = sha256(b.resetPassword);
  else {
    if (!ROLE_LEVEL[b.role]) return { status: 'fail', message: '권한 값 오류' };
    if (u.username === a.user.username && b.role !== 'admin') return { status: 'fail', message: '본인의 관리자 권한은 해제할 수 없습니다.' };
    u.role = b.role;
  }
  writeT(T.users, rows); return { status: 'success' };
}

/* ---------------- V1 → V2 자동 이관 ---------------- */
function migrateV1IfNeeded() {
  if (readT(T.eq).length) return false;
  const old = getSS().getSheetByName('교정대상');
  if (!old || old.getLastRow() < 2) return false;
  const F = ['id', 'name', 'unit', 'cycle', 'lastCal', 'nextDue', 'agency', 'serial', 'location', 'owner', 'dept', 'grr', 'method', 'round',
    'inDate', 'inGiver', 'inTaker', 'appearance', 'accessory', 'calDate', 'result', 'certNo', 'newDue', 'label', 'outDate', 'outReceiver', 'note', 'active', 'updatedAt', 'updatedBy'];
  const vals = old.getRange(2, 1, old.getLastRow() - 1, F.length).getValues();
  const t = nowStr(); const eqs = [], evs = []; let batchId = '', round = '';
  vals.forEach(function (r) {
    const v = {}; F.forEach(function (f, i) { v[f] = cellStr(r[i]); }); if (!v.id) return;
    const e = blank(T.eq);
    ['id', 'name', 'unit', 'cycle', 'lastCal', 'nextDue', 'agency', 'serial', 'location', 'owner', 'dept', 'grr', 'method', 'note'].forEach(function (f) { e[f] = v[f]; });
    e.method = e.method || '입고'; e.status = v.active === 'N' ? '보관' : '사용'; e.createdAt = t; e.updatedAt = t; e.updatedBy = 'V1 이관';
    if (v.inDate || v.calDate || v.outDate) {
      round = round || v.round; batchId = batchId || 'B-V1-' + (v.round || 'import');
      const ev = blank(T.ev);
      ['inDate', 'inGiver', 'inTaker', 'appearance', 'accessory', 'calDate', 'result', 'certNo', 'label', 'outDate', 'outReceiver', 'note'].forEach(function (f) { ev[f] = v[f]; });
      ev.recId = newId('R') + eqs.length; ev.batchId = batchId; ev.eqId = e.id; ev.eqName = e.name; ev.dept = e.dept; ev.owner = e.owner;
      ev.dueDate = e.nextDue; ev.state = ev.outDate ? '완료' : '진행'; ev.createdAt = t; ev.updatedAt = t; ev.updatedBy = 'V1 이관';
      applyCalToMaster(ev, e);
      evs.push(ev);
    }
    eqs.push(e);
  });
  writeT(T.eq, eqs);
  if (evs.length) {
    appendT(T.ev, evs);
    appendT(T.batch, [{ batchId: batchId, title: (round || '') + ' 교정 (V1 이관)', agency: eqs[0] ? eqs[0].agency : '', planDate: '', memo: 'V1 진행기록 자동 이관', closed: '', createdAt: t, createdBy: 'V1 이관' }]);
  }
  log([['시스템', '-', 'V1 이관', '장비 ' + eqs.length + '대, 진행기록 ' + evs.length + '건']]);
  return true;
}

/* ---------------- 교정결과 → 장비마스터 반영 ---------------- */
// ev: 실적, e: 장비. 처음 반영할 때 이전 값 저장(되돌리기용)
function applyCalToMaster(ev, e) {
  if (!ev.calDate || !ev.result) return;
  if (!ev.prevStatus) { ev.prevLastCal = e.lastCal; ev.prevNextDue = e.nextDue; ev.prevStatus = e.status || '사용'; }
  if (isPass(ev.result)) {
    ev.newDue = addMonths(ev.calDate, e.cycle);
    e.lastCal = ev.calDate; if (ev.newDue) e.nextDue = ev.newDue;
    if (e.status === '사용중지' || e.status === '수리') e.status = ev.prevStatus === '사용중지' || ev.prevStatus === '수리' ? '사용' : (ev.prevStatus || '사용');
  } else if (isFail(ev.result)) {
    ev.newDue = ''; e.lastCal = ev.prevLastCal; e.nextDue = ev.prevNextDue; e.status = '사용중지';
  }
}
function restoreMaster(ev, e) {
  if (!ev.prevStatus) return;
  e.lastCal = ev.prevLastCal; e.nextDue = ev.prevNextDue; e.status = ev.prevStatus;
  ev.prevLastCal = ''; ev.prevNextDue = ''; ev.prevStatus = ''; ev.newDue = '';
}

/* ---------------- 장비 대장 ---------------- */
function handleImportMaster(b) {
  const a = need(b, 'admin'); if (a.err) return a.err;
  const rows = readT(T.eq); const map = {}; rows.forEach(function (o, i) { map[o.id] = o; });
  const t = nowStr(); let added = 0, updated = 0, dateUpd = 0;
  (b.rows || []).forEach(function (s) {
    const id = String(s.id || '').trim(); if (!id) return;
    let e = map[id];
    if (!e) { e = blank(T.eq); e.id = id; e.status = '사용'; e.method = s.method === '현장' ? '현장' : '입고'; e.createdAt = t; rows.push(e); map[id] = e; added++; }
    else updated++;
    EQ_MASTER_IMPORT.forEach(function (f) { if (s[f] !== undefined && s[f] !== null && String(s[f]) !== '') e[f] = String(s[f]); });
    // 교정일은 더 최근 값만 반영(웹앱에서 이미 처리한 최신 교정을 옛 엑셀이 덮어쓰지 않게)
    if (s.lastCal && (!e.lastCal || s.lastCal > e.lastCal)) { e.lastCal = s.lastCal; if (s.nextDue) e.nextDue = s.nextDue; dateUpd++; }
    else if (!e.nextDue && s.nextDue) e.nextDue = s.nextDue;
    e.updatedAt = t; e.updatedBy = a.user.name;
  });
  writeT(T.eq, rows);
  log([[a.user.name, '-', '대장 가져오기', '신규 ' + added + ', 갱신 ' + updated + ', 교정일 갱신 ' + dateUpd]]);
  return { status: 'success', added: added, updated: updated, dateUpdated: dateUpd, data: readAllData() };
}
function handleSaveEquipment(b) {
  const a = need(b, 'admin'); if (a.err) return a.err;
  const s = b.eq || {}; const id = String(s.id || '').trim();
  if (!id) return { status: 'fail', message: '관리번호를 입력해주세요.' };
  const rows = readT(T.eq); let e = rows.filter(function (o) { return o.id === id; })[0]; const t = nowStr();
  if (b.isNew) {
    if (e) return { status: 'fail', message: '이미 있는 관리번호입니다.' };
    e = blank(T.eq); e.id = id; e.status = '사용'; e.createdAt = t; rows.push(e);
  } else if (!e) return { status: 'fail', message: '장비 없음' };
  const changed = [];
  EQ_EDITABLE.forEach(function (f) { if (s[f] !== undefined && String(s[f]) !== e[f]) { changed.push(label(T.eq, f) + ' ' + (e[f] || '-') + '→' + (s[f] || '-')); e[f] = String(s[f]); } });
  if (s.lastCal !== undefined && s.cycle && s.lastCal && !s.nextDue) e.nextDue = addMonths(s.lastCal, s.cycle);
  e.updatedAt = t; e.updatedBy = a.user.name;
  writeT(T.eq, rows);
  log([[a.user.name, id, b.isNew ? '장비 등록' : '장비 수정', changed.join(', ')]]);
  return { status: 'success', eq: e };
}
function handleSetEquipStatus(b) {
  const a = need(b, 'admin'); if (a.err) return a.err;
  if (EQ_STATUSES.indexOf(b.status) < 0) return { status: 'fail', message: '상태 값 오류' };
  const ids = {}; (b.ids || []).forEach(function (x) { ids[x] = 1; });
  const rows = readT(T.eq); const out = [], hist = []; const t = nowStr();
  rows.forEach(function (e) {
    if (!ids[e.id] || e.status === b.status) return;
    hist.push([a.user.name, e.id, '상태 변경', (e.status || '-') + ' → ' + b.status + (b.reason ? ' (' + b.reason + ')' : '')]);
    e.status = b.status; e.updatedAt = t; e.updatedBy = a.user.name; out.push(e);
  });
  writeT(T.eq, rows); log(hist);
  return { status: 'success', items: out };
}

/* ---------------- 교정 차수 ---------------- */
function handleCreateBatch(b) {
  const a = need(b, 'editor'); if (a.err) return a.err;
  const title = String(b.title || '').trim(); if (!title) return { status: 'fail', message: '차수명을 입력해주세요.' };
  const bt = { batchId: newId('B'), title: title, agency: b.agency || '', planDate: b.planDate || '', memo: b.memo || '', closed: '', createdAt: nowStr(), createdBy: a.user.name };
  appendT(T.batch, [bt]);
  const r = addEvents(bt.batchId, b.eqIds || [], a.user.name);
  log([[a.user.name, '-', '차수 생성', title + ' (' + r.added.length + '대)']]);
  return { status: 'success', batch: bt, events: r.added, skipped: r.skipped };
}
function addEvents(batchId, eqIds, who) {
  const eqs = readT(T.eq); const evs = readT(T.ev); const t = nowStr();
  const open = {}; evs.forEach(function (v) { if (v.state === '진행') open[v.eqId] = v.batchId; });
  const added = [], skipped = [];
  eqIds.forEach(function (id, i) {
    const e = eqs.filter(function (o) { return o.id === id; })[0];
    if (!e) return;
    if (open[id]) { skipped.push(id); return; }
    const v = blank(T.ev);
    v.recId = newId('R') + i; v.batchId = batchId; v.eqId = id; v.eqName = e.name; v.dept = e.dept; v.owner = e.owner; v.dueDate = e.nextDue;
    v.state = '진행'; v.createdAt = t; v.updatedAt = t; v.updatedBy = who;
    added.push(v); open[id] = batchId;
  });
  appendT(T.ev, added);
  return { added: added, skipped: skipped };
}
function handleAddToBatch(b) {
  const a = need(b, 'editor'); if (a.err) return a.err;
  if (!readT(T.batch).some(function (x) { return x.batchId === b.batchId; })) return { status: 'fail', message: '차수 없음' };
  const r = addEvents(b.batchId, b.eqIds || [], a.user.name);
  log([[a.user.name, '-', '차수 추가', b.batchId + ' +' + r.added.length + '대']]);
  return { status: 'success', events: r.added, skipped: r.skipped };
}
function handleUpdateBatch(b) {
  const a = need(b, 'editor'); if (a.err) return a.err;
  const rows = readT(T.batch); const bt = rows.filter(function (x) { return x.batchId === b.batchId; })[0];
  if (!bt) return { status: 'fail', message: '차수 없음' };
  ['title', 'agency', 'planDate', 'memo', 'closed'].forEach(function (f) { if (b[f] !== undefined) bt[f] = String(b[f]); });
  writeT(T.batch, rows);
  return { status: 'success', batch: bt };
}
// 진행 중 실적 제외(취소) — 교정 결과가 아직 없을 때만
function handleRemoveEvents(b) {
  const a = need(b, 'editor'); if (a.err) return a.err;
  const ids = {}; (b.recIds || []).forEach(function (x) { ids[x] = 1; });
  const evs = readT(T.ev); let n = 0; const hist = [];
  const keep = evs.filter(function (v) {
    if (!ids[v.recId]) return true;
    if (v.calDate || v.outDate) return true;
    n++; hist.push([a.user.name, v.eqId, '차수 제외', v.batchId]); return false;
  });
  writeT(T.ev, keep); log(hist);
  return { status: 'success', removed: n };
}

/* ---------------- 단계 처리 ---------------- */
function handleUpdateEvents(b) {
  const a = need(b, 'editor'); if (a.err) return a.err;
  const patch = b.patch || {}; const keys = Object.keys(patch).filter(function (k) { return EV_EDITABLE.indexOf(k) >= 0; });
  const ids = {}; (b.recIds || []).forEach(function (x) { ids[x] = 1; });
  if (!keys.length) return { status: 'fail', message: '변경할 내용이 없습니다.' };
  const evs = readT(T.ev), eqs = readT(T.eq); const eqMap = {}; eqs.forEach(function (e) { eqMap[e.id] = e; });
  const ncs = readT(T.nc); const t = nowStr(); const who = a.user.name;
  const outEv = [], outEq = [], newNc = [], hist = [];
  evs.forEach(function (v) {
    if (!ids[v.recId]) return;
    const e = eqMap[v.eqId];
    keys.forEach(function (k) { v[k] = patch[k] === null || patch[k] === undefined ? '' : String(patch[k]); });
    if (e && (keys.indexOf('calDate') >= 0 || keys.indexOf('result') >= 0)) {
      if (v.prevStatus) restoreMaster(v, e);
      applyCalToMaster(v, e);
      e.updatedAt = t; e.updatedBy = who; outEq.push(e);
      if (isFail(v.result) && !ncs.some(function (n) { return n.recId === v.recId; }) && !newNc.some(function (n) { return n.recId === v.recId; })) {
        newNc.push({ ncId: newId('N') + newNc.length, eqId: v.eqId, recId: v.recId, foundDate: v.calDate, content: '교정결과 ' + v.result, action: '', impact: '', state: '조치중', closedDate: '', updatedAt: t, updatedBy: who });
      }
    }
    v.state = v.outDate ? '완료' : '진행';
    v.updatedAt = t; v.updatedBy = who;
    hist.push([who, v.eqId, b.kind || '수정', keys.map(function (k) { return label(T.ev, k) + ' ' + (v[k] || '-'); }).join(', ')]);
    outEv.push(v);
  });
  if (!outEv.length) return { status: 'fail', message: '대상 실적을 찾지 못했습니다.' };
  writeT(T.ev, evs); if (outEq.length) writeT(T.eq, eqs); if (newNc.length) appendT(T.nc, newNc); log(hist);
  return { status: 'success', events: outEv, eqs: outEq, ncs: newNc };
}
function handleUndoStep(b) {
  const a = need(b, 'admin'); if (a.err) return a.err;
  const fields = STEP_FIELDS[b.step]; if (!fields) return { status: 'fail', message: '단계 값 오류' };
  const evs = readT(T.ev), eqs = readT(T.eq);
  const v = evs.filter(function (x) { return x.recId === b.recId; })[0]; if (!v) return { status: 'fail', message: '실적 없음' };
  const e = eqs.filter(function (x) { return x.id === v.eqId; })[0];
  if (b.step === 'cal' && e) restoreMaster(v, e);
  fields.forEach(function (f) { v[f] = ''; });
  v.state = v.outDate ? '완료' : '진행'; v.updatedAt = nowStr(); v.updatedBy = a.user.name;
  writeT(T.ev, evs); if (e) writeT(T.eq, eqs);
  log([[a.user.name, v.eqId, '되돌리기', ({ in: '입고', cal: '교정', out: '출하' })[b.step] + ' 취소' + (b.reason ? ' — ' + b.reason : '')]]);
  return { status: 'success', events: [v], eqs: e ? [e] : [] };
}

/* ---------------- 부적합 조치 ---------------- */
function handleSaveNc(b) {
  const a = need(b, 'editor'); if (a.err) return a.err;
  const rows = readT(T.nc); const s = b.nc || {}; let n = rows.filter(function (x) { return x.ncId === s.ncId; })[0];
  if (!n) {
    if (!s.eqId) return { status: 'fail', message: '관리번호 필요' };
    n = blank(T.nc); n.ncId = newId('N'); n.eqId = s.eqId; n.foundDate = s.foundDate || today(); n.state = '조치중'; rows.push(n);
  }
  ['content', 'action', 'impact', 'state', 'closedDate', 'foundDate'].forEach(function (f) { if (s[f] !== undefined) n[f] = String(s[f]); });
  if (n.state === '완료' && !n.closedDate) n.closedDate = today();
  if (n.state !== '완료') n.closedDate = '';
  n.updatedAt = nowStr(); n.updatedBy = a.user.name;
  writeT(T.nc, rows);
  log([[a.user.name, n.eqId, '부적합', n.state + ' · ' + (n.action || '')]]);
  return { status: 'success', nc: n };
}

/* ---------------- 조회 ---------------- */
function readAllData() {
  return { eqs: readT(T.eq), events: readT(T.ev), batches: readT(T.batch), ncs: readT(T.nc), version: GAS_VERSION };
}
function getLog(eqId) {
  const rows = readT(T.log); const out = [];
  for (let i = rows.length - 1; i >= 0 && out.length < 300; i--) if (!eqId || rows[i].eqId === eqId) out.push(rows[i]);
  return out;
}

/* ---------------- 진입점 ---------------- */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {}; const action = p.action || 'read';
    if (action === 'version') return jsonResponse({ status: 'success', data: GAS_VERSION });
    if (action === 'checkAccess') return jsonResponse({ status: 'success', required: !!accessCodeValue(), ok: accessOk(p.ak), version: GAS_VERSION });
    if (!accessOk(p.ak)) return jsonResponse({ status: 'fail', accessDenied: true, message: '접속 비밀번호가 필요합니다.' });
    if (action === 'read') {
      let migrated = false;
      try { const lock = LockService.getScriptLock(); if (lock.tryLock(20000)) { try { migrated = migrateV1IfNeeded(); } finally { lock.releaseLock(); } } } catch (e1) {}
      const d = readAllData(); d.migrated = migrated;
      return jsonResponse({ status: 'success', data: d });
    }
    if (action === 'getLog') return jsonResponse({ status: 'success', data: getLog(p.eqId) });
    return jsonResponse({ status: 'fail', message: 'Unknown GET action: ' + action });
  } catch (err) { return jsonResponse({ status: 'error', message: err.toString() }); }
}
function doPost(e) {
  let lock = null;
  try {
    const action = e.parameter.action;
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (!accessOk(e.parameter.ak || body._ak)) return jsonResponse({ status: 'fail', accessDenied: true, message: '접속 비밀번호가 필요합니다.' });
    lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) return jsonResponse({ status: 'error', busy: true, message: '다른 저장을 처리 중입니다. 잠시 후 다시 시도해주세요.' });
    const reqId = body._reqId ? String(body._reqId) : ''; const cache = CacheService.getScriptCache();
    if (reqId) { const c = cache.get('req_' + reqId); if (c) return ContentService.createTextOutput(c).setMimeType(ContentService.MimeType.JSON); }
    const text = JSON.stringify(routePost(action, body));
    if (reqId && text.length < 90000) cache.put('req_' + reqId, text, 600);
    return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
  } catch (err) { return jsonResponse({ status: 'error', message: err.toString() }); }
  finally { if (lock) { try { lock.releaseLock(); } catch (e2) {} } }
}
function routePost(action, b) {
  switch (action) {
    case 'login': return handleLogin(b);
    case 'register': return handleRegister(b);
    case 'changePassword': return handleChangePassword(b);
    case 'listUsers': return handleListUsers(b);
    case 'updateUser': return handleUpdateUser(b);
    case 'importMaster': return handleImportMaster(b);
    case 'saveEquipment': return handleSaveEquipment(b);
    case 'setEquipStatus': return handleSetEquipStatus(b);
    case 'createBatch': return handleCreateBatch(b);
    case 'addToBatch': return handleAddToBatch(b);
    case 'updateBatch': return handleUpdateBatch(b);
    case 'removeEvents': return handleRemoveEvents(b);
    case 'updateEvents': return handleUpdateEvents(b);
    case 'undoStep': return handleUndoStep(b);
    case 'saveNc': return handleSaveNc(b);
    default: return { status: 'fail', message: 'Unknown POST action: ' + action };
  }
}
function setup() { Object.keys(T).forEach(function (k) { sheetOf(T[k]); }); tokenSecret(); migrateV1IfNeeded(); }
