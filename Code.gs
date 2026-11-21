/* 산도장 서버 (Google Apps Script) — 화면은 GitHub Pages, 자료는 이 서버의 구글 시트·드라이브에만.
   ■ 처음 한 번: 편집기에서 setup() 실행 → 허용. 실행 기록에 서버 주소·시트 주소가 나온다.
   ■ 첫 가입자는 자동으로 관리자. 그다음부터는 관리자가 승인해야 쓸 수 있다.
   ■ 위치 원칙: 휴대폰이 '인증됨/기록만'을 판정하고 서버에는 어느 산·인증 여부·거리(100m 단위)만 온다. 사람의 좌표는 받지도 저장하지도 않는다.
   ■ 사진: 드라이브 '산도장 사진' 폴더(공유하지 않음). 화면은 로그인한 사람에게만 이 서버를 거쳐 보낸다.
   ■ 100대 명산: 공공데이터포털 인증키를 스크립트 속성 DATA_GO_KR_KEY 에 넣고 debugTop100() → importTop100() */

const TZ = 'Asia/Seoul';
const SHEETS = {
  '회원': ['id', 'name', 'pinHash', 'token', 'role', 'status', 'color', 'at', 'lastAt'],
  '산': ['id', 'name', 'region', 'alt', 'lat', 'lon', 'top100', 'kind', 'by', 'code'],
  '기록': ['id', 'user', 'mt', 'date', 'verified', 'dist', 'nights', 'at'],
  '산행': ['id', 'title', 'mt', 'start', 'end', 'members', 'gear', 'check', 'by', 'at'],
  '사진': ['id', 'user', 'mt', 'date', 'file', 'thumb', 'at'],
  '설정': ['key', 'value'],
  '접속': ['at', 'user', 'dev', 'via', 'action'],
};
const COLORS = ['#4f93c9', '#e0806a', '#9b7fd1', '#5aa57a', '#d6a33c', '#d46a9a', '#3fa5a5', '#8a8f3c', '#c9643c', '#6a7fd4'];
const TEMPLATE = [
  { cat: '잠자리', items: [['텐트', 1], ['침낭'], ['매트'], ['베개'], ['그라운드시트', 1]] },
  { cat: '주방', items: [['버너', 1], ['가스', 1], ['코펠', 1], ['수저·컵'], ['라이터']] },
  { cat: '옷', items: [['경량 패딩'], ['우모복'], ['장갑'], ['비니'], ['여벌 양말']] },
  { cat: '안전', items: [['헤드랜턴'], ['보조배터리'], ['구급약', 1], ['스틱']] },
  { cat: '물·먹거리', items: [['식수 2L'], ['행동식'], ['저녁·아침', 1]] },
  { cat: '기타', items: [['쓰레기봉투'], ['물티슈·휴지'], ['랜턴', 1]] }];

/* ── 처음 한 번 ── */
function setup() {
  const P = PropertiesService.getScriptProperties();
  let ss = P.getProperty('SHEET_ID') ? SpreadsheetApp.openById(P.getProperty('SHEET_ID')) : null;
  if (!ss) { ss = SpreadsheetApp.create('산도장 자료'); P.setProperty('SHEET_ID', ss.getId()); }
  Object.keys(SHEETS).forEach(n => {
    let s = ss.getSheetByName(n);
    if (!s) s = ss.insertSheet(n);
    if (s.getLastRow() === 0) { s.appendRow(SHEETS[n]); s.setFrozenRows(1); s.getRange(1, 1, 1000, SHEETS[n].length).setNumberFormat('@'); }
  });
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1'); if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  if (!P.getProperty('FOLDER_ID')) P.setProperty('FOLDER_ID', DriveApp.createFolder('산도장 사진').getId());
  if (!getSetting_('template', null)) setSetting_('template', TEMPLATE);
  console.log('✓ 준비 끝\n서버 주소: ' + ScriptApp.getService().getUrl() + '\n자료 시트: ' + ss.getUrl() +
    '\n사진 폴더: https://drive.google.com/drive/folders/' + P.getProperty('FOLDER_ID') +
    '\n앱 설정 탭의 "서버 주소" 칸에 위 서버 주소(/exec 로 끝남)를 넣고, 처음 가입하는 사람이 관리자가 돼요.');
}

/* ── 요청 받기 ── */
function doPost(e) { return out_(handle_(JSON.parse(e.postData.contents || '{}'))); }
function doGet(e) {
  const p = e.parameter || {};
  if (!p.q) return ContentService.createTextOutput('산도장 서버가 동작하고 있어요.');
  const r = handle_(JSON.parse(p.q)), s = JSON.stringify(r);
  return p.callback ? ContentService.createTextOutput(p.callback.replace(/[^\w$.]/g, '') + '(' + s + ')').setMimeType(ContentService.MimeType.JAVASCRIPT) : out_(r);
}
const out_ = r => ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);

function handle_(q) {
  try {
    const a = q.action;
    if (a === 'join') return { ok: true, data: join_(q) };
    if (a === 'login') return { ok: true, data: login_(q) };
    const me = who_(q.auth);
    if (!me) return { ok: false, code: 'auth', message: '다시 로그인해 주세요' };
    if (me.status === 'pending') return { ok: false, code: 'pending', message: '관리자 승인을 기다리고 있어요' };
    if (me.status !== 'active') return { ok: false, code: 'auth', message: '쓸 수 없는 계정이에요. 관리자에게 물어봐 주세요' };
    touch_(me, q);
    const W = { checkin: checkin_, photo: photo_, delPhoto: delPhoto_, check: check_, gear: gear_, trip: trip_, delTrip: delTrip_, joinTrip: joinTrip_,
      approve: approve_, reject: reject_, addMt: addMt_, template: template_, setting: settingSet_, delVisit: delVisit_ };
    if (a === 'load') return { ok: true, data: load_(me) };
    if (a === 'img') return { ok: true, data: img_(me, q) };
    if (W[a]) {
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try { W[a](me, q); SpreadsheetApp.flush(); rowsDirty_(); } finally { lock.releaseLock(); }
      return { ok: true, data: load_(me) };
    }
    return { ok: false, message: '모르는 요청이에요: ' + a };
  } catch (err) {
    console.error(err);
    return { ok: false, message: String(err && err.message || err) };
  }
}

/* ── 로그인: 비밀번호는 해시로만 저장, 휴대폰에는 토큰만 ── */
const hash_ = s => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'sandojang:' + s));
const newId_ = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
function join_(q) {
  const name = String(q.name || '').trim().slice(0, 12), pin = String(q.pin || '');
  if (!name) throw new Error('이름을 넣어 주세요');
  if (!/^\d{6}$/.test(pin)) throw new Error('비밀번호는 숫자 6자리예요');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const us = rows_('회원');
    if (us.some(u => u.name === name)) throw new Error('같은 이름이 이미 있어요. 로그인하거나 다른 이름을 써 주세요');
    const first = !us.length, token = Utilities.getUuid();
    const u = { id: newId_('u'), name, pinHash: hash_(name + ':' + pin), token: hash_(token), role: first ? 'admin' : 'member', status: first ? 'active' : 'pending',
      color: COLORS[us.length % COLORS.length], at: now_(), lastAt: now_() };
    append_('회원', u);
    if (!first) notifyAdmins_(name);
    return { token, status: u.status, name };
  } finally { lock.releaseLock(); }
}
function login_(q) {
  const name = String(q.name || '').trim(), pin = String(q.pin || '');
  const u = rows_('회원').find(x => x.name === name && x.pinHash === hash_(name + ':' + pin));
  if (!u) throw new Error('이름이나 비밀번호가 맞지 않아요');
  if (u.status === 'rejected') throw new Error('가입이 거절됐어요. 관리자에게 물어봐 주세요');
  const token = Utilities.getUuid();
  set_('회원', u._row, 'token', hash_(token));
  return { token, status: u.status, name };
}
function who_(auth) {
  if (!auth || !auth.token) return null;
  const h = hash_(auth.token);
  return rows_('회원').find(u => u.token === h) || null;
}
function touch_(me, q) {
  if (q.action !== 'load') return;
  const last = me.lastAt ? new Date(me.lastAt).getTime() : 0;
  if (Date.now() - last < 30 * 60e3) return;   // 30분에 한 번만 적는다 (자동 새로고침마다 쓰지 않게)
  set_('회원', me._row, 'lastAt', now_());
  const m = q.meta || {}; append_('접속', { at: now_(), user: me.name, dev: m.dev || '', via: m.via || '', action: 'open' });
}
function notifyAdmins_(name) {
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '[산도장] ' + name + '님이 가입을 신청했어요', '앱의 설정 탭(또는 종 아이콘)에서 승인해 주세요.');
  } catch (e) { console.warn('가입 알림 메일', e); }
}

/* ── 불러오기: 보여 줄 것만 보낸다 (비밀번호·토큰·남의 거리 X) ── */
function load_(me) {
  const admin = me.role === 'admin';
  const us = rows_('회원');
  const people = us.filter(u => u.status === 'active').map(u => ({ id: u.id, name: u.name, color: u.color, role: u.role }));
  const visits = rows_('기록').map(v => ({ id: v.id, user: v.user, mt: v.mt, date: v.date, verified: v.verified === 'true' || v.verified === true, nights: +v.nights || 0,
    dist: v.user === me.id && v.dist !== '' ? +v.dist : undefined }));
  const trips = rows_('산행').map(t => ({ id: t.id, title: t.title, mt: t.mt, start: t.start, end: t.end, by: t.by,
    members: String(t.members || '').split(',').filter(String), gear: json_(t.gear, {}), check: json_(t.check, {}) }));
  const photos = rows_('사진').map(p => ({ id: p.id, user: p.user, mt: p.mt, date: p.date }));
  const mountains = rows_('산').map(m => ({ id: m.id, name: m.name, region: m.region, alt: +m.alt || 0, lat: m.lat === '' ? null : +m.lat, lon: m.lon === '' ? null : +m.lon,
    top100: m.top100 === 'true' || m.top100 === true, kind: m.kind || '산' }));
  const data = { me: me.id, today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), people, mountains, visits, trips, photos,
    template: getSetting_('template', TEMPLATE), radius: +getSetting_('radius', 300) };
  if (admin) data.admin = { pending: us.filter(u => u.status === 'pending').map(u => ({ id: u.id, name: u.name, at: String(u.at).slice(0, 16).replace('T', ' ') })) };
  return data;
}

/* ── 쓰기 ── */
function checkin_(me, q) {
  const m = rows_('산').find(x => x.id === q.mt); if (!m) throw new Error('없는 산이에요');
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  // 인증은 오늘만. '기록만'은 지난 산행 날짜를 받을 수 있다
  const today = !q.verified && /^\d{4}-\d\d-\d\d$/.test(q.date || '') && q.date <= now ? q.date : now;
  // 같은 날 같은 산은 한 번만 (인증됨이 오면 기록만을 인증됨으로 바꾼다)
  const old = rows_('기록').find(v => v.user === me.id && v.mt === q.mt && v.date === today);
  const dist = Math.round((+q.dist || 0) / 100) * 100;
  if (old) { if (q.verified && old.verified !== 'true') { set_('기록', old._row, 'verified', 'true'); set_('기록', old._row, 'dist', dist); } return; }
  append_('기록', { id: newId_('v'), user: me.id, mt: q.mt, date: today, verified: q.verified ? 'true' : 'false', dist, nights: +q.nights || 0, at: now_() });
}
function delVisit_(me, q) {
  const v = rows_('기록').find(x => x.id === q.id); if (!v) return;
  if (v.user !== me.id && me.role !== 'admin') throw new Error('내 기록만 지울 수 있어요');
  del_('기록', v._row);
}
function photo_(me, q) {
  if (!/^data:image\/jpeg;base64,/.test(q.data || '')) throw new Error('사진을 읽지 못했어요');
  const folder = DriveApp.getFolderById(prop_('FOLDER_ID')), id = newId_('p');
  const save = (d, n) => folder.createFile(Utilities.newBlob(Utilities.base64Decode(d.split(',')[1]), 'image/jpeg', n)).getId();
  const date = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const file = save(q.data, date + '_' + me.name + '_' + id + '.jpg');
  const thumb = q.thumb ? save(q.thumb, date + '_' + id + '_작게.jpg') : '';
  append_('사진', { id, user: me.id, mt: q.mt, date, file, thumb, at: now_() });
}
function delPhoto_(me, q) {
  const p = rows_('사진').find(x => x.id === q.id); if (!p) return;
  if (p.user !== me.id && me.role !== 'admin') throw new Error('내 사진만 지울 수 있어요');
  [p.file, p.thumb].filter(String).forEach(f => { try { DriveApp.getFileById(f).setTrashed(true); } catch (e) {} });
  del_('사진', p._row);
}
function img_(me, q) {
  const p = rows_('사진').find(x => x.id === q.id); if (!p) throw new Error('사진이 없어요');
  const f = DriveApp.getFileById(q.size === 'f' || !p.thumb ? p.file : p.thumb);
  return 'data:image/jpeg;base64,' + Utilities.base64Encode(f.getBlob().getBytes());
}
function tripRow_(id) { const t = rows_('산행').find(x => x.id === id); if (!t) throw new Error('없는 산행이에요'); return t; }
function check_(me, q) {
  const t = tripRow_(q.trip), c = json_(t.check, {}), a = c[me.id] = c[me.id] || [], i = a.indexOf(q.item);
  if (q.on && i < 0) a.push(q.item); if (!q.on && i >= 0) a.splice(i, 1);
  set_('산행', t._row, 'check', JSON.stringify(c));
}
function gear_(me, q) {
  const t = tripRow_(q.trip), g = json_(t.gear, {});
  g[q.item] = q.who || ''; set_('산행', t._row, 'gear', JSON.stringify(g));
}
function trip_(me, q) {
  const t = q.trip || {};
  if (!t.mt || !/^\d{4}-\d\d-\d\d$/.test(t.start) || !/^\d{4}-\d\d-\d\d$/.test(t.end) || t.end < t.start) throw new Error('날짜를 확인해 주세요');
  const members = [...new Set([me.id].concat(t.members || []))].join(',');
  append_('산행', { id: newId_('t'), title: String(t.title || '').slice(0, 40), mt: t.mt, start: t.start, end: t.end, members, gear: '{}', check: '{}', by: me.id, at: now_() });
}
function delTrip_(me, q) {
  const t = tripRow_(q.trip); if (t.by !== me.id && me.role !== 'admin') throw new Error('만든 사람만 지울 수 있어요');
  del_('산행', t._row);
}
function joinTrip_(me, q) {
  const t = tripRow_(q.trip), ms = String(t.members || '').split(',').filter(String), i = ms.indexOf(me.id);
  if (q.on && i < 0) ms.push(me.id); if (!q.on && i >= 0) ms.splice(i, 1);
  set_('산행', t._row, 'members', ms.join(','));
}
function admin_(me) { if (me.role !== 'admin') throw new Error('관리자만 할 수 있어요'); }
function approve_(me, q) { admin_(me); const u = rows_('회원').find(x => x.id === q.id); if (u) set_('회원', u._row, 'status', 'active'); }
function reject_(me, q) { admin_(me); const u = rows_('회원').find(x => x.id === q.id); if (u) set_('회원', u._row, 'status', 'rejected'); }
function addMt_(me, q) {
  const m = q.mt || {}, name = String(m.name || '').trim().slice(0, 30);
  if (!name) throw new Error('이름을 넣어 주세요');
  const lat = +m.lat, lon = +m.lon;
  if (!(lat > 32 && lat < 39.5 && lon > 124 && lon < 132)) throw new Error('우리나라 안의 좌표가 아니에요');
  if (q.id) { const r = rows_('산').find(x => x.id === q.id); if (!r) throw new Error('없는 산이에요'); set_('산', r._row, 'lat', lat); set_('산', r._row, 'lon', lon); return; }
  append_('산', { id: newId_('m'), name, region: String(m.region || '직접 추가').slice(0, 20), alt: Math.round(+m.alt || 0), lat, lon, top100: 'false', kind: m.kind === '야영지' ? '야영지' : '산', by: me.id, code: '' });
}
function template_(me, q) {
  if (!Array.isArray(q.template)) throw new Error('준비물 목록이 이상해요');
  setSetting_('template', q.template.map(c => ({ cat: String(c.cat).slice(0, 12), items: (c.items || []).map(i => [String(i[0]).slice(0, 20)].concat(i[1] ? [1] : [])) })));
}
function settingSet_(me, q) { admin_(me); if (q.radius) setSetting_('radius', Math.max(100, Math.min(1000, +q.radius))); }

/* ── 100대 명산 가져오기 (공공데이터포털 '한국등산트레킹지원센터_100대명산' 계열) ──
   응답 모양을 먼저 debugTop100() 으로 확인한다. 필드 이름이 달라도 이름·위도·경도·높이를 찾아 쓴다. */
const TOP100_URL = 'https://apis.data.go.kr/B553662/top100FamtListBasiInfoService/getTop100FamtListBasiInfoList';
function top100Fetch_(page) {
  const key = prop_('DATA_GO_KR_KEY'); if (!key) throw new Error('스크립트 속성 DATA_GO_KR_KEY 에 공공데이터포털 인증키를 넣어 주세요');
  const url = (prop_('TOP100_URL') || TOP100_URL) + '?serviceKey=' + encodeURIComponent(key) + '&pageNo=' + page + '&numOfRows=100&type=json&_type=json';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true }), txt = res.getContentText();
  let j; try { j = JSON.parse(txt); } catch (e) { throw new Error('JSON 이 아닌 응답 (' + res.getResponseCode() + '): ' + txt.slice(0, 300).replace(/serviceKey=[^&\s]+/g, 'serviceKey=***')); }
  if (j.OpenAPI_ServiceResponse) throw new Error('인증 오류: ' + JSON.stringify(j.OpenAPI_ServiceResponse.cmmMsgHeader));
  const body = (j.response || j).body || {}, it = body.items && (body.items.item || body.items);
  return { items: Array.isArray(it) ? it : it ? [it] : [], total: +body.totalCount || 0, raw: j };
}
function debugTop100() { const r = top100Fetch_(1); console.log('전체 ' + r.total + '건, 첫 줄:\n' + JSON.stringify(r.items[0], null, 1)); }
function importTop100() {
  const pick = (o, re) => { const k = Object.keys(o).find(k => re.test(k)); return k ? o[k] : ''; };
  const have = rows_('산'), byName = {}; have.forEach(m => byName[m.name] = m);
  let page = 1, n = 0, all = [];
  while (page < 10) { const r = top100Fetch_(page); all = all.concat(r.items); if (!r.items.length || all.length >= r.total) break; page++; }
  all.forEach(o => {
    const name = String(pick(o, /^(mntn|frtrl|mtn).*nm$|^mntnnm$|name|nm$/i) || '').trim();
    const lat = +pick(o, /^lat|lat$|^la$|ycrd|ypos/i), lon = +pick(o, /^lo?n|lot$|^lo$|lng|xcrd|xpos/i);
    const alt = Math.round(+pick(o, /hg|height|alt|aslAltd|hght/i) || 0), region = String(pick(o, /addr|adres|loc|sido|ctpv/i) || '').split(' ').slice(0, 2).join(' ');
    if (!name || byName[name]) return;
    append_('산', { id: newId_('m'), name, region, alt, lat: isFinite(lat) && lat > 30 ? lat : '', lon: isFinite(lon) && lon > 120 ? lon : '', top100: 'true', kind: '산', by: 'data.go.kr', code: String(pick(o, /cd$|code|id$/i) || '') });
    byName[name] = 1; n++;
  });
  console.log('✓ 새로 넣은 산 ' + n + '곳 (받은 자료 ' + all.length + '건)');
}

/* ── 점검: 편집기에서 실행하면 불러오기 시간을 잰다 ── */
function healthCheck() {
  const t = Date.now(), u = rows_('회원').find(x => x.role === 'admin');
  if (!u) return console.log('아직 가입한 사람이 없어요. 앱에서 먼저 가입해 주세요');
  const d = load_(u), ms = Date.now() - t;
  console.log((ms < 20000 ? '✓ 정상' : '✗ 느림') + ' · 불러오기 ' + ms + 'ms · 사람 ' + d.people.length + ' · 산 ' + d.mountains.length + ' · 기록 ' + d.visits.length + ' · 사진 ' + d.photos.length);
}

/* ── 시트 도우미 (한 요청 안에서는 한 번만 열고 한 번만 읽는다) ── */
let ss0_ = null, sh0_ = {}, rows0_ = {};
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function ss_() { if (ss0_) return ss0_; const id = prop_('SHEET_ID'); if (!id) throw new Error('서버 준비가 안 됐어요. 편집기에서 setup() 을 실행해 주세요'); return (ss0_ = SpreadsheetApp.openById(id)); }
function sh_(name) { return sh0_[name] || (sh0_[name] = ss_().getSheetByName(name)); }
function rowsDirty_(name) { if (name) delete rows0_[name]; else rows0_ = {}; }
function rows_(name) {
  if (!rows0_[name]) {
    const v = sh_(name).getDataRange().getValues(), h = v.shift();
    rows0_[name] = v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = cellIn_(r[j])); return o; }).filter(o => o[h[0]] !== '');
  }
  return rows0_[name].map(o => Object.assign({}, o));
}
function append_(name, o) { rowsDirty_(name); const s = sh_(name); s.appendRow(SHEETS[name].map(k => cellOut_(o[k]))); }
function set_(name, row, key, val) { rowsDirty_(name); sh_(name).getRange(row, SHEETS[name].indexOf(key) + 1).setValue(cellOut_(val)); }
function del_(name, row) { rowsDirty_(name); sh_(name).deleteRow(row); }
const cellIn_ = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v == null ? '' : v);
const cellOut_ = v => (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : (v == null ? '' : v);
const json_ = (s, d) => { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } };
const now_ = () => Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
function getSetting_(k, d) { const r = rows_('설정').find(x => x.key === k); return r ? json_(r.value, d) : d; }
function setSetting_(k, v) { const r = rows_('설정').find(x => x.key === k); if (r) set_('설정', r._row, 'value', JSON.stringify(v)); else append_('설정', { key: k, value: JSON.stringify(v) }); }
