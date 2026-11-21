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
  '산행': ['id', 'title', 'mt', 'start', 'end', 'members', 'gear', 'check', 'by', 'at', 'type'],
  '차박지': ['id', 'name', 'region', 'lat', 'lon', 'floor', 'fee', 'cook', 'water', 'toilet', 'level', 'note', 'by', 'at'],
  '차박기록': ['id', 'spot', 'user', 'date', 'nights', 'score', 'note', 'at'],
  '화장실좌표': ['id', 'lat', 'lon', 'ok', 'at'],
  '사진': ['id', 'user', 'mt', 'date', 'file', 'thumb', 'at'],
  '설정': ['key', 'value'],
  '접속': ['at', 'user', 'dev', 'via', 'action'],
};
const COLORS = ['#4f93c9', '#e0806a', '#9b7fd1', '#5aa57a', '#d6a33c', '#d46a9a', '#3fa5a5', '#8a8f3c', '#c9643c', '#6a7fd4'];
const CAR_TEMPLATE = [
  { cat: '잠자리', items: [['차량용 매트'], ['침낭'], ['베개'], ['암막 커튼·가림막'], ['방충망']] },
  { cat: '전기·난방', items: [['대용량 파워뱅크', 1], ['전기요'], ['무시동 히터', 1], ['랜턴']] },
  { cat: '안전', items: [['일산화탄소 경보기'], ['환기용 창문 틈막이'], ['소화기', 1], ['비상 휴대용 변기']] },
  { cat: '기타', items: [['물·물티슈'], ['쓰레기봉투'], ['수평계 앱 확인']] }];
const TEMPLATE = [
  { cat: '잠자리', items: [['텐트', 1], ['침낭'], ['매트'], ['베개'], ['그라운드시트', 1]] },
  { cat: '주방', items: [['버너', 1], ['가스', 1], ['코펠', 1], ['수저·컵'], ['라이터']] },
  { cat: '옷', items: [['경량 패딩'], ['우모복'], ['장갑'], ['비니'], ['여벌 양말']] },
  { cat: '안전', items: [['헤드랜턴'], ['보조배터리'], ['구급약', 1], ['스틱']] },
  { cat: '물·먹거리', items: [['식수 2L'], ['행동식'], ['저녁·아침', 1]] },
  { cat: '기타', items: [['쓰레기봉투'], ['물티슈·휴지'], ['랜턴', 1]] }];

/* ── 처음 한 번 ── */
const SCHEMA = 2;
function migrate_() {
  const P = PropertiesService.getScriptProperties();
  if (+P.getProperty('SCHEMA') === SCHEMA) return;
  Object.keys(SHEETS).forEach(n => {
    let sh = ss_().getSheetByName(n);
    if (!sh) { sh = ss_().insertSheet(n); sh.appendRow(SHEETS[n]); sh.setFrozenRows(1); sh.getRange(1, 1, 1000, SHEETS[n].length).setNumberFormat('@'); return; }
    const h = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    SHEETS[n].forEach((k, i) => { if (h[i] !== k) sh.getRange(1, i + 1).setValue(k); });
  });
  if (!getSetting_('carTemplate', null)) setSetting_('carTemplate', CAR_TEMPLATE);
  P.setProperty('SCHEMA', String(SCHEMA)); sh0_ = {}; rowsDirty_();
}
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
  PropertiesService.getScriptProperties().deleteProperty('SCHEMA'); migrate_();
  if (prop_('DATA_GO_KR_KEY')) installParkingTrigger_();
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
    migrate_(); touch_(me, q);
    const W = { checkin: checkin_, photo: photo_, delPhoto: delPhoto_, check: check_, gear: gear_, trip: trip_, delTrip: delTrip_, joinTrip: joinTrip_,
      approve: approve_, addCar: addCar_, carLog: carLog_, delCar: delCar_, delCarLog: delCarLog_, reject: reject_, addMt: addMt_, template: template_, setting: settingSet_, delVisit: delVisit_ };
    if (a === 'load') return { ok: true, data: load_(me) };
    if (a === 'img') return { ok: true, data: img_(me, q) };
    if (a === 'near') return { ok: true, data: near_(me, q) };
    if (a === 'weather') return { ok: true, data: weather_(+q.lat, +q.lon) };
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
    type: t.type === 'car' ? 'car' : 'back', members: String(t.members || '').split(',').filter(String), gear: json_(t.gear, {}), check: json_(t.check, {}) }));
  const logs = rows_('차박기록');
  const carspots = rows_('차박지').map(c => ({ id: c.id, name: c.name, region: c.region, lat: +c.lat, lon: +c.lon, floor: c.floor, fee: c.fee, cook: c.cook, water: c.water, toilet: c.toilet, level: c.level, note: c.note, by: c.by,
    logs: logs.filter(l => l.spot === c.id).map(l => ({ id: l.id, user: l.user, date: l.date, nights: +l.nights || 0, score: +l.score || 0, note: l.note })) }));
  const photos = rows_('사진').map(p => ({ id: p.id, user: p.user, mt: p.mt, date: p.date }));
  const mountains = rows_('산').map(m => ({ id: m.id, name: m.name, region: m.region, alt: +m.alt || 0, lat: m.lat === '' ? null : +m.lat, lon: m.lon === '' ? null : +m.lon,
    top100: m.top100 === 'true' || m.top100 === true, kind: m.kind || '산' }));
  const data = { me: me.id, today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), people, mountains, visits, trips, photos, carspots, carTemplate: getSetting_('carTemplate', CAR_TEMPLATE),
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
  append_('산행', { id: newId_('t'), title: String(t.title || '').slice(0, 40), mt: t.mt, start: t.start, end: t.end, members, gear: '{}', check: '{}', by: me.id, at: now_(), type: t.type === 'car' ? 'car' : 'back' });
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
  setSetting_(q.kind === 'car' ? 'carTemplate' : 'template', q.template.map(c => ({ cat: String(c.cat).slice(0, 12), items: (c.items || []).map(i => [String(i[0]).slice(0, 20)].concat(i[1] ? [1] : [])) })));
}
function settingSet_(me, q) { admin_(me); if (q.radius) setSetting_('radius', Math.max(100, Math.min(1000, +q.radius))); }

/* ── 주변 정보: 화장실·주차장·차박지 ──
   화장실: 저장소 data/wc/<지역코드>.json (전국공중화장실표준데이터, 좌표 없음) → 산 근처 시·군·구 것만 주소를 좌표로 바꿔(구글 지오코더) 시트 '화장실좌표'에 기억.
           한 번에 25곳까지만 바꾸고(실행 시간·하루 한도), 남은 수를 앱에 알려 주면 앱이 다시 부른다.
   주차장: 전국주차장정보표준데이터 API 를 일주일에 한 번 통째로 받아 드라이브 파일(parking.json)에 둔다.
   '내 위치' 근처를 찾을 때 받은 좌표는 찾기에만 쓰고 저장하지 않는다. */
const PAGES_URL = 'https://friese8807.github.io/sandojang/';
const WC_KW = /산|등산|공원|휴양림|캠핑|야영|오토|주차|해수욕|해변|유원지|계곡|쉼터|전망|정상|약수|숲|둘레길|자연|관광|휴게|생태|수목원|저수지|댐|섬|항|포구|폭포|사찰|사$|암$|성$|광장|탐방|매표|입구|고개|재$|령$/;
function near_(me, q) {
  const lat = +q.lat, lon = +q.lon;
  if (!(lat > 32 && lat < 39.5 && lon > 124 && lon < 132)) throw new Error('위치를 확인해 주세요');
  const t0 = Date.now(), out = { wc: [], pk: [], car: [], left: 0, more: 0, note: [] };
  // 1) 이 근처 시·군·구 (산은 기억, 내 위치는 기억하지 않음)
  let areas = q.mt ? getSetting_('area:' + q.mt, null) : null;
  if (!areas) { areas = areasNear_(lat, lon); if (q.mt) setSetting_('area:' + q.mt, areas); }
  // 2) 화장실
  try {
    const have = {}; rows_('화장실좌표').forEach(r => have[r.id] = r);
    const todo = [];
    areas.codes.forEach(code => {
      const list = wcList_(code);
      list.forEach(w => {
        const id = w[0], c = have[id];
        if (c) { if (c.ok === '1') { const km = km_(lat, lon, +c.lat, +c.lon); if (km <= 5) out.wc.push(wcOut_(w, +c.lat, +c.lon, km)); } return; }
        if (w[1] === '개방' && !WC_KW.test(w[2])) return;   // 가게·관공서 개방화장실은 산 근처 후보에서 뺀다
        todo.push({ w, pri: areas.dong.some(d => (w[3] + ' ' + w[4]).includes(d)) ? 0 : 1 });
      });
    });
    todo.sort((a, b) => a.pri - b.pri);
    const geo = Maps.newGeocoder().setRegion('kr').setLanguage('ko'), rows = [];
    // 같은 읍·면·동 먼저 다 찾고, 시·군 전체는 '더 넓게 찾기'(q.all)를 눌렀을 때만 (구글 주소 변환은 하루 한도가 있다)
    const go = q.all ? todo : todo.filter(x => x.pri === 0);
    for (const { w } of go) {
      if (rows.length >= 25 || Date.now() - t0 > 15000) break;
      let p = null;
      for (const a of [w[3], w[4]].filter(String)) { try { const r = geo.geocode(a); const g = r.results && r.results[0]; if (g) { p = g.geometry.location; break; } } catch (e) { out.note.push('주소 변환 한도'); break; } }
      rows.push([w[0], p ? p.lat : '', p ? p.lng : '', p ? '1' : '0', now_()]);
      if (p) { const km = km_(lat, lon, p.lat, p.lng); if (km <= 5) out.wc.push(wcOut_(w, p.lat, p.lng, km)); }
    }
    if (rows.length) { const sh = sh_('화장실좌표'); sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setNumberFormat('@').setValues(rows); rowsDirty_('화장실좌표'); }
    out.left = Math.max(0, go.length - rows.length); out.more = q.all ? 0 : todo.filter(x => x.pri === 1).length;
  } catch (e) { out.note.push('화장실: ' + e.message); }
  out.wc.sort((a, b) => a.km - b.km); out.wc = out.wc.slice(0, 30);
  // 3) 주차장
  try {
    const pk = parkingAll_();
    if (!pk) out.note.push(prop_('DATA_GO_KR_KEY') ? '주차장 자료를 받는 중이에요. 몇 분 뒤 다시 봐 주세요' : '주차장: 공공데이터 인증키가 아직 없어요');
    else out.pk = pk.map(p => ({ name: p[0], type: p[1], lots: p[2], fee: p[3], feeD: p[4], hours: p[5], tel: p[6], lat: p[7], lon: p[8], addr: p[9], km: km_(lat, lon, p[7], p[8]) }))
      .filter(p => p.km <= 5).sort((a, b) => a.km - b.km).slice(0, 20);
  } catch (e) { out.note.push('주차장: ' + e.message); }
  // 4) 차박지 (우리가 남긴 곳)
  out.car = rows_('차박지').map(c => ({ id: c.id, km: km_(lat, lon, +c.lat, +c.lon) })).filter(c => c.km <= 30).sort((a, b) => a.km - b.km);
  out.areas = areas.names;
  return out;
}
const km_ = (a, b, c, d) => { const r = Math.PI / 180, x = Math.sin((c - a) * r / 2) ** 2 + Math.cos(a * r) * Math.cos(c * r) * Math.sin((d - b) * r / 2) ** 2; return 12742 * Math.asin(Math.sqrt(x)); };
function wcOut_(w, lat, lon, km) {
  return { id: w[0], kind: w[1] + '화장실', name: w[2], addr: w[3] || w[4], m: w[5], u: w[6], w: w[7], dm: w[8], dw: w[9], org: w[10], tel: w[11], open: w[12], openD: w[13], bell: w[14], cctv: w[15], baby: w[16], lat, lon, km: Math.round(km * 100) / 100 };
}
/* 지역 파일은 GitHub Pages 에서 받아 6시간 기억 (100KB 넘으면 나눠서) */
let wcMem_ = {};
function wcList_(code) {
  if (wcMem_[code]) return wcMem_[code];
  const C = CacheService.getScriptCache(), n = +C.get('wc:' + code + ':n');
  let txt = '';
  if (n) { const parts = C.getAll([...Array(n).keys()].map(i => 'wc:' + code + ':' + i)); if (Object.keys(parts).length === n) txt = [...Array(n).keys()].map(i => parts['wc:' + code + ':' + i]).join(''); }
  if (!txt) {
    const r = UrlFetchApp.fetch(PAGES_URL + 'data/wc/' + code + '.json', { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return (wcMem_[code] = []);
    txt = r.getContentText('UTF-8');
    const ch = {}, k = Math.ceil(txt.length / 90000); for (let i = 0; i < k; i++) ch['wc:' + code + ':' + i] = txt.slice(i * 90000, (i + 1) * 90000);
    try { C.putAll(ch, 21600); C.put('wc:' + code + ':n', String(k), 21600); } catch (e) {}
  }
  return (wcMem_[code] = JSON.parse(txt));
}
/* 가운데와 사방 3km 지점의 주소로 시·군·구를 찾는다 (경계에 걸친 산: 선자령=평창·강릉) */
function areasNear_(lat, lon) {
  const idx = JSON.parse(UrlFetchApp.fetch(PAGES_URL + 'data/wc/index.json').getContentText('UTF-8')).areas;
  const geo = Maps.newGeocoder().setRegion('kr').setLanguage('ko'), d = .027;
  const codes = {}, names = {}, dong = {};
  [[0, 0], [d, 0], [-d, 0], [0, d * 1.25], [0, -d * 1.25]].forEach(([a, b], i) => {
    let res; try { res = geo.reverseGeocode(lat + a, lon + b); } catch (e) { return; }
    const g = res.results && res.results[0]; if (!g) return;
    const comp = g.address_components || [], get = t => (comp.find(c => c.types.indexOf(t) >= 0) || {}).long_name || '';
    const sido = get('administrative_area_level_1'), sgg = get('locality') || get('sublocality_level_1') || get('administrative_area_level_2');
    const gu = get('sublocality_level_1');
    Object.keys(idx).forEach(code => {
      const nm = idx[code];
      if (sgg && nm.includes(sgg) && (!sido || nm.slice(0, 2) === sido.slice(0, 2)) && (!gu || gu === sgg || nm.includes(gu))) { codes[code] = 1; names[nm] = 1; }
    });
    if (i === 0) { const dd = get('sublocality_level_2') || get('sublocality_level_1'); if (dd) dong[dd] = 1; }
  });
  return { codes: Object.keys(codes), names: Object.keys(names), dong: Object.keys(dong) };
}
/* ── 주차장 (전국주차장정보표준데이터) ── 필드 이름이 조금 달라도 찾아 쓴다. 먼저 debugParking() 으로 응답 확인 */
const PARK_URL = 'https://api.data.go.kr/openapi/tn_pubr_prkplce_info_api';
function parkingFetch_(page, rows) {
  const key = prop_('DATA_GO_KR_KEY'); if (!key) throw new Error('스크립트 속성 DATA_GO_KR_KEY 에 공공데이터포털 인증키를 넣어 주세요');
  const url = (prop_('PARK_URL') || PARK_URL) + '?serviceKey=' + encodeURIComponent(key) + '&pageNo=' + page + '&numOfRows=' + rows + '&type=json';
  let res = null; for (let i = 0; i < 3 && !res; i++) { try { res = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); } catch (e) { Utilities.sleep(2000 * (i + 1)); } }
  if (!res) throw new Error('주차장 API 연결 실패');
  const txt = res.getContentText(); let j;
  try { j = JSON.parse(txt); } catch (e) { throw new Error('JSON 이 아닌 응답 (' + res.getResponseCode() + '): ' + txt.slice(0, 300).replace(/serviceKey=[^&\s]+/g, 'serviceKey=***')); }
  if (j.OpenAPI_ServiceResponse) throw new Error('인증 오류: ' + JSON.stringify(j.OpenAPI_ServiceResponse.cmmMsgHeader));
  const b = (j.response || j).body || {}, h = (j.response || j).header || {};
  if (h.resultCode && h.resultCode !== '00') throw new Error('주차장 API: ' + h.resultCode + ' ' + (h.resultMsg || ''));
  const it = b.items && (b.items.item || b.items);
  return { items: Array.isArray(it) ? it : it ? [it] : [], total: +b.totalCount || 0 };
}
function debugParking() { const r = parkingFetch_(1, 3); console.log('전체 ' + r.total + '곳, 첫 줄:\n' + JSON.stringify(r.items[0], null, 1)); }
function syncParking() {
  const pick = (o, re) => { const k = Object.keys(o).find(k => re.test(k)); return k ? String(o[k] == null ? '' : o[k]).trim() : ''; };
  let page = 1, all = [], total = 0;
  do { const r = parkingFetch_(page, 1000); total = r.total; all = all.concat(r.items); page++; if (!r.items.length) break; } while (all.length < total && page < 60);
  const out = all.map(o => {
    const lat = +pick(o, /^latitude$|위도/i), lon = +pick(o, /^longitude$|경도/i);
    if (!(lat > 32 && lat < 39.5 && lon > 124 && lon < 132)) return null;
    const fee = pick(o, /parkingchrgeInfo|요금정보/i), bt = pick(o, /basicTime|기본시간/i), bc = pick(o, /basicCharge|기본요금/i), day = pick(o, /dayCmmtkt$|1일주차권요금$/i);
    const feeD = fee === '무료' ? '' : [bt && bc ? bt + '분 ' + Number(bc).toLocaleString() + '원' : '', day && day !== '0' ? '하루 ' + Number(day).toLocaleString() + '원' : ''].filter(String).join(' · ');
    const o1 = pick(o, /weekdayOperOpenHhmm|평일운영시작/i), o2 = pick(o, /weekdayOperColseHhmm|weekdayOperCloseHhmm|평일운영종료/i);
    return [pick(o, /prkplceNm|주차장명/i), pick(o, /prkplceSe|주차장구분/i), +pick(o, /prkcmprt|주차구획수/i) || 0, fee || '-', feeD,
      o1 && o2 ? (o1 === '00:00' && /^(00:00|23:59|24:00)$/.test(o2) ? '24시간' : o1 + '~' + o2) : '', pick(o, /phoneNumber|전화번호/i), Math.round(lat * 1e5) / 1e5, Math.round(lon * 1e5) / 1e5, pick(o, /rdnmadr|소재지도로명주소/i) || pick(o, /lnmadr|지번/i)];
  }).filter(Boolean);
  const blob = Utilities.newBlob(JSON.stringify(out), 'application/json', 'parking.json'), id = prop_('PARK_FILE');
  let f = null; if (id) { try { f = DriveApp.getFileById(id); f.setContent(blob.getDataAsString()); } catch (e) { f = null; } }
  if (!f) { f = DriveApp.getFolderById(prop_('FOLDER_ID')).createFile(blob); PropertiesService.getScriptProperties().setProperty('PARK_FILE', f.getId()); }
  PropertiesService.getScriptProperties().setProperty('PARK_AT', now_());
  CacheService.getScriptCache().remove('pk:n');
  console.log('✓ 주차장 ' + out.length + '곳 저장 (받은 자료 ' + all.length + '건)');
}
function installParkingTrigger_() {
  if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'syncParking')) return;
  ScriptApp.newTrigger('syncParking').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
  ScriptApp.newTrigger('syncParking').timeBased().after(60 * 1000).create();   // 처음 한 번은 바로
}
let pkMem_ = null;
function parkingAll_() {
  if (pkMem_) return pkMem_;
  const id = prop_('PARK_FILE');
  if (!id) { if (prop_('DATA_GO_KR_KEY')) try { installParkingTrigger_(); } catch (e) {} return null; }
  return (pkMem_ = JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString()));
}
/* ── 날씨 (기상청 단기예보) ── 5km 격자 기준이라 정상은 이보다 추워요(100m 오를 때 약 0.6도) */
function grid_(lat, lon) {
  const RE = 6371.00877 / 5.0, SL1 = 30 * Math.PI / 180, SL2 = 60 * Math.PI / 180, OLON = 126 * Math.PI / 180, OLAT = 38 * Math.PI / 180, r = Math.PI / 180;
  let sn = Math.log(Math.cos(SL1) / Math.cos(SL2)) / Math.log(Math.tan(Math.PI * .25 + SL2 * .5) / Math.tan(Math.PI * .25 + SL1 * .5));
  let sf = Math.pow(Math.tan(Math.PI * .25 + SL1 * .5), sn) * Math.cos(SL1) / sn, ro = RE * sf / Math.pow(Math.tan(Math.PI * .25 + OLAT * .5), sn);
  let ra = RE * sf / Math.pow(Math.tan(Math.PI * .25 + lat * r * .5), sn), th = lon * r - OLON;
  if (th > Math.PI) th -= 2 * Math.PI; if (th < -Math.PI) th += 2 * Math.PI; th *= sn;
  return [Math.floor(ra * Math.sin(th) + 43 + .5), Math.floor(ro - ra * Math.cos(th) + 136 + .5)];
}
function weather_(lat, lon) {
  if (!(lat > 32 && lat < 39.5)) throw new Error('위치를 확인해 주세요');
  const [nx, ny] = grid_(lat, lon), C = CacheService.getScriptCache(), ck = 'wx:' + nx + ',' + ny, hit = C.get(ck);
  if (hit) return JSON.parse(hit);
  const key = prop_('KMA_KEY') || prop_('DATA_GO_KR_KEY'); if (!key) throw new Error('날씨: 공공데이터 인증키가 아직 없어요');
  // 발표 시각: 02·05·08·11·14·17·20·23시 (+10분 뒤 나옴)
  const now = new Date(Date.now() - 15 * 60e3), kst = new Date(now.getTime() + 9 * 36e5);
  let h = kst.getUTCHours(), bd = new Date(kst); const hs = [23, 20, 17, 14, 11, 8, 5, 2];
  let bh = hs.find(x => x <= h); if (bh == null) { bh = 23; bd = new Date(kst.getTime() - 864e5); }
  const ymd = Utilities.formatDate(bd, 'UTC', 'yyyyMMdd'), bt = ('0' + bh).slice(-2) + '00';
  const url = 'https://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getVilageFcst?serviceKey=' + encodeURIComponent(key) + '&pageNo=1&numOfRows=1000&dataType=JSON&base_date=' + ymd + '&base_time=' + bt + '&nx=' + nx + '&ny=' + ny;
  let res = null; for (let i = 0; i < 3 && !res; i++) { try { res = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); } catch (e) { Utilities.sleep(1500 * (i + 1)); } }
  if (!res) throw new Error('기상청 연결 실패');
  let j; try { j = JSON.parse(res.getContentText()); } catch (e) { throw new Error('기상청 응답을 읽지 못했어요 (' + res.getResponseCode() + ')'); }
  if (j.OpenAPI_ServiceResponse) throw new Error('기상청 인증 오류: ' + JSON.stringify(j.OpenAPI_ServiceResponse.cmmMsgHeader));
  const hd = j.response && j.response.header; if (hd && hd.resultCode !== '00') throw new Error('기상청: ' + hd.resultMsg);
  const items = (((j.response || {}).body || {}).items || {}).item || [];
  const T = {}; items.forEach(x => { const k = x.fcstDate + x.fcstTime; (T[k] = T[k] || { d: x.fcstDate, t: x.fcstTime })[x.category] = x.fcstValue; });
  const hours = Object.keys(T).sort().map(k => T[k]), days = {};
  hours.forEach(x => {
    const d = days[x.d] = days[x.d] || { d: x.d, min: 99, max: -99, pop: 0, wsd: 0, sky: 0, pty: 0 };
    if (x.TMP != null) { d.min = Math.min(d.min, +x.TMP); d.max = Math.max(d.max, +x.TMP); }
    if (x.TMN != null) d.min = Math.min(d.min, +x.TMN); if (x.TMX != null) d.max = Math.max(d.max, +x.TMX);
    d.pop = Math.max(d.pop, +x.POP || 0); d.wsd = Math.max(d.wsd, +x.WSD || 0); d.sky = Math.max(d.sky, +x.SKY || 0); if (+x.PTY) d.pty = +x.PTY;
  });
  const out = { at: ymd + bt, nx, ny, now: hours[0] ? { t: +hours[0].TMP, sky: +hours[0].SKY, pty: +hours[0].PTY, pop: +hours[0].POP, wsd: +hours[0].WSD, reh: +hours[0].REH } : null,
    days: Object.keys(days).sort().slice(0, 4).map(k => days[k]),
    hours: hours.slice(0, 30).map(x => ({ d: x.d, t: x.t, tmp: +x.TMP, sky: +x.SKY, pty: +x.PTY, pop: +x.POP, wsd: +x.WSD })) };
  try { C.put(ck, JSON.stringify(out), 3600); } catch (e) {}
  return out;
}
/* ── 차박지 ── 장소 좌표는 '이 자리를 차박지로 남기기'를 누른 사람이 일부러 올린 것만 */
function addCar_(me, q) {
  const c = q.car || {}, name = String(c.name || '').trim().slice(0, 30), lat = +c.lat, lon = +c.lon;
  if (!name) throw new Error('이름을 넣어 주세요');
  if (!(lat > 32 && lat < 39.5 && lon > 124 && lon < 132)) throw new Error('우리나라 안의 좌표가 아니에요');
  const pick = (v, ok) => ok.indexOf(v) >= 0 ? v : '';
  const id = newId_('c');
  append_('차박지', { id, name, region: String(c.region || '').slice(0, 20), lat: Math.round(lat * 1e5) / 1e5, lon: Math.round(lon * 1e5) / 1e5,
    floor: pick(c.floor, ['아스팔트', '파쇄석', '흙', '잔디', '모래']), fee: pick(c.fee, ['무료', '유료']), cook: pick(c.cook, ['취사 가능', '취사 금지', '모름']),
    water: pick(c.water, ['개수대 있음', '개수대 없음']), toilet: String(c.toilet || '').slice(0, 30), level: pick(c.level, ['평탄', '약간 경사', '경사']),
    note: String(c.note || '').slice(0, 300), by: me.id, at: now_() });
  if (c.date) carLog_(me, { spot: id, date: c.date, nights: c.nights, score: c.score, note: '' });
}
function carLog_(me, q) {
  if (!rows_('차박지').some(c => c.id === q.spot)) throw new Error('없는 차박지예요');
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), date = /^\d{4}-\d\d-\d\d$/.test(q.date || '') && q.date <= now ? q.date : now;
  append_('차박기록', { id: newId_('l'), spot: q.spot, user: me.id, date, nights: Math.max(0, Math.min(30, +q.nights || 1)), score: Math.max(0, Math.min(5, +q.score || 0)), note: String(q.note || '').slice(0, 300), at: now_() });
}
function delCar_(me, q) {
  const c = rows_('차박지').find(x => x.id === q.id); if (!c) return;
  if (c.by !== me.id && me.role !== 'admin') throw new Error('남긴 사람만 지울 수 있어요');
  rows_('차박기록').filter(l => l.spot === c.id).sort((a, b) => b._row - a._row).forEach(l => del_('차박기록', l._row));
  del_('차박지', rows_('차박지').find(x => x.id === q.id)._row);
}
function delCarLog_(me, q) {
  const l = rows_('차박기록').find(x => x.id === q.id); if (!l) return;
  if (l.user !== me.id && me.role !== 'admin') throw new Error('내 기록만 지울 수 있어요');
  del_('차박기록', l._row);
}
function debugWeather() { console.log(JSON.stringify(weather_(37.689, 128.752), null, 1).slice(0, 2000)); }

/* ── 100대 명산 가져오기 (공공데이터포털 '한국등산트레킹지원센터_100대명산' 계열) ──
   응답 모양을 먼저 debugTop100() 으로 확인한다. 필드 이름이 달라도 이름·위도·경도·높이를 찾아 쓴다. */
const TOP100_URL = 'https://apis.data.go.kr/B553662/top100FamtListBasiInfoService/getTop100FamtListBasiInfoList';
function top100Fetch_(page, rows) {
  const key = prop_('DATA_GO_KR_KEY'); if (!key) throw new Error('스크립트 속성 DATA_GO_KR_KEY 에 공공데이터포털 인증키를 넣어 주세요');
  const url = (prop_('TOP100_URL') || TOP100_URL) + '?serviceKey=' + encodeURIComponent(key) + '&pageNo=' + page + '&numOfRows=' + (rows || 20) + '&type=json';
  // 이 기관 서버는 한 번에 많이 달라면 30초 넘게 걸려 포털이 'SERVICETIMEOUT_ERROR'를 준다 → 조금씩, 시간 초과면 쉬었다 다시 (최대 4번)
  let last = '';
  for (let i = 0; i < 4; i++) {
    if (i) Utilities.sleep(3000 * i);
    let res; try { res = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); } catch (e) { last = '연결 실패: ' + e.message; continue; }
    const txt = res.getContentText(); let j;
    try { j = JSON.parse(txt); } catch (e) { last = 'JSON 이 아닌 응답 (' + res.getResponseCode() + '): ' + txt.slice(0, 300).replace(/serviceKey=[^&\s]+/g, 'serviceKey=***'); if (/TIMEOUT/i.test(txt)) continue; throw new Error(last); }
    if (j.OpenAPI_ServiceResponse) { const h = j.OpenAPI_ServiceResponse.cmmMsgHeader || {}; last = '포털 오류: ' + JSON.stringify(h); if (/TIMEOUT/i.test(h.errMsg || '')) continue; throw new Error(last); }
    const hd = (j.response || j).header || {}; if (hd.resultCode && hd.resultCode !== '00') throw new Error('100대명산 API: ' + hd.resultCode + ' ' + (hd.resultMsg || ''));
    const body = (j.response || j).body || {}, it = body.items && (body.items.item || body.items);
    return { items: Array.isArray(it) ? it : it ? [it] : [], total: +body.totalCount || 0, raw: j };
  }
  throw new Error(last + ' — 기관 서버가 응답하지 않아요. 잠시 뒤 다시 실행해 주세요');
}
function debugTop100() { const r = top100Fetch_(1, 3); console.log('전체 ' + r.total + '건, 첫 줄:\n' + JSON.stringify(r.items[0], null, 1)); }
function importTop100() {
  const pick = (o, re) => { const k = Object.keys(o).find(k => re.test(k)); return k ? o[k] : ''; };
  const have = rows_('산'), byName = {}; have.forEach(m => byName[m.name] = m);
  let page = 1, n = 0, all = [];
  while (page < 20) { const r = top100Fetch_(page, 20); all = all.concat(r.items); if (!r.items.length || all.length >= r.total) break; page++; }
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
