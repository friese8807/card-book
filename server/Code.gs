/* ═══ 명함첩 서버 (Google Apps Script) ═══
   하는 일: ① 명함 사진 읽기(인식·분류) ② 명함을 구글 시트에, 사진을 구글 드라이브 폴더에 저장 ③ 앱에 목록·사진 내려주기

   처음 한 번:
   1) script.google.com → 새 프로젝트 → 이 파일 내용을 Code.gs 에 붙여 넣기
   2) 위 함수 고르는 칸에서 setup 실행 → 권한 허용 → 실행 기록에 나온 '서버 주소·연결 암호'는 배포 뒤에 다시 확인
   3) 배포 → 새 배포 → 유형 '웹 앱', 실행 '나', 액세스 '모든 사용자' → 배포 → /exec 로 끝나는 주소 복사
   4) 명함 읽기는 구글 드라이브 글자 인식(OCR, 무료)으로 해요. 따로 넣을 키가 없어요.
      (나중에 유료 AI 로 바꾸고 싶으면 스크립트 속성 ANTHROPIC_API_KEY 에 키를 넣으면 돼요)
   5) 앱 설정 탭에 서버 주소와 연결 암호를 넣고 '연결하기'
   코드를 고치면: 배포 → 배포 관리 → 연필 → 버전 '새 버전' → 배포 (주소는 그대로) */

const TZ = 'Asia/Seoul';
const COLS = ['id', 'name', 'nameEn', 'title', 'dept', 'org', 'cat', 'mobile', 'tel', 'fax', 'email', 'addr', 'web', 'met', 'place', 'memo', 'fav', 'checked', 'photoId', 'at', 'updated'];
const CATS = ['중앙부처', '소속·산하기관', '지자체', '공공기관', '기업', '사내', '기타'];
const MY_COMPANY = /올포랜드|allforland/i;          // 우리 회사 명함은 '사내'로
const MODEL = 'claude-opus-5-5';

/* ── 처음 설정 ── */
function setup() {
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('APP_TOKEN')) P.setProperty('APP_TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 20));
  sheet_(); folder_();
  const url = ScriptApp.getService().getUrl();
  console.log('서버 주소: ' + (url || '(아직 배포 전이에요. 배포한 뒤 setup 을 한 번 더 실행하세요)'));
  console.log('연결 암호: ' + P.getProperty('APP_TOKEN'));
  console.log('명함 시트: ' + SpreadsheetApp.openById(P.getProperty('SHEET_ID')).getUrl());
  console.log('사진 폴더: ' + DriveApp.getFolderById(P.getProperty('FOLDER_ID')).getUrl());
  console.log('명함 읽기: ' + (P.getProperty('ANTHROPIC_API_KEY') ? 'AI(Claude)' : '구글 OCR (ANTHROPIC_API_KEY 를 넣으면 AI 로 읽어요)'));
  if (url) console.log('한 번 누르는 연결 링크(사파리에서 열기): https://friese8807.github.io/card-book/#setup=' + Utilities.base64EncodeWebSafe(JSON.stringify({ url, token: P.getProperty('APP_TOKEN') })));
}
function sheet_() {
  const P = PropertiesService.getScriptProperties(); let id = P.getProperty('SHEET_ID'), ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create('명함첩'); P.setProperty('SHEET_ID', ss.getId()); }
  let sh = ss.getSheetByName('명함');
  if (!sh) { sh = ss.getSheets()[0]; sh.setName('명함'); sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold'); sh.setFrozenRows(1);
  }
  return sh;
}
function folder_() {
  const P = PropertiesService.getScriptProperties(); let id = P.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('명함첩 사진'); P.setProperty('FOLDER_ID', f.getId()); return f;
}

/* ── 요청 받기 ── 앱은 text/plain 으로 JSON 을 보낸다(사전 요청이 생기지 않게) */
function doPost(e) {
  let q = {}; try { q = JSON.parse(e.postData.contents); } catch (x) {}
  return out_(handle_(q));
}
/* 홈 화면 앱에서 POST 가 막힐 때를 위한 JSONP (작은 요청만) */
function doGet(e) {
  const p = e.parameter || {}, cb = /^[A-Za-z_$][\w$]*$/.test(p.callback || '') ? p.callback : '';
  let q = {}; try { q = JSON.parse(p.q || '{}'); } catch (x) {}
  const r = cb ? handle_(q) : { ok: true, data: '명함첩 서버가 동작 중이에요' };
  return cb ? ContentService.createTextOutput(cb + '(' + JSON.stringify(r) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT) : out_(r);
}
const out_ = r => ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);

function handle_(q) {
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('APP_TOKEN')) return { ok: false, code: 'setup', message: '서버에서 setup 을 먼저 실행해 주세요' };
  if (q.token !== P.getProperty('APP_TOKEN')) return { ok: false, code: 'auth', message: '연결 암호가 맞지 않아요' };
  try {
    switch (q.action) {
      case 'ping': return { ok: true, data: { ai: !!P.getProperty('ANTHROPIC_API_KEY'), count: Math.max(0, sheet_().getLastRow() - 1) } };
      case 'list': return { ok: true, data: { cards: list_() } };
      case 'scan': return { ok: true, data: scan_(q.img) };
      case 'save': return { ok: true, data: save_(q.card, q.photo) };
      case 'del': return { ok: true, data: del_(q.id) };
      case 'photos': return { ok: true, data: photos_(q.ids) };
      default: return { ok: false, message: '알 수 없는 요청이에요' };
    }
  } catch (err) {
    console.error(q.action, err.stack || err);
    return { ok: false, message: String(err.message || err).slice(0, 200) };
  }
}

/* ── 시트 읽고 쓰기 ── */
function rows_() {
  const sh = sheet_(), n = sh.getLastRow() - 1; if (n < 1) return { sh, rows: [] };
  return { sh, rows: sh.getRange(2, 1, n, COLS.length).getDisplayValues() };
}
const toCard_ = r => { const c = {}; COLS.forEach((k, i) => c[k] = r[i]); c.fav = c.fav === 'TRUE' || c.fav === 'true' || c.fav === '1'; c.checked = c.checked !== 'FALSE' && c.checked !== 'false' && c.checked !== '0'; c.at = Number(c.at) || 0; c.photo = !!c.photoId; return c; };
function list_() { return rows_().rows.filter(r => r[0]).map(toCard_); }
function save_(card, photo) {
  if (!card || !card.id || !String(card.name || '').trim()) throw new Error('이름이 없는 명함은 저장할 수 없어요');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const { sh, rows } = rows_(), i = rows.findIndex(r => r[0] === card.id);
    const old = i >= 0 ? toCard_(rows[i]) : {};
    const c = Object.assign({}, old, card, { updated: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') });
    if (!c.at) c.at = Date.now();
    if (photo && /^data:image\/\w+;base64,/.test(photo)) {
      const b = Utilities.newBlob(Utilities.base64Decode(photo.split(',')[1]), 'image/jpeg', (c.name + '_' + (c.org || '') + '_' + c.id + '.jpg').replace(/[\\/:*?"<>|]/g, ''));
      if (old.photoId) { try { DriveApp.getFileById(old.photoId).setTrashed(true); } catch (e) {} }
      c.photoId = folder_().createFile(b).getId();
    }
    // 모든 칸을 글자로(앞에 '): 전화번호 앞자리 0, 날짜가 '2026. 11. 12' 로 바뀌는 것, 수식(=…) 실행을 막는다
    const row = COLS.map(k => k === 'fav' || k === 'checked' ? (c[k] ? 'TRUE' : 'FALSE') : c[k] == null || c[k] === '' ? '' : "'" + String(c[k]));
    if (i >= 0) sh.getRange(i + 2, 1, 1, COLS.length).setValues([row]); else sh.appendRow(row);
    c.photo = !!c.photoId; return c;
  } finally { lock.releaseLock(); }
}
function del_(id) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const { sh, rows } = rows_(), i = rows.findIndex(r => r[0] === id); if (i < 0) return { id };
    const pid = rows[i][COLS.indexOf('photoId')]; if (pid) { try { DriveApp.getFileById(pid).setTrashed(true); } catch (e) {} }
    sh.deleteRow(i + 2); return { id };
  } finally { lock.releaseLock(); }
}
function photos_(ids) {
  const { rows } = rows_(), out = {};
  (ids || []).slice(0, 8).forEach(id => {
    const r = rows.find(x => x[0] === id), pid = r && r[COLS.indexOf('photoId')]; if (!pid) return;
    try { out[id] = 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(pid).getBlob().getBytes()); } catch (e) {}
  });
  return out;
}

/* ── 명함 읽기 ── AI 키가 있으면 Claude, 없으면 구글 드라이브 OCR + 규칙 */
function scan_(img) {
  if (!/^data:image\/\w+;base64,/.test(img || '')) throw new Error('사진이 오지 않았어요');
  const b64 = img.split(',')[1], key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  let r, via;
  if (key) { try { r = claude_(b64, key); via = 'ai'; } catch (e) { console.warn('AI 읽기 실패 → OCR', e.message); r = null; } }
  if (!r) { r = parse_(ocr_(b64)); via = 'ocr'; }
  learn_(r);
  r.cat = CATS.indexOf(r.cat) >= 0 && r.cat !== '기타' ? r.cat : classify_(r.org, r.email, r.dept);
  if (MY_COMPANY.test(r.org || '')) r.cat = '사내';
  ['mobile', 'tel', 'fax'].forEach(k => r[k] = phone_(r[k]));
  r.via = via; return r;
}

const SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['isCard', 'name', 'nameEn', 'title', 'dept', 'org', 'cat', 'mobile', 'tel', 'fax', 'email', 'addr', 'web', 'box'],
  properties: {
    isCard: { type: 'boolean', description: '사진에 명함이 있으면 true' },
    name: { type: 'string', description: '한글 이름(없으면 영문 이름)' },
    nameEn: { type: 'string' },
    title: { type: 'string', description: '직위·직급 (예: 사무관, 팀장, 대표이사 / 공학박사)' },
    dept: { type: 'string', description: '부서 (상위 조직부터 공백으로 이어서, 예: 해양정책실 해양개발과)' },
    org: { type: 'string', description: '기관·회사 이름 (로고 글자 포함, 정식 이름)' },
    cat: { type: 'string', enum: CATS },
    mobile: { type: 'string' }, tel: { type: 'string' }, fax: { type: 'string' },
    email: { type: 'string' }, addr: { type: 'string', description: '주소 (본사·지사가 여럿이면 첫 번째)' }, web: { type: 'string' },
    box: { type: 'array', items: { type: 'number' }, description: '사진 안 명함 영역 [왼쪽, 위, 오른쪽, 아래] 0~1 비율. 명함이 사진을 꽉 채우면 [0,0,1,1]' },
  },
};
const PROMPT = '이 사진은 받은 명함이에요. 명함에 적힌 정보를 그대로 옮겨 주세요. 없는 칸은 빈 글자로 두고, 추측해서 채우지 마세요.\n' +
  '전화번호는 010-1234-5678 처럼 하이픈으로. 국가번호 +82 10 은 010 으로 바꿔 주세요.\n' +
  '분류(cat): 중앙부처(해양수산부·국토교통부·행정안전부 같은 ○○부·처·위원회 본부), 소속·산하기관(국립해양조사원·국토지리정보원·○○청·지방해양수산청 같은 정부 소속기관), ' +
  '지자체(시·도·군·구청), 공공기관(공단·공사·진흥원·정부출연연구원·협회·재단·대학), 기업(민간 회사), 사내(올포랜드), 기타.';

/* Claude API — Apps Script 에는 SDK 가 없어 HTTP 로 부른다. 결과는 JSON 스키마로 받는다 */
function claude_(b64, key) {
  const body = {
    model: MODEL, max_tokens: 4000,
    output_config: { effort: 'low', format: { type: 'json_schema', schema: SCHEMA } },
    fallbacks: 'default',
    messages: [{ role: 'user', content: [
      { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: b64 } },
      { type: 'text', text: PROMPT },
    ] }],
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body),
    headers: { 'x-api-key': key, 'anthropic-version': '2026-10-21', 'anthropic-beta': 'server-side-fallback-2026-07-01' },
  });
  const code = res.getResponseCode(), j = JSON.parse(res.getContentText() || '{}');
  if (code !== 200) throw new Error('Claude ' + code + ' ' + ((j.error && j.error.message) || ''));
  if (j.stop_reason === 'refusal') throw new Error('Claude 가 이 사진을 처리하지 않았어요');
  const t = (j.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  const r = JSON.parse(t);
  if (!r.isCard && !r.name) throw new Error('명함을 찾지 못했어요');
  delete r.isCard;
  if (!Array.isArray(r.box) || r.box.length !== 4) r.box = [0, 0, 1, 1];
  return r;
}

/* 구글 드라이브 OCR: 사진을 구글 문서로 바꾸면 글자를 읽어 준다 → 글자만 꺼내고 문서는 지운다 */
function ocr_(b64) {
  const boundary = 'cardbook' + Date.now(), meta = JSON.stringify({ name: 'card-ocr-' + Date.now(), mimeType: 'application/vnd.google-apps.document' });
  const pre = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n--' + boundary + '\r\nContent-Type: image/jpeg\r\n\r\n';
  const bytes = [].concat(Utilities.newBlob(pre).getBytes(), Utilities.base64Decode(b64), Utilities.newBlob('\r\n--' + boundary + '--').getBytes());
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=ko', {
    method: 'post', contentType: 'multipart/related; boundary=' + boundary, payload: bytes, muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  });
  if (res.getResponseCode() !== 200) throw new Error('글자 인식(OCR)에 실패했어요 (' + res.getResponseCode() + ')');
  const id = JSON.parse(res.getContentText()).id;
  try { return DocumentApp.openById(id).getBody().getText(); } finally { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  // 아래 줄은 권한(드라이브·문서) 요청을 위해 남겨 둔다: DriveApp.createFile, DocumentApp.create
}

/* OCR 글자 → 칸 나누기 (규칙). AI 보다 부정확하므로 앱에서 '확인 필요'로 표시된다 */
const TITLES = '대표이사|대표|회장|부회장|사장|부사장|전무이사|전무|상무이사|상무|이사|본부장|실장|센터장|국장|부장|차장|과장|팀장|파트장|소장|원장|부원장|단장|위원장|위원|책임연구원|선임연구원|수석연구원|연구원|연구관|연구사|연구위원|책임|선임|수석|주임|대리|사원|매니저|사무관|서기관|주무관|부이사관|이사관|행정관|기술서기관|공업사무관|해양수산사무관|교수|부교수|조교수|박사|기자|변호사|프로|PM|Manager|Director|CEO|CTO|COO';
const TITLE_RE = new RegExp('(?<![가-힣])(' + TITLES + ')');   // '조사원'의 '사원'처럼 낱말 안에서 잡지 않게
function parse_(text) {
  const lines = String(text || '').split(/\n+/).map(s => s.replace(/\s+/g, ' ').trim()).filter(Boolean);
  const all = lines.join('\n'), r = { name: '', nameEn: '', title: '', dept: '', org: '', cat: '', mobile: '', tel: '', fax: '', email: '', addr: '', web: '', box: [0, 0, 1, 1] };
  r.email = (all.match(/[\w.+-]+@[\w-]+(\.[\w-]+)+/) || [''])[0];
  r.web = (all.match(/(https?:\/\/)?www\.[\w.-]+\.[a-z]{2,}(\/\S*)?/i) || [''])[0];
  // 한 줄에 'T. 051-… F. 051-…' 처럼 여러 번호가 있어도 번호마다 앞의 표시(M·T·F)로 가른다
  const PH = /(m\.?|mobile|h\.?p\.?|c\.?p\.?|휴대폰?|핸드폰|t\.?|tel\.?|전화|d\.?|direct|f\.?|fax\.?|팩스)?[\s.:)\]]*((?:\+?82[\s.-]?0?\d{1,2}|0\d{1,3})[\s.)-]*\d{3,4}[\s.-]*\d{4})/gi;
  let m; while ((m = PH.exec(all))) {
    const lab = (m[1] || '').toLowerCase(), v = m[2], mob = /^0?1[016789]/.test(v.replace(/[^\d]/g, '').replace(/^82/, '0'));
    if (/^(f|fax|팩스)/.test(lab)) { if (!r.fax) r.fax = v; }
    else if (/^(m|h|c|휴|핸)/.test(lab) || mob) { if (!r.mobile) r.mobile = v; }
    else if (!r.tel) r.tel = v;
  }
  r.addr = lines.find(l => /(서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충청|충북|충남|전라|전북|전남|경상|경북|경남|제주)/.test(l) && /(로|길|동|번지)\s*\d/.test(l)) || '';
  // 이름: 성 + 1~3글자 한글이 따로 있는 줄 (띄어 쓴 '이 희 용' 도)
  const SUR = '김이박최정강조윤장임한오서신권황안송류유전홍고문양손배백허남심노하곽성차주우구민진나지엄채원천방공현함변염여추도소석선설마길연위표명기반왕금옥육인맹제모탁국어은편용예봉경';
  for (const l of lines) {
    const t = l.replace(TITLE_RE, ' ').replace(/[A-Za-z].*$/, '').trim(), s = t.replace(/\s/g, '');
    if (/^[가-힣]{2,4}$/.test(s) && SUR.indexOf(s[0]) >= 0 && (t.length <= 7)) { r.name = s; const en = l.match(/[A-Z][a-z]+(\s?[A-Z][a-z-]+)+/); if (en) r.nameEn = en[0]; break; }
  }
  const tl = lines.find(l => TITLE_RE.test(l)); if (tl) r.title = (tl.match(TITLE_RE) || [''])[0];
  // 기관 줄 후보 중 더 구체적인 것(소속기관·회사)을 먼저: '해양수산부 / 국립해양조사원' 이면 조사원
  const orgs = lines.filter(l => /(\(주\)|㈜|주식회사|부$|조사원|정보원|공단|공사|진흥원|연구원$|협회|재단|청$|시청|도청|구청|군청|대학교)/.test(l) && !/@|\d{3}/.test(l));
  r.org = (orgs.find(l => !/부$/.test(l)) || orgs[0] || '').replace(/\s*(대표|Tel).*$/i, '');
  r.dept = (lines.find(l => /(과|팀|실|센터|본부|부서|국|사업부|연구소)$/.test(l.replace(TITLE_RE, '').trim()) && l !== r.org && !/@|\d{3}/.test(l)) || '').replace(TITLE_RE, '').trim();
  if (!r.org && r.email) r.org = r.email.split('@')[1].split('.')[0];
  if (!r.name) r.name = '(이름 확인 필요)';
  return r;
}

/* 이미 넣어 둔 명함에서 배우기: 같은 회사 메일 주소(@rian.kr)면 그 명함의 기관 이름·분류를,
   같은 기관이면 그 분류를 그대로 쓴다. 한 번 고쳐 두면 다음 명함부터 맞게 들어간다.
   korea.kr·gmail 처럼 여러 기관이 함께 쓰는 주소는 기관을 가르지 못하므로 빼고 본다 */
const SHARED_MAIL = /@(korea\.kr|gmail\.com|naver\.com|daum\.net|hanmail\.net|nate\.com|kakao\.com|outlook\.com|hotmail\.com)$/i;
function learn_(r) {
  const cards = list_().filter(c => c.checked);   // 사람이 확인한 명함만 믿는다
  const dom = r.email && !SHARED_MAIL.test(r.email) ? r.email.split('@')[1].toLowerCase() : '';
  const same = dom && cards.find(c => c.email && c.email.toLowerCase().endsWith('@' + dom) && c.org);
  if (same) { r.org = same.org; r.cat = same.cat; return; }
  const key = s => String(s || '').replace(/\s|\(주\)|㈜|주식회사/g, '');
  const o = r.org && cards.find(c => key(c.org) === key(r.org));
  if (o) { r.org = o.org; r.cat = o.cat; }
}

/* 분류 규칙 — AI 가 없거나 '기타'라고 했을 때 */
const MINISTRY = /(기획재정부|교육부|과학기술정보통신부|외교부|통일부|법무부|국방부|행정안전부|국가보훈부|문화체육관광부|농림축산식품부|산업통상자원부|보건복지부|환경부|기후에너지환경부|고용노동부|성평등가족부|여성가족부|국토교통부|해양수산부|중소벤처기업부|인사혁신처|법제처|식품의약품안전처|국무조정실|위원회)/;
function classify_(org, email, dept) {
  org = String(org || ''); email = String(email || '').toLowerCase();
  if (MY_COMPANY.test(org) || /allforland/.test(email)) return '사내';
  if (/(시청|도청|군청|구청|광역시|특별시|특별자치|^[가-힣]{2,4}(도|시|군|구)$)/.test(org)) return '지자체';
  if (/(국립|조사원|정보원|지방해양수산청|해양경찰|[가-힣]청$|관리단|사무소$)/.test(org)) return '소속·산하기관';
  if (MINISTRY.test(org)) return '중앙부처';
  if (/(공단|공사|진흥원|연구원|협회|재단|진흥회|대학교|대학원|\.or\.kr|\.re\.kr)/.test(org) || /\.(or|re|ac)\.kr$/.test(email)) return '공공기관';
  if (/(\(주\)|㈜|주식회사|Inc|Co\.|Corp|Ltd)/i.test(org) || /\.(co\.kr|com|ai|io|net)$/.test(email)) return '기업';
  if (/korea\.kr$/.test(email)) return '소속·산하기관';
  return '기타';
}
function phone_(v) {
  let d = String(v || '').replace(/[^\d+]/g, ''); if (!d) return '';
  d = d.replace(/^\+?82/, '0').replace(/^00/, '0');
  if (/^050\d{8}$/.test(d)) return d.replace(/^(050\d)(\d{3})(\d{4})$/, '$1-$2-$3');   // 0504-078-1340 같은 안심번호
  if (/^02\d{7,8}$/.test(d)) return d.replace(/^(02)(\d{3,4})(\d{4})$/, '$1-$2-$3');
  if (/^0\d{9,10}$/.test(d)) return d.replace(/^(0\d{2})(\d{3,4})(\d{4})$/, '$1-$2-$3');
  if (/^1\d{3}\d{4}$/.test(d)) return d.replace(/^(\d{4})(\d{4})$/, '$1-$2');
  return String(v).trim();
}

/* 편집기에서 시험: 드라이브의 명함 사진 한 장으로 읽기만 해 본다 (재배포 필요 없음) */
function testScan() {
  const f = DriveApp.searchFiles("mimeType contains 'image/' and trashed = false").next();
  console.log(f.getName(), JSON.stringify(scan_('data:image/jpeg;base64,' + Utilities.base64Encode(f.getBlob().getBytes())), null, 1));
}
