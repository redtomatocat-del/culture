/**
 * BEYOND THE SCREEN — 생기부 기록 워크북 백엔드 (HTML 워크북 패턴 · 명단 없는 버전)
 * ────────────────────────────────────────────────────
 * 1. 학생은 HTML 워크북에서 작성 → POST 제출 (학년·반·번호·이름 직접 입력)
 * 2. 'BTS_제출' 시트에 upsert (학년-반-번호-이름 키, 빈 값은 기존 값 보존 = 병합 저장)
 * 3. 선생님은 시트 메뉴에서 마스터 탭·학생 카드 탭·세특 초안 사용
 * 4. 보드는 같은 /exec URL의 GET 응답 (HTML 파일명: board_sa)
 * 5. 메뉴의 [📋 성찰 설문지 생성]으로 구글 폼을 1회 자동 생성
 * ────────────────────────────────────────────────────
 */

// ─── 상수 ───
const SHEET_NAME = 'BTS_제출';
const MASTER_NAME = '📋 학생 현황';
const GUIDE_NAME = '📖 사용 안내';

const HEADERS = [
  '제출일시', '학년', '반', '번호', '이름',
  '전공', '제목', '이론렌즈', '네열쇠',
  '활동개요', '세션장면', '문화원형', '이론접근', '실제한일', '질문주장', '생각변화', '나의생각', '후속연구'
];
const CONTENT_START = 5; // 콘텐츠 칸 시작 인덱스 (제목부터) — 병합 저장 대상

const GUIDANCE = {
  overview: '이틀 활동 개요 + 테이블·모둠에서 내 역할 (캐릭터가 아니라 나).',
  scene:    '누가·무엇을·주사위 결과·그래서 테이블에 벌어진 일. 실패·부분 성공 장면 환영.',
  archetype: '서구의 원형(원탁 서사·영웅담) ↔ 우리 문화의 원형(마당놀이·판소리·이야기판) 비교: 닮은 점 + 다른 점.',
  theory:    '이론 렌즈 개념어 2개 이상으로 장면 해석 + 경험경제·연대의 역설 조사를 이론에 연결.',
  essay:    '지식 질문 + 주장 + 인용한 제시문·체험 근거 요약.',
  followup: '후속 탐구 예고: 주제 → 이유 → 탐구 질문 → 자료·방법.',
  change:   '수업 전 생각 ↔ 지금 생각 대비 + 변화의 계기.',
  work:     'Day 1/2로 나눠 동사 중심 행동 기록. 세특의 사실 근거.',
  thought:  '형식 없는 자기 생각·경험담. 솔직할수록 좋음.',
};

const SECTIONS = [
  { key: 'overview', col: '활동개요',   label: '① 활동 개요와 내 역할',   row: 9  },
  { key: 'scene',    col: '세션장면',   label: '② 기억에 남는 세션 장면', row: 13 },
  { key: 'archetype', col: '문화원형', label: '③ 문화 원형으로 비교하기', row: 17 },
  { key: 'theory',    col: '이론접근', label: '④ 사회를 보는 이론과 접근', row: 21 },
  { key: 'work',     col: '실제한일',   label: '⑤ 실제로 한 일',          row: 25 },
  { key: 'essay',    col: '질문주장',   label: '⑥ 지식 질문과 나의 주장', row: 29 },
  { key: 'change',   col: '생각변화',   label: '⑦ 생각의 변화 (전→후)',  row: 33 },
  { key: 'thought',  col: '나의생각',   label: '⑧ 나의 생각과 경험담',    row: 37 },
  { key: 'followup', col: '후속연구',   label: '⑨ 후속 탐구 예고',        row: 41 },
];
const PROGRESS_KEYS = ['overview','scene','archetype','theory','work','essay','change','thought','followup'];
const SETUK_LABEL_ROW = 45;
const SETUK_BODY_ROW = 46;

const COLOR_HEADER = '#3E1E17';
const COLOR_SECTION = '#7A3B2E';
const COLOR_TOPIC = '#A05A3F';
const COLOR_GUIDE_BG = '#FAEEDA';
const COLOR_GUIDE_TEXT = '#854F0B';
const COLOR_SETUK_BG = '#FFF7ED';
const COLOR_SETUK_BODY = '#FAEEDA';
const COLOR_SETUK_TEXT = '#C2410C';
const TAB_PALETTE = ['#E4C9A8','#B5D4F4','#F5C4B3','#E9D8A6','#CDB4DB','#A9DEB8','#F4D58A','#F5B7C4','#C9CBA3','#B8C0FF','#FFC9B9','#D8E2DC','#9FC5E1'];
function tabColorFor(ban){ const n = parseInt(ban, 10); return TAB_PALETTE[(isNaN(n) ? 0 : n) % TAB_PALETTE.length]; }


// ─── 시트 보장 ───
function getSubmitSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setBackground(COLOR_HEADER).setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
    sh.setColumnWidth(1, 130);
    [2,3,4].forEach(c => sh.setColumnWidth(c, 50));
    sh.setColumnWidth(5, 90);
    sh.setColumnWidth(6, 250);
    [7,8].forEach(c => sh.setColumnWidth(c, 100));
    for (let c = 9; c <= HEADERS.length; c++) sh.setColumnWidth(c, 320);
    try { sh.setTabColor('#7A3B2E'); } catch (e) {}
  }
  return sh;
}


// ─── 라우팅 ───
function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('board_sa');
  tpl.banParam = (e && e.parameter && e.parameter.ban) ? String(e.parameter.ban).trim() : '';
  return tpl.evaluate()
    .setTitle('Beyond the Screen — 학급 보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'submit') return saveSubmission(data);
    return jsonOut({ ok: false, msg: 'Unknown action' });
  } catch (err) {
    return jsonOut({ ok: false, msg: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function s(v) { return v == null ? '' : String(v); }
function pad2(v){ v = String(v == null ? '' : v).trim(); return /^[0-9]$/.test(v) ? '0' + v : v; }


// ─── 제출 저장 (병합 upsert) ───
function saveSubmission(d) {
  d.num = pad2(d.num); d.name = String(d.name == null ? '' : d.name).trim();
  const sh = getSubmitSheet();
  const ts = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm:ss');
  const row = [
    ts,
    s(d.grade), s(d.ban), s(d.num), s(d.name),
    s(d.major), s(d.title), s(d.lens1), s(d.keyword),
    s(d.overview), s(d.scene), s(d.archetype), s(d.theory), s(d.work), s(d.essay), s(d.change), s(d.thought), s(d.followup)
  ];

  const key = [s(d.grade), s(d.ban), s(d.num), s(d.name)].join('|');
  const all = sh.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < all.length; i++) {
    const k = [all[i][1], all[i][2], pad2(all[i][3]), String(all[i][4]).trim()].join('|');
    if (k === key) { foundRow = i + 1; break; }
  }
  var resultMsg;
  if (foundRow > 0) {
    const prev = all[foundRow - 1];
    for (let c = CONTENT_START; c < HEADERS.length; c++) {
      if (String(row[c]).trim() === '' && String(prev[c]).trim() !== '') {
        row[c] = prev[c];
      }
    }
    sh.getRange(foundRow, 1, 1, HEADERS.length).setValues([row]);
    resultMsg = '제출 갱신됨(병합)';
  } else {
    sh.appendRow(row);
    resultMsg = '새 제출 저장됨';
  }

  try {
    var student = { grade: s(d.grade), ban: s(d.ban), num: s(d.num), name: s(d.name) };
    buildStudentTab(student, rowToRec(row));
  } catch (err) {}

  return jsonOut({ ok: true, msg: resultMsg });
}


// ─── 메뉴 ───
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎲 Beyond the Screen')
    .addItem('🔄 학생 현황 마스터 탭 갱신', 'updateMasterTab')
    .addSeparator()
    .addItem('🆕 학생 카드 탭 전체 생성 (제출 기반)', 'generateAllStudentTabs')
    .addItem('🃏 선택 학생 카드 탭 새로 만들기', 'rebuildSelectedTab')
    .addSeparator()
    .addItem('✍ 세특 초안 일괄 생성 (모든 제출 학생)', 'generateAllSetukDrafts')
    .addItem('✍ 선택 학생만 세특 초안 생성', 'generateSetukForCurrent')
    .addSeparator()
    .addItem('📋 성찰 설문지(구글 폼) 생성 — 1회만', 'createSurveyForm')
    .addSeparator()
    .addItem('🔀 탭 순서 정리 (학년 → 반 → 번호)', 'arrangeTabsByOrder')
    .addItem('🖥 학급 보드 열기 (새 탭)', 'openBoard')
    .addItem('📄 선택 학생 → Google Docs', 'exportCurrentToDoc')
    .addSeparator()
    .addItem('⚠ 제출 데이터 모두 초기화 (위험)', 'resetAllSubmissions')
    .addToUi();
}


// ─── 제출 row → 학생 레코드 ───
function rowToRec(row) {
  const rec = {};
  HEADERS.forEach((h, i) => { rec[h] = row[i]; });
  return {
    ts: rec['제출일시'],
    grade: s(rec['학년']), ban: s(rec['반']), num: pad2(rec['번호']), name: String(rec['이름'] == null ? '' : rec['이름']).trim(),
    major: s(rec['전공']), title: s(rec['제목']), lens1: s(rec['이론렌즈']), keyword: s(rec['네열쇠']),
    overview: s(rec['활동개요']), scene: s(rec['세션장면']), archetype: s(rec['문화원형']), theory: s(rec['이론접근']),
    work: s(rec['실제한일']), essay: s(rec['질문주장']), change: s(rec['생각변화']), thought: s(rec['나의생각']), followup: s(rec['후속연구'])
  };
}

function getAllSubmissions() {
  const sh = getSubmitSheet();
  const all = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < all.length; i++) {
    const r = rowToRec(all[i]);
    if (!r.name) continue;
    const k = [r.grade, r.ban, r.num, r.name].join('|');
    map[k] = { rec: r, rowNum: i + 1 };
  }
  return map;
}

function submittedStudents() {
  const subs = getAllSubmissions();
  const list = Object.keys(subs).map(k => {
    const r = subs[k].rec;
    return { grade: r.grade, ban: r.ban, num: r.num, name: r.name, key: k };
  });
  list.sort((a, b) => {
    if (+a.grade !== +b.grade) return +a.grade - +b.grade;
    if (+a.ban !== +b.ban) return +a.ban - +b.ban;
    return +a.num - +b.num;
  });
  return list;
}

function filledCount(rec) {
  return PROGRESS_KEYS.filter(k => String(rec[k] || '').trim().length >= 5).length;
}


// ─── 학생 탭 생성 ───
function studentTabName(stu) {
  return stu.name + '(' + stu.grade + '-' + stu.ban + '-' + stu.num + ')';
}

function generateAllStudentTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const subs = getAllSubmissions();
  const list = submittedStudents();
  if (list.length === 0) { ui.alert('아직 제출한 학생이 없습니다.'); return; }
  const r = ui.alert('제출 학생 ' + list.length + '명의 카드 탭을 생성/갱신합니다. 진행?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  let created = 0, updated = 0;
  list.forEach(stu => {
    const rec = subs[stu.key].rec;
    const existed = !!ss.getSheetByName(studentTabName(stu));
    buildStudentTab(stu, rec);
    if (existed) updated++; else created++;
  });

  try { ensureGuideTab(); } catch (e) {}
  try { updateMasterTab(); } catch (e) {}
  try { arrangeTabsByOrder(); } catch (e) {}

  ui.alert('완료. 새로 생성 ' + created + '개, 기존 갱신 ' + updated + '개.');
}

function rebuildSelectedTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getActiveSheet();
  const m = sh.getName().match(/^(.+?)\((\d+)-(\d+)-(\d+)\)$/);
  if (!m) { ui.alert('학생 탭에서 실행하세요 (이름(학년-반-번호) 형식)'); return; }
  const [_, name, grade, ban, num] = m;
  const subs = getAllSubmissions();
  const key = [grade, ban, pad2(num), name].join('|');
  if (!subs[key]) { ui.alert('이 학생의 제출 데이터가 없습니다.'); return; }
  buildStudentTab({ grade: grade, ban: ban, num: pad2(num), name: name }, subs[key].rec);
  ui.alert('카드 탭 재생성 완료.');
}


function buildStudentTab(student, rec) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabName = studentTabName(student);
  let sh = ss.getSheetByName(tabName);
  if (sh) {
    sh.clear();
    try { sh.getRange(1, 1, sh.getMaxRows(), 2).breakApart(); } catch (e) {}
  } else {
    sh = ss.insertSheet(tabName);
  }

  const content = new Array(SETUK_BODY_ROW).fill(['', '']);
  content[0] = [student.name + '(' + student.grade + '-' + student.ban + '-' + student.num + ') — Beyond the Screen 기록', ''];
  content[1] = [student.grade + '학년 ' + student.ban + '반 ' + student.num + '번  ·  제출: ' + (rec.ts || '미제출'), ''];
  content[2] = ['▣ 기록 개요', ''];
  content[3] = ['제목', rec.title || ''];
  content[4] = ['희망 전공·계열', rec.major || '-'];
  content[5] = ['렌즈 · 열쇠', (rec.lens1 || '-') + '  ·  ' + (rec.keyword || '-')];

  SECTIONS.forEach(sec => {
    content[sec.row - 1] = ['▣ ' + sec.label, ''];
    content[sec.row] = ['(안내) ' + GUIDANCE[sec.key], ''];
    content[sec.row + 1] = [rec[sec.key] || '', ''];
  });

  content[SETUK_LABEL_ROW - 1] = ['📋 자동 생성 세특 초안 (제출 시 자동 · 다듬어서 NEIS에)', ''];
  content[SETUK_BODY_ROW - 1] = [generateSetukDraft(rec), ''];

  sh.getRange(1, 1, content.length, 2).setValues(content);

  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 600);

  sh.getRange(1, 1, 1, 2).merge()
    .setBackground(COLOR_HEADER).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sh.setRowHeight(1, 36);

  sh.getRange(2, 1, 1, 2).merge()
    .setBackground('#F2ECE2').setFontColor('#3F3F3A')
    .setFontSize(10.5).setHorizontalAlignment('center').setFontFamily('맑은 고딕');

  sh.getRange(3, 1, 1, 2).merge()
    .setBackground(COLOR_TOPIC).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('left');
  sh.setRowHeight(3, 28);

  sh.getRange(4, 1, 3, 1).setBackground('#EDE6D8').setFontWeight('bold').setFontColor('#444441').setFontSize(11);
  sh.getRange(4, 2, 3, 1).setBackground('#FFFFFF').setFontSize(11).setVerticalAlignment('middle').setWrap(true);
  [4,5,6].forEach(r => sh.setRowHeight(r, 30));

  sh.setRowHeight(8, 8);

  SECTIONS.forEach(sec => {
    sh.getRange(sec.row, 1, 1, 2).merge()
      .setBackground(COLOR_SECTION).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('left');
    sh.setRowHeight(sec.row, 28);

    sh.getRange(sec.row + 1, 1, 1, 2).merge()
      .setBackground(COLOR_GUIDE_BG).setFontColor(COLOR_GUIDE_TEXT)
      .setFontSize(10.5).setFontStyle('italic').setWrap(true)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.setRowHeight(sec.row + 1, 30);

    sh.getRange(sec.row + 2, 1, 1, 2).merge()
      .setBackground('#FFFFFF').setFontColor('#221A15').setFontSize(12)
      .setWrap(true).setVerticalAlignment('top').setFontFamily('맑은 고딕');
    sh.setRowHeight(sec.row + 2, (sec.key === 'archetype' || sec.key === 'theory') ? 130 : 90);

    sh.setRowHeight(sec.row + 3, 8);
  });

  sh.getRange(SETUK_LABEL_ROW, 1, 1, 2).merge()
    .setBackground(COLOR_SETUK_BG).setFontColor(COLOR_SETUK_TEXT)
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left');
  sh.setRowHeight(SETUK_LABEL_ROW, 26);
  sh.getRange(SETUK_BODY_ROW, 1, 1, 2).merge()
    .setBackground(COLOR_SETUK_BODY).setFontColor('#412402').setFontSize(11)
    .setWrap(true).setVerticalAlignment('top');
  sh.setRowHeight(SETUK_BODY_ROW, 100);

  sh.setHiddenGridlines(true);
  sh.setFrozenRows(2);
  if (sh.getMaxColumns() > 2) {
    sh.hideColumns(3, sh.getMaxColumns() - 2);
  }

  try { sh.setTabColor(tabColorFor(student.ban)); } catch (e) {}
}


// ─── 마스터 탭 (제출 기반) ───
function updateMasterTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MASTER_NAME);
  if (!sh) sh = ss.insertSheet(MASTER_NAME, 0);
  sh.clear();
  try { sh.getRange(1, 1, sh.getMaxRows(), 11).breakApart(); } catch (e) {}

  const header = ['학년', '반', '번호', '이름', '전공', '이론 렌즈', '네 열쇠', '제목', '진척도', '세특 초안 미리보기'];
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setBackground(COLOR_HEADER).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);

  const subs = getAllSubmissions();
  const list = submittedStudents();
  const rows = list.map(stu => {
    const rec = subs[stu.key].rec;
    const filled = filledCount(rec);
    const pct = Math.round(filled / PROGRESS_KEYS.length * 100);
    const card = ss.getSheetByName(studentTabName(stu));
    let setuk = '';
    if (card) setuk = String(card.getRange(SETUK_BODY_ROW, 1).getValue() || '');
    if (!setuk) setuk = generateSetukDraft(rec);
    if (setuk.length > 200) setuk = setuk.slice(0, 200) + '…';
    return [stu.grade, stu.ban, stu.num, stu.name, rec.major || '', rec.lens1 || '', rec.keyword || '', rec.title || '', pct + '%  (' + filled + '/9)', setuk];
  });

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    sh.getRange(2, 1, rows.length, header.length).setWrap(true).setVerticalAlignment('top');
  }

  sh.setColumnWidth(1, 45);
  sh.setColumnWidth(2, 45);
  sh.setColumnWidth(3, 50);
  sh.setColumnWidth(4, 85);
  sh.setColumnWidth(5, 130);
  sh.setColumnWidth(6, 100);
  sh.setColumnWidth(7, 200);
  sh.setColumnWidth(8, 90);
  sh.setColumnWidth(9, 240);
  sh.setColumnWidth(10, 340);

  for (let i = 0; i < rows.length; i++) {
    const m = String(rows[i][7]).match(/\((\d+)\/9\)/);
    if (m && parseInt(m[1], 10) < 4) {
      sh.getRange(i + 2, 1, 1, header.length).setBackground('#FFFBEB');
    }
  }

  try { sh.setTabColor('#4D7C0F'); } catch (e) {}
}


function ensureGuideTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(GUIDE_NAME)) return;
  const sh = ss.insertSheet(GUIDE_NAME, 1);
  const lines = [
    ['Beyond the Screen — 선생님 안내', ''],
    ['', ''],
    ['📌 흐름', ''],
    ['  1. 학생 → HTML 워크북에서 작성·제출 (학년·반·번호·이름 직접 입력)', ''],
    ['  2. 시트 → 자동으로 [BTS_제출] 탭에 row 추가/갱신 (병합 저장 — 빈 칸은 기존 값 보존)', ''],
    ['  3. 메뉴 → 마스터 탭 갱신 / 카드 탭 생성 / 세특 초안 생성', ''],
    ['  4. 성찰 설문지는 메뉴 [📋 성찰 설문지 생성]으로 1회 생성 (구글 폼 링크가 뜸)', ''],
    ['', ''],
    ['📋 구조', ''],
    ['  • 9칸: 활동개요·역할 → 세션 장면 → 문화 원형 비교 → 이론과 접근 → 실제 한 일 → 지식 질문·주장 → 생각 변화 → 나의 생각·경험담 → 후속 탐구 예고', ''],
    ['  • 마스터 탭에서 전공·이론 렌즈·네 열쇠 별로 필터 가능', ''],
    ['', ''],
    ['⚠ 주의', ''],
    ['  • 학생은 시트에 직접 접근하지 않아요. HTML 링크만 줍니다.', ''],
    ['  • 명단이 없는 버전 — 마스터 탭에는 제출한 학생만 표시. 미제출 파악은 명렬표와 대조.', ''],
    ['  • 같은 학년·반·번호·이름 재제출 = 같은 row 병합 갱신 (빈 칸은 기존 값 보존).', ''],
    ['  • 학생이 신원을 다르게 입력하면 별도 행 — 오타 행은 시트에서 직접 삭제.', ''],
  ];
  sh.getRange(1, 1, lines.length, 2).setValues(lines);
  sh.getRange(1, 1, 1, 2).merge()
    .setBackground(COLOR_HEADER).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sh.setRowHeight(1, 36);
  sh.setColumnWidth(1, 700);
  sh.setColumnWidth(2, 50);
  sh.getRange(1, 1, lines.length, 2).setFontFamily('맑은 고딕').setFontSize(11).setWrap(true).setVerticalAlignment('middle');
  sh.setHiddenGridlines(true);
  try { sh.setTabColor('#888780'); } catch (e) {}
}


// ─── 세특 초안 ───
function generateSetukDraft(rec) {
  const title = String(rec.title || '').trim();
  const lens1 = String(rec.lens1 || '').trim();
  const keyword = String(rec.keyword || '').trim();
  const overview = String(rec.overview || '').trim();
  const scene = String(rec.scene || '').trim();
  const archetype = String(rec.archetype || '').trim();
  const theory = String(rec.theory || '').trim();
  const major = String(rec.major || '').trim();
  const essay = String(rec.essay || '').trim();
  const followup = String(rec.followup || '').trim();
  const change = String(rec.change || '').trim();
  const work = String(rec.work || '').trim();
  const thought = String(rec.thought || '').trim();

  if (!title && !overview && !archetype && !theory) {
    return '(작성된 내용이 부족하여 자동 생성된 초안이 없습니다)';
  }

  const parts = [];
  if (title) {
    parts.push("영어·일반사회 융합수업(TRPG 체험)에서 '" + clip(title, 70) + "' 기록을 작성함.");
  }
  if (overview) parts.push('(역할) ' + clip(firstSent(overview), 160));
  if (scene) parts.push('(장면) ' + clip(firstSent(scene), 180));
  if (archetype) parts.push('(문화 원형 비교) ' + clip(firstSent(archetype), 180));
  if (theory) parts.push('(이론 접근)' + (lens1 ? ' [' + lens1 + '] ' : ' ') + clip(firstSent(theory), 200));
  if (essay) parts.push('(지식 질문·주장) ' + clip(firstSent(essay), 180));
  if (change) parts.push('(변화) ' + clip(firstSent(change), 160));
  if (thought) parts.push('(소감) ' + clip(firstSent(thought), 140));
  if (work) parts.push('(과정) ' + clip(firstSent(work), 180));
  if (followup) parts.push('(후속 예고) ' + clip(firstSent(followup), 150));
  parts.push('각 문화의 원형을 비교하는 시각으로 면대면 상호작용의 체험을 해석하고' + (major ? " 이를 '" + major + "' 분야의 관심으로 확장했으며" : '') + (keyword ? " '" + keyword + "'을(를) 자기 삶의 개념으로 끌어와" : '') + ' 경험을 지식으로 번역하는 융합적 탐구 태도가 드러남.');
  return parts.join(' ');
}

function firstSent(t) {
  const m = String(t).match(/[\s\S]{1,300}?[.!?。！？]/);
  return m ? m[0].trim() : String(t).slice(0, 300).trim();
}

function clip(str, n) {
  if (!str) return '';
  str = String(str).replace(/\s+/g, ' ').trim();
  return str.length > n ? str.slice(0, n) + '…' : str;
}


function generateAllSetukDrafts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const subs = getAllSubmissions();
  const eligible = submittedStudents().filter(stu => ss.getSheetByName(studentTabName(stu)));
  if (eligible.length === 0) {
    ui.alert('카드 탭이 생성된 학생이 없습니다. 먼저 [학생 카드 탭 전체 생성]을 실행하세요.');
    return;
  }
  const r = ui.alert(eligible.length + '명 학생의 세특 초안을 일괄 생성합니다. (각 학생 카드 탭의 주황 박스에 덮어씀)\n진행?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  let done = 0;
  eligible.forEach(stu => {
    const rec = subs[stu.key].rec;
    const sh = ss.getSheetByName(studentTabName(stu));
    if (!sh) return;
    sh.getRange(SETUK_BODY_ROW, 1).setValue(generateSetukDraft(rec));
    done++;
  });

  updateMasterTab();
  ui.alert('세특 초안 생성 완료: ' + done + '명');
}

function generateSetukForCurrent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getActiveSheet();
  const m = sh.getName().match(/^(.+?)\((\d+)-(\d+)-(\d+)\)$/);
  if (!m) { ui.alert('학생 카드 탭을 클릭한 뒤 실행하세요.'); return; }
  const [_, name, grade, ban, num] = m;
  const subs = getAllSubmissions();
  const key = [grade, ban, pad2(num), name].join('|');
  if (!subs[key]) { ui.alert('이 학생의 제출 데이터가 없습니다.'); return; }
  sh.getRange(SETUK_BODY_ROW, 1).setValue(generateSetukDraft(subs[key].rec));
  ui.alert('세특 초안 생성 완료.');
}


// ─── 성찰 설문지(구글 폼) 생성 ───
function createSurveyForm() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const existing = props.getProperty('BTS_FORM_URL');
  if (existing) {
    ui.alert('이미 생성된 설문지가 있습니다.\n\n학생 배포용: ' + existing + '\n\n(새로 만들려면 시트의 문서 속성에서 BTS_FORM_URL을 지우세요)');
    return;
  }

  const form = FormApp.create('Beyond the Screen 수업 성찰 설문');
  form.setDescription(
    '안녕하세요! 이 설문은 여러분이 이틀의 융합수업(Beyond the Screen)에서 경험하고 성장한 것을 학교생활기록부에 정확히 반영하기 위해 마련되었습니다.\n\n' +
    '이 설문은 점수에 반영되는 시험이 절대 아닙니다. 정답도 오답도 없습니다. 활동 과정, 산출물, 얻은 깨달음을 구체적이고 솔직하게 작성할수록 여러분의 노력과 성장이 뚜렷하게 드러납니다.\n\n' +
    '질문을 꼼꼼히 읽고, 이틀 동안의 자신을 천천히 돌아보며 편안한 마음으로 답변해 주세요.'
  );
  form.setCollectEmail(false);

  form.addTextItem().setTitle('1. 자신의 이름을 적어보세요.').setRequired(true);
  form.addTextItem().setTitle('2. 자신의 학년/반/번호를 적어보세요. (예시: 20305)').setRequired(true);

  // 사전-사후 그리드
  const gridStatements = [
    '나는 처음 만나는 사람과 얼굴을 맞대고 협력하는 것에 자신이 있었다.',
    '나는 규칙이 자유를 제한하기만 하는 것이 아니라, 무언가를 가능하게 할 수도 있다고 생각했다.',
    '나는 게임·놀이가 사회를 이해하는 진지한 도구가 될 수 있다고 생각했다.',
    '나는 사회 현상이나 일상의 장면을 이론(개념)의 렌즈로 해석할 수 있었다.',
    '나는 영어 지문을 읽을 때 단어의 뜻과 시대적 맥락을 추론하며 읽을 수 있었다.',
    '나는 다른 사람의 선택과 반응에 따라 나의 행동을 조율하는 상호작용에 능숙했다.',
    '나는 AI 시대에 인간에게 남는 고유한 역량이 무엇인지 나름의 답을 가지고 있었다.',
    '나는 예상 밖의 상황(실패·우연)을 회피하지 않고 이야기의 일부로 받아들일 수 있었다.'
  ];
  gridStatements.forEach(function(stmt, i) {
    form.addGridItem()
      .setTitle((i + 3) + '. ' + stmt)
      .setHelpText('위 설명에 대해 수업을 듣기 전과 후, 당신의 상태는 어떠했나요? 각 시점별로 해당하는 점수를 선택해주세요.\n(1: 전혀 그렇지 않다 / 2: 그렇지 않다 / 3: 보통이다 / 4: 그렇다 / 5: 매우 그렇다)')
      .setRows(['수업 전', '수업 후'])
      .setColumns(['1점', '2점', '3점', '4점', '5점'])
      .setRequired(true);
  });

  // 서술형 (생기부용)
  form.addParagraphTextItem().setTitle('11. 이틀 동안 어떤 활동들을 했는지 구체적으로 설명해주세요. 테이블과 모둠에서 자신이 맡았던 역할은 무엇이었나요? (예: 장면을 여는 질문, 출처 검증, 갈등 조율, 기록 등)');
  form.addParagraphTextItem().setTitle('12. 활동을 통해 얻은 구체적인 성과나 결과물은 무엇인가요? (예: 영문 지문 분석, 캐릭터·세션 기록, 이론 해석문, 경험경제 조사, 지식 질문 글쓰기 등)');
  form.addParagraphTextItem().setTitle('13. 활동 전과 비교하여 자신의 생각이나 태도에 어떤 변화가 있었나요? 특히 주체성·타자·상호작용·규칙 네 개념 중 무엇이 가장 와닿았는지, 그 이유와 함께 적어주세요.');
  form.addParagraphTextItem().setTitle('14. 자신의 전공·진로와 관련해서 새롭게 배우거나 연결한 점이 있다면 적어주세요. (Day 2에 세운 지식 질문을 포함해도 좋습니다)');
  form.addParagraphTextItem().setTitle('15. 활동 중 가장 기억에 남는 순간이나 어려웠던 점, 그리고 그것을 어떻게 극복했는지 자유롭게 이야기해주세요. (이를 통해 배우고 성장한 것을 설명)');
  form.addParagraphTextItem().setTitle('16. 이 수업과 연계하여 더 배우거나 탐구하고 싶은 후속 연구는 무엇인가요?\n\n다음 항목을 포함해 작성해주세요:\n1. 탐구하고 싶은 주제\n2. 이 주제를 선택한 이유\n3. 구체적으로 알아보고 싶은 점(탐구 질문)\n4. 자료를 찾고 연구를 진행할 방법 (예: 전공 서적, 뉴스·칼럼, 연구기관 보고서 등)\n5. 예상되는 최종 결과물의 형태');

  const editUrl = form.getEditUrl();
  const pubUrl = form.getPublishedUrl();
  props.setProperty('BTS_FORM_URL', pubUrl);
  props.setProperty('BTS_FORM_EDIT_URL', editUrl);

  ui.alert('✅ 설문지 생성 완료!\n\n학생 배포용 링크:\n' + pubUrl + '\n\n선생님 편집용 링크:\n' + editUrl + '\n\n(드라이브 최상위에 "Beyond the Screen 수업 성찰 설문"으로 저장됨. 응답은 폼의 [응답] 탭에서 시트로 연결하세요.)');
}


// ─── 탭 정렬 ───
function arrangeTabsByOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(MASTER_NAME);
  const guide = ss.getSheetByName(GUIDE_NAME);
  const submit = ss.getSheetByName(SHEET_NAME);

  let pos = 1;
  if (master) { try { master.activate(); ss.moveActiveSheet(pos++); master.setTabColor('#4D7C0F'); } catch (e) {} }
  if (submit) { try { submit.activate(); ss.moveActiveSheet(pos++); submit.setTabColor('#7A3B2E'); } catch (e) {} }
  if (guide) { try { guide.activate(); ss.moveActiveSheet(pos++); guide.setTabColor('#888780'); } catch (e) {} }

  const studentTabs = ss.getSheets().filter(sh => /\(\d+-\d+-\d+\)$/.test(sh.getName()));
  studentTabs.sort((a, b) => {
    const ma = a.getName().match(/\((\d+)-(\d+)-(\d+)\)$/);
    const mb = b.getName().match(/\((\d+)-(\d+)-(\d+)\)$/);
    if (!ma || !mb) return 0;
    if (+ma[1] !== +mb[1]) return +ma[1] - +mb[1];
    if (+ma[2] !== +mb[2]) return +ma[2] - +mb[2];
    return +ma[3] - +mb[3];
  });
  studentTabs.forEach((sh, i) => {
    try {
      sh.activate(); ss.moveActiveSheet(pos + i);
      const m = sh.getName().match(/\(\d+-(\d+)-\d+\)$/);
      if (m) sh.setTabColor(tabColorFor(m[1]));
    } catch (e) {}
  });
}


function openBoard() {
  const ui = SpreadsheetApp.getUi();
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  if (!url) { ui.alert('웹 앱이 아직 배포되지 않았습니다.'); return; }
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:13px;padding:6px;">새 탭에서 <b>보드</b>를 엽니다…<br>' +
    '안 열리면 <a href="' + url + '" target="_blank">여기 클릭</a></div>' +
    '<script>window.open(' + JSON.stringify(url) + ',"_blank");google.script.host.close();<\/script>'
  ).setWidth(330).setHeight(110);
  ui.showModalDialog(html, '보드 여는 중…');
}


// ─── Docs 내보내기 ───
function exportCurrentToDoc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getActiveSheet();
  const m = sh.getName().match(/^(.+?)\((\d+)-(\d+)-(\d+)\)$/);
  if (!m) { ui.alert('학생 카드 탭에서 실행하세요.'); return; }
  const [_, name, grade, ban, num] = m;
  const subs = getAllSubmissions();
  const key = [grade, ban, pad2(num), name].join('|');
  if (!subs[key]) { ui.alert('이 학생의 제출 데이터가 없습니다.'); return; }
  const rec = subs[key].rec;

  const title = 'BTS기록_' + name + '_' + Utilities.formatDate(new Date(), 'GMT+9', 'yyyyMMdd');
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  const H1 = DocumentApp.ParagraphHeading.HEADING1;
  const H2 = DocumentApp.ParagraphHeading.HEADING2;

  body.appendParagraph('Beyond the Screen 기록 — ' + name).setHeading(H1);
  body.appendParagraph(grade + '학년 ' + ban + '반 ' + num + '번');
  body.appendParagraph('제목: ' + (rec.title || '-'));
  body.appendParagraph('이론 렌즈: ' + (rec.lens1 || '-') + ' · 네 개의 열쇠: ' + (rec.keyword || '-'));
  body.appendParagraph('희망 전공·계열: ' + (rec.major || '-'));
  body.appendParagraph('').appendHorizontalRule();

  SECTIONS.forEach(sec => {
    body.appendParagraph(sec.label).setHeading(H2);
    body.appendParagraph(rec[sec.key] || '(미작성)');
  });

  body.appendParagraph('── 자동 생성 세특 초안 ──').setHeading(H2);
  body.appendParagraph(generateSetukDraft(rec));

  doc.saveAndClose();
  ui.alert('✅ Google Docs 생성 완료!\n\n' + doc.getUrl());
}


// ─── 위험: 초기화 ───
function resetAllSubmissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('⚠ [BTS_제출] 시트의 모든 학생 행을 지웁니다. 학생 카드 탭은 유지됩니다. 정말?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  const r2 = ui.alert('진짜요? 한 번 더 확인.', ui.ButtonSet.YES_NO);
  if (r2 !== ui.Button.YES) return;

  const sh = getSubmitSheet();
  const last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  updateMasterTab();
  ui.alert('제출 데이터 초기화 완료.');
}


// ─── 보드 API ───
function getBoardData() {
  const subs = getAllSubmissions();
  const out = { activities: [], updatedAt: Utilities.formatDate(new Date(), 'GMT+9', 'HH:mm:ss') };
  submittedStudents().forEach(stu => {
    const rec = subs[stu.key].rec;
    const filled = filledCount(rec);
    const progress = Math.round(filled / PROGRESS_KEYS.length * 100);
    out.activities.push({
      key: stu.key, rowNum: subs[stu.key].rowNum,
      grade: stu.grade, ban: stu.ban, num: stu.num, name: stu.name,
      title: rec.title || '(제목 미작성)',
      lens1: rec.lens1 || '',
      major: rec.major || '',
      keyword: rec.keyword || '',
      summary_preview: clip(rec.thought || rec.scene, 80),
      progress: progress,
      submitted: true
    });
  });
  return out;
}

function getStudentDetail(key) {
  const subs = getAllSubmissions();
  if (!subs[key]) return null;
  const rec = subs[key].rec;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stu = { grade: rec.grade, ban: rec.ban, num: rec.num, name: rec.name };
  let setuk = '';
  const card = ss.getSheetByName(studentTabName(stu));
  if (card) setuk = String(card.getRange(SETUK_BODY_ROW, 1).getValue() || '');
  if (!setuk) setuk = generateSetukDraft(rec);
  return Object.assign({}, rec, { setuk: setuk });
}
