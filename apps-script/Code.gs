/** @OnlyCurrentDoc */
/**
 * GHU 성과지표 대시보드 ↔ 구글 스프레드시트 연동
 *
 * 대시보드의 "등록" 버튼을 누르면 입력한 실적·목표가 이 스크립트로 전송되어
 * "등록" 시트에 한 값당 한 줄씩 쌓입니다. 관리자는 대시보드의 관리자 화면에서
 * 관리자 키를 넣고 "시트에서 불러오기"를 누르면 쌓인 값을 가져와 조정합니다.
 *
 * 설치 방법은 같은 폴더의 설치안내.md를 보세요.
 */

var SHEET_NAME = '등록';
var HEADER = ['등록 시각', '등록 번호', '부서', '작성자', '지표코드', '지표명', '연도', '구분', '값', '단위', '브라우저 제출 시각'];
var MAX_ITEMS = 2000;

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.kind !== 'kpi-submit') return out({ ok: false, error: '알 수 없는 요청입니다.' });
    var dept = clean(body.dept, 40), who = clean(body.who, 40), at = clean(body.at, 30);
    if (!dept) return out({ ok: false, error: '부서가 비어 있습니다.' });
    var items = Array.isArray(body.items) ? body.items.slice(0, MAX_ITEMS) : [];
    var rows = [], now = new Date(), id = Utilities.formatDate(now, 'Asia/Seoul', 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 9000 + 1000);
    items.forEach(function (it) {
      var c = String(it.c || ''), y = String(it.y || ''), k = it.k === 't' ? '목표' : (it.k === 'a' ? '실적' : ''), v = Number(it.v);
      if (!/^[SMILEG]\d-\d-\d$/.test(c) || !/^20\d\d$/.test(y) || !k || !isFinite(v)) return;
      rows.push([now, id, dept, who, c, clean(it.n, 80), y, k, v, clean(it.u, 10), at]);
    });
    if (!rows.length) return out({ ok: false, error: '등록할 숫자가 없습니다.' });
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var sh = sheet();
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADER.length).setValues(rows);
    } finally { lock.releaseLock(); }
    return out({ ok: true, id: id, n: rows.length });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

// 관리자 화면의 "시트에서 불러오기": ?action=list&key=관리자키
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'ping') return out({ ok: true, sheet: SHEET_NAME });
  if (p.action !== 'list') return out({ ok: false, error: '알 수 없는 요청입니다.' });
  var key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!key || p.key !== key) return out({ ok: false, error: '관리자 키가 맞지 않습니다.' });
  var sh = sheet(), n = sh.getLastRow() - 1;
  var data = n > 0 ? sh.getRange(2, 1, n, HEADER.length).getValues() : [];
  var rows = data.map(function (r) {
    return { t: r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]), id: String(r[1]), dept: String(r[2]), who: String(r[3]), c: String(r[4]), y: String(r[6]), k: r[7] === '목표' ? 't' : 'a', v: Number(r[8]) };
  });
  return out({ ok: true, rows: rows });
}

function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
  }
  return sh;
}

function clean(s, n) { return String(s == null ? '' : s).replace(/[\r\n\t]/g, ' ').replace(/^[=+\-@]/, "'$&").slice(0, n).trim(); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// 처음 한 번 실행: 시트 머리글을 만들고 관리자 키를 정합니다(키는 원하는 값으로 바꾸세요).
function setup() {
  sheet();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_KEY')) props.setProperty('ADMIN_KEY', Utilities.getUuid().slice(0, 8));
  Logger.log('관리자 키: ' + props.getProperty('ADMIN_KEY'));
}
