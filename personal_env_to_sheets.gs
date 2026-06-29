/** 개인 온습도(여러 대) → Google Sheets 1개에 통합 저장 + 차트 데이터 제공
 *  - 시트 1개 / 웹앱 1개만 배포. 보드들은 각자 'name'(기기명)을 달아 전송.
 *  - doPost: 기록 (기기명 포함)
 *  - doGet?action=data : 차트 HTML이 읽을 데이터(JSONP)
 *
 *  배포:
 *   1) 구글시트 > 확장 프로그램 > Apps Script > 이 코드 붙여넣기 > 저장
 *   2) 배포 > 새 배포 > 유형 '웹 앱' (실행: 나, 액세스: 모든 사용자)
 *   3) 웹앱 URL 복사 → 모든 ino 의 GS_URL 에 같은 URL 사용 / 차트 HTML 에 1개만 등록
 *   4) 보관 자동삭제: installCleanup 1회 실행 (매일 새벽 3시 30일 초과분 삭제)
 *
 *  테스트: URL 끝에 ?action=data 붙여 열면 JSON 보임.
 */

const SHEET_NAME = '온습도';
const TZ = 'Asia/Seoul';
const RETAIN_DAYS = 30;   // 보관일수 (이보다 오래된 행 자동 삭제)

function sheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['측정시각(KST)', '기기명', '온도(℃)', '습도(%)', '체감(℃)']);
    sh.setFrozenRows(1);
    sh.getRange('A1:E1').setFontWeight('bold');
  }
  return sh;
}

function num_(v){ return (v === null || v === undefined || v === '') ? '' : Number(v); }

function doPost(e){
  try {
    const d = JSON.parse(e.postData.contents);
    const ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const name = (d.name || d.device || '온습도계');
    sheet_().appendRow([ts, String(name), num_(d.temperature), num_(d.humidity), num_(d.feels_like)]);
    return ContentService.createTextOutput('OK');
  } catch (err) {
    return ContentService.createTextOutput('ERR: ' + err);
  }
}

function readRecent_(n){
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - n + 1);
  const vals = sh.getRange(start, 1, last - start + 1, 5).getValues();
  return vals.map(function(r){
    return { ts: String(r[0]), name: String(r[1]), t: r[2], h: r[3], f: r[4] };
  });
}

function doGet(e){
  if (e && e.parameter && e.parameter.action === 'data') {
    const n = parseInt(e.parameter.n, 10) || 3000;
    const out = JSON.stringify({ ok: true, rows: readRecent_(n) });
    const cb = e.parameter.callback;
    if (cb) {
      return ContentService.createTextOutput(cb + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('OK (GET)');
}

/** 30일(RETAIN_DAYS) 지난 행 삭제 — 매일 자동 */
function cleanupOld(){
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const ts = sh.getRange(2, 1, last - 1, 1).getValues();
  const cutoff = new Date().getTime() - RETAIN_DAYS * 86400000;
  let del = 0;
  for (let i = 0; i < ts.length; i++) {
    const d = parseKst_(ts[i][0]);
    if (d && d.getTime() < cutoff) del++;
    else break;
  }
  if (del > 0) sh.deleteRows(2, del);
}

function parseKst_(v){
  if (v instanceof Date) return v;
  const m = String(v).match(/(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  if (!m) return null;
  return new Date(+m[1], +m[2]-1, +m[3], +m[4], +m[5]);
}

function installCleanup(){
  removeCleanup();
  ScriptApp.newTrigger('cleanupOld').timeBased().atHour(3).everyDays(1).inTimezone(TZ).create();
}
function removeCleanup(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'cleanupOld') ScriptApp.deleteTrigger(t);
  });
}
