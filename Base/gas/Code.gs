// ═══════════════════════════════════════════════════════════════
// E-PPA Apps Script 백엔드 — 기존 리드게이트 + 설정 공유 기능 통합
//
// 기존에 배포돼 있던 리드게이트 스크립트(요금분석 탭에서 "결과 확인하기"
// 제출 시 시트에 한 줄 기록)에, ④ 설정 탭에서 저장한 값을 모든 방문자에게
// 똑같이 적용하는 기능을 action 파라미터로 분기해 추가했다.
//
// - action 파라미터가 없는 요청(=summary.html의 submitLead()가 보내는 것과
//   동일) → 예전과 똑같이 시트에 리드 한 줄 기록
// - action=load / action=save → 설정값 조회/저장(PropertiesService, 별도
//   시트 불필요)
//
// [적용 방법]
// 1. 기존 EPPA Apps Script 프로젝트를 연다(SHEET_ID로 식별되는 바로 그 것)
// 2. 이 파일 내용 전체로 기존 코드를 덮어쓴다
// 3. 배포 → 배포 관리 → 기존 활성 배포의 연필(수정) 아이콘 클릭
//    → 버전: "새 버전" 선택 → 배포
//    (※ "새 배포"가 아니라 "새 버전"으로 기존 배포를 수정해야 URL이
//      그대로 유지되어 summary.html의 GATE_URL을 안 바꿔도 됨)
// 4. URL이 바뀌지 않으므로 summary.html의 SETTINGS_API_URL은 GATE_URL과
//    동일한 값으로 이미 설정해뒀다 — 추가로 할 일 없음
//
// [비밀번호]
// 설정 저장(action=save)은 pw 파라미터가 ADMIN_PW_HASH와 일치해야 허용된다.
// 이 값은 summary.html의 ADMIN_PW_HASH와 반드시 같아야 하며, 비밀번호를
// 바꾸면(현재 8890) summary.html에서 새로 계산한 해시로 두 곳 다 갱신해야 함.
// 클라이언트 코드에 이미 공개돼 있는 값이라 강한 보안은 아니고, 아무나
// 실수로 덮어쓰는 걸 막는 최소한의 장치임.
// ═══════════════════════════════════════════════════════════════

const SHEET_ID = '1-XdCFD3s5Gft3Ln0xBbQ0RljCkx8oWxZGs7cQj2rJ8A';
var ADMIN_PW_HASH = 1723927; // '8890' 해시 — summary.html의 ADMIN_PW_HASH와 동일하게 유지
var SETTINGS_PROP_KEY = 'eppa_admin_settings';

function doGet(e) {
  var p = e.parameter;

  /* ── ④ 설정 탭 공유 저장/조회 — 기존 리드게이트와 무관한 별도 기능 ── */
  if (p.action === 'load') {
    var json = PropertiesService.getScriptProperties().getProperty(SETTINGS_PROP_KEY) || '{}';
    var data;
    try { data = JSON.parse(json); } catch (err) { data = {}; }
    if (p.callback) return jsonpResp(p.callback, data);
    return jsonResp(data);
  }
  if (p.action === 'save') {
    if (String(p.pw) !== String(ADMIN_PW_HASH)) {
      var authErr = { success: false, error: 'auth' };
      if (p.callback) return jsonpResp(p.callback, authErr);
      return jsonResp(authErr);
    }
    if (!p.data) {
      var dataErr = { success: false, error: 'no_data' };
      if (p.callback) return jsonpResp(p.callback, dataErr);
      return jsonResp(dataErr);
    }
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(5000);
      PropertiesService.getScriptProperties().setProperty(SETTINGS_PROP_KEY, p.data);
    } finally {
      lock.releaseLock();
    }
    var ok = { success: true };
    if (p.callback) return jsonpResp(p.callback, ok);
    return jsonResp(ok);
  }

  /* ── 기존 리드게이트: action 파라미터 없는 요청은 전부 예전 그대로 시트에 기록 ── */
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '접수일시','회사명','담당자명','연락처/이메일','RE100 공급시기',
        '20년누적절감','총전력량','연간발전량','초과발전량',
        '설비용량','연평균발전시간','RE100이행률','PPA단가'
      ]);
    }
    sheet.appendRow([
      Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm:ss'),
      p.company  ||'',
      p.name     ||'',
      p.contact  ||'',
      p.re100    ||'',
      p.cum20    ||'',
      p.totalKwh ||'',
      p.totalGen ||'',
      p.excessKwh||'',
      p.capacity ||'',
      p.hours    ||'',
      p.re100rate||'',
      p.ppaPrice ||''
    ]);
    return ContentService.createTextOutput('ok');
  } catch(err) {
    return ContentService.createTextOutput('error:'+err.toString());
  }
}

function doPost(e){ return doGet(e); }

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpResp(callback, obj) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(obj) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
