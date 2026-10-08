// ═══════════════════════════════════════════════════════════════
// E-PPA 설정 공유 백엔드 — ④ 설정 탭에서 저장한 값을 모든 방문자에게 적용
//
// 지금까지는 저장 버튼을 눌러도 localStorage(이 브라우저에만 저장)라서
// 관리자 본인 외에는 적용되지 않았음. 이 스크립트를 배포하면 저장값을
// Apps Script의 PropertiesService(스크립트 전체가 공유하는 저장소, 별도
// 시트 불필요)에 보관하고, summary.html이 페이지 로드 시 여기서 읽어와
// 모든 방문자에게 동일하게 적용한다.
//
// [배포 순서]
// 1. script.google.com 에서 새 프로젝트 생성
// 2. 이 파일 내용 전체 붙여넣기
// 3. 배포 → 새 배포 → 유형: 웹 앱
//    실행 계정: 나(본인), 액세스 권한: 모든 사용자
// 4. 배포 후 나오는 웹 앱 URL을 복사
// 5. summary.html에서 SETTINGS_API_URL = '...' 부분에 그 URL을 붙여넣기
//
// [비밀번호]
// 쓰기(action=save)는 pw 파라미터가 ADMIN_PW_HASH와 일치해야 허용된다.
// 이 값은 summary.html의 ADMIN_PW_HASH와 반드시 같아야 하며, 비밀번호를
// 바꾸면(현재 8890) summary.html에서 새로 계산한 해시로 두 곳 다 갱신해야 함.
// 클라이언트 코드에 이미 공개돼 있는 값이라 강한 보안은 아니고, 아무나
// 실수로 덮어쓰는 걸 막는 최소한의 장치임.
// ═══════════════════════════════════════════════════════════════

var ADMIN_PW_HASH = 1723927; // '8890' 해시 — summary.html의 ADMIN_PW_HASH와 동일하게 유지
var PROP_KEY = 'eppa_admin_settings';

function doGet(e) {
  var p = e.parameter;

  if (p.action === 'load') {
    var json = PropertiesService.getScriptProperties().getProperty(PROP_KEY) || '{}';
    var data;
    try { data = JSON.parse(json); } catch (err) { data = {}; }
    if (p.callback) return jsonpResp(p.callback, data);
    return jsonResp(data);
  }

  if (p.action === 'save') {
    if (String(p.pw) !== String(ADMIN_PW_HASH)) {
      var err = { success: false, error: 'auth' };
      if (p.callback) return jsonpResp(p.callback, err);
      return jsonResp(err);
    }
    if (!p.data) {
      var err2 = { success: false, error: 'no_data' };
      if (p.callback) return jsonpResp(p.callback, err2);
      return jsonResp(err2);
    }
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(5000);
      PropertiesService.getScriptProperties().setProperty(PROP_KEY, p.data);
    } finally {
      lock.releaseLock();
    }
    var ok = { success: true };
    if (p.callback) return jsonpResp(p.callback, ok);
    return jsonResp(ok);
  }

  return ContentService.createTextOutput('E-PPA Settings API');
}

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
