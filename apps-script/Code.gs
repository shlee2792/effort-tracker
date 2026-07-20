/**
 * 공수 측정 프로그램 백엔드 (Google Apps Script)
 *
 * 스프레드시트 첫 번째 시트를 key-value 저장소로 사용한다.
 *   A열: key, B열: value(JSON 문자열), 1행: 헤더
 *
 * 지원 액션 (모두 GET — 웹앱 리다이렉트 시 POST 본문이 유실되는 문제 회피):
 *   ?action=get&key=...              → { key, value }  (없으면 value: null)
 *   ?action=set&key=...&value=...    → { key, ok: true }
 *   ?action=append&key=...&value=... → { key, ok: true, appended: true|false }
 *       value는 배열에 추가할 항목 1건의 JSON.
 *       같은 id가 이미 배열에 있으면 추가하지 않는다(멱등) — 버튼 연타/재시도에 안전.
 *       LockService로 직렬화하므로 여러 명이 동시에 저장해도 기록이 유실되지 않는다.
 *
 * 배포: 스크립트 편집기에서 코드 교체 후 "배포 > 배포 관리 > 연필 아이콘 > 새 버전"으로
 * 재배포해야 기존 URL이 유지된다. (새 배포를 만들면 URL이 바뀌므로 주의)
 */

function doGet(e) {
  var action = e.parameter.action;
  var key = e.parameter.key;
  var out;

  try {
    if (action === 'get') {
      out = { key: key, value: getValue_(key) };
    } else if (action === 'set') {
      out = withLock_(function () {
        setValue_(key, e.parameter.value);
        return { key: key, ok: true };
      });
    } else if (action === 'append') {
      out = withLock_(function () {
        return appendItem_(key, e.parameter.value);
      });
    } else {
      out = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    out = { error: String(err) };
  }

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// 실제 key-value 데이터가 들어 있는 시트를 식별하는 단서:
// A열에 'key' 헤더 또는 앱이 쓰는 키가 존재하는 시트
var KNOWN_KEYS = ['key', 'work-logs', 'active-timers', 'meta-members', 'meta-vehicles', 'meta-admin-pin'];
var SHEET_CACHE = null;

function sheet_() {
  if (SHEET_CACHE) return SHEET_CACHE;
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var best = sheets[0];
  var bestScore = -1;
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var last = s.getLastRow();
    if (last < 1) continue;
    var colA = s.getRange(1, 1, last, 1).getValues();
    var score = 0;
    for (var j = 0; j < colA.length; j++) {
      if (KNOWN_KEYS.indexOf(String(colA[j][0])) !== -1) score++;
    }
    if (score > bestScore) { bestScore = score; best = s; }
  }
  SHEET_CACHE = best;
  return SHEET_CACHE;
}

/** key가 있는 행 번호(1-base)를 반환, 없으면 -1. 헤더 유무와 무관하게 전체를 훑는다 */
function findRow_(sheet, key) {
  var last = sheet.getLastRow();
  if (last < 1) return -1;
  var keys = sheet.getRange(1, 1, last, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === String(key)) return i + 1;
  }
  return -1;
}

function getValue_(key) {
  var sheet = sheet_();
  var row = findRow_(sheet, key);
  if (row === -1) return null;
  var v = sheet.getRange(row, 2).getValue();
  return v === '' ? null : String(v);
}

function setValue_(key, value) {
  var sheet = sheet_();
  var row = findRow_(sheet, key);
  if (row === -1) {
    sheet.appendRow([key, value]);
  } else {
    sheet.getRange(row, 2).setValue(value);
  }
}

function appendItem_(key, itemJson) {
  var item = JSON.parse(itemJson);
  var raw = getValue_(key);
  var arr = raw ? JSON.parse(raw) : [];
  if (Object.prototype.toString.call(arr) !== '[object Array]') {
    return { key: key, ok: false, error: 'value is not an array' };
  }
  var exists = item && item.id && arr.some(function (x) { return x && x.id === item.id; });
  if (!exists) {
    arr.push(item);
    setValue_(key, JSON.stringify(arr));
  }
  return { key: key, ok: true, appended: !exists };
}
