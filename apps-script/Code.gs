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
 *   ?action=start-timer&value=...     → { ok: true, started: true|false, timers: [...] }
 *       value는 타이머 1건의 JSON. 같은 id 또는 같은 작업(인원·차종·MID·세부·단계)의
 *       타이머가 이미 있으면 추가하지 않고 started:false를 돌려준다.
 *       읽기+중복검사+쓰기를 서버에서 한 번에 처리하므로 클라이언트 왕복이 1회로 줄어든다.
 *   ?action=remove-timer&id=...       → { ok: true, removed: true|false, timers: [...] }
 *       active-timers에서 해당 id를 제거한다. (타이머 종료 저장/취소용)
 *
 * 미러 시트 '공수로그':
 *   work-logs가 바뀔 때마다 로그를 "한 행 = 기록 1건" 표 형태로 자동 반영한다.
 *   Looker Studio 등 BI 도구 연결용이며 조회 전용 — 이 시트를 직접 수정해도
 *   앱 데이터(work-logs JSON)에는 반영되지 않고, 다음 갱신 때 덮어써진다.
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
        if (key === 'work-logs') rebuildMirror_(e.parameter.value);
        return { key: key, ok: true };
      });
    } else if (action === 'append') {
      out = withLock_(function () {
        return appendItem_(key, e.parameter.value);
      });
    } else if (action === 'start-timer') {
      out = withLock_(function () {
        return startTimer_(e.parameter.value);
      });
    } else if (action === 'remove-timer') {
      out = withLock_(function () {
        return removeTimer_(e.parameter.id);
      });
    } else if (action === 'update-log') {
      out = withLock_(function () {
        return updateLog_(e.parameter.value);
      });
    } else if (action === 'delete-log') {
      out = withLock_(function () {
        return deleteLog_(e.parameter.id);
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

/** key의 값을 배열로 파싱해서 반환. 값이 없으면 빈 배열, 배열이 아니면 null */
function readArray_(key) {
  var raw = getValue_(key);
  var arr = raw ? JSON.parse(raw) : [];
  return Object.prototype.toString.call(arr) === '[object Array]' ? arr : null;
}

function startTimer_(timerJson) {
  var t = JSON.parse(timerJson);
  var arr = readArray_('active-timers');
  if (arr === null) return { ok: false, error: 'active-timers is not an array' };
  var dup = arr.some(function (x) {
    if (!x) return false;
    if (x.id === t.id) return true;
    return x.member === t.member && x.vehicle === t.vehicle && x.mid === t.mid &&
           x.sub === t.sub && (x.stage || null) === (t.stage || null) &&
           (x.round || 1) === (t.round || 1);
  });
  if (!dup) {
    arr.push(t);
    setValue_('active-timers', JSON.stringify(arr));
  }
  return { ok: true, started: !dup, timers: arr };
}

function removeTimer_(id) {
  var arr = readArray_('active-timers');
  if (arr === null) return { ok: false, error: 'active-timers is not an array' };
  var next = arr.filter(function (x) { return x && x.id !== id; });
  var removed = next.length !== arr.length;
  if (removed) setValue_('active-timers', JSON.stringify(next));
  return { ok: true, removed: removed, timers: next };
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
    if (key === 'work-logs') mirrorAppend_(item);
  }
  return { key: key, ok: true, appended: !exists };
}

/** 로그 1건을 id 기준으로 교체 (관리자 수정용). 전체 배열을 URL로 보내지 않아 로그가 많아져도 안전 */
function updateLog_(entryJson) {
  var entry = JSON.parse(entryJson);
  if (!entry || !entry.id) return { ok: false, error: 'entry id required' };
  var arr = readArray_('work-logs');
  if (arr === null) return { ok: false, error: 'work-logs is not an array' };
  var found = false;
  for (var i = 0; i < arr.length; i++) {
    if (arr[i] && arr[i].id === entry.id) { arr[i] = entry; found = true; break; }
  }
  if (found) {
    var json = JSON.stringify(arr);
    setValue_('work-logs', json);
    rebuildMirror_(json);
  }
  return { ok: true, updated: found };
}

/** 로그 1건을 id 기준으로 삭제 */
function deleteLog_(id) {
  var arr = readArray_('work-logs');
  if (arr === null) return { ok: false, error: 'work-logs is not an array' };
  var next = arr.filter(function (x) { return x && x.id !== id; });
  var deleted = next.length !== arr.length;
  if (deleted) {
    var json = JSON.stringify(next);
    setValue_('work-logs', json);
    rebuildMirror_(json);
  }
  return { ok: true, deleted: deleted };
}

/* ---------------- 미러 시트 (Looker Studio 연결용) ---------------- */
var MIRROR_SHEET_NAME = '공수로그';
var MIRROR_HEADER = ['날짜', '담당자', '월', '차수', '차종', '구분', '중분류', '소분류', '작업단계', '소요시간(분)', '방식', '비고', 'id'];

function mirrorSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(MIRROR_SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(MIRROR_SHEET_NAME);
    s.getRange(1, 1, 1, MIRROR_HEADER.length).setValues([MIRROR_HEADER]).setFontWeight('bold');
  }
  return s;
}

function groupLabel_(g) {
  if (g === 'verify') return '검증수행';
  if (g === 'pm') return 'PM업무';
  return '환경설정';
}

function logToRow_(e) {
  return [
    e.date || '', e.member || '', e.month || '', (e.round || 1) + '차', e.vehicle || '',
    groupLabel_(e.group), e.mid || '', e.sub || '', e.stage || '',
    e.minutes || 0, e.source === 'timer' ? '타이머' : '직접입력', e.note || '', e.id || ''
  ];
}

function mirrorAppend_(item) {
  try {
    mirrorSheet_().appendRow(logToRow_(item));
  } catch (err) {
    // 미러 실패가 본 저장을 막으면 안 되므로 삼킨다 (원본 work-logs가 항상 기준)
  }
}

/** work-logs 전체 교체(set) 시 미러 시트를 원본과 일치하게 재작성 */
function rebuildMirror_(logsJson) {
  try {
    var logs = JSON.parse(logsJson);
    if (Object.prototype.toString.call(logs) !== '[object Array]') return;
    var s = mirrorSheet_();
    var last = s.getLastRow();
    if (last > 1) s.getRange(2, 1, last - 1, MIRROR_HEADER.length).clearContent();
    if (logs.length) {
      var rows = logs.map(logToRow_);
      s.getRange(2, 1, rows.length, MIRROR_HEADER.length).setValues(rows);
    }
  } catch (err) {
    // 미러 실패가 본 저장을 막으면 안 되므로 삼킨다
  }
}
