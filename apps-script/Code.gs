/**
 * 공수 측정 프로그램 백엔드 (Google Apps Script)
 *
 * work-logs를 제외한 키는 스프레드시트의 key-value 저장소 시트를 사용한다.
 *   A열: key, B열: value(JSON 문자열), 1행: 헤더
 *
 * work-logs만 예외적으로 전용 시트 'work-logs-db'에 "한 행 = 기록 1건"으로 저장한다
 * (JSON 블롭 방식은 스프레드시트 셀 글자수 제한(약 5만자)에 근접해 폐기됨).
 * doGet의 액션 이름/파라미터/응답 형태는 기존과 동일하게 유지되므로 클라이언트(index.html,
 * dashboard.html)는 이 변경을 알 필요가 없다 — 내부 저장 방식만 바뀐 것.
 *
 * 지원 액션 (모두 GET — 웹앱 리다이렉트 시 POST 본문이 유실되는 문제 회피):
 *   ?action=get&key=...              → { key, value }  (없으면 value: null)
 *       key=work-logs이면 work-logs-db 시트를 읽어 배열 JSON 문자열로 반환한다.
 *   ?action=set&key=...&value=...    → { key, ok: true }
 *       key=work-logs이면 value(전체 배열 JSON)로 work-logs-db 시트를 통째로 재작성한다.
 *       ⚠ 로그가 많으면 URL 길이 제한에 걸릴 수 있으므로 정상 경로에서는 쓰지 않는다
 *       (일괄 수정은 반드시 update-log를 건별로 사용할 것).
 *   ?action=append&key=...&value=... → { key, ok: true, appended: true|false }
 *       value는 배열에 추가할 항목 1건의 JSON.
 *       같은 id가 이미 있으면 추가하지 않는다(멱등) — 버튼 연타/재시도에 안전.
 *       LockService로 직렬화하므로 여러 명이 동시에 저장해도 기록이 유실되지 않는다.
 *   ?action=start-timer&value=...     → { ok: true, started: true|false, timers: [...] }
 *       value는 타이머 1건의 JSON. 같은 id 또는 같은 작업(인원·차종·MID·세부·단계)의
 *       타이머가 이미 있으면 추가하지 않고 started:false를 돌려준다.
 *       읽기+중복검사+쓰기를 서버에서 한 번에 처리하므로 클라이언트 왕복이 1회로 줄어든다.
 *   ?action=remove-timer&id=...       → { ok: true, removed: true|false, timers: [...] }
 *       active-timers에서 해당 id를 제거한다. (타이머 종료 저장/취소용)
 *   ?action=update-log&value=...      → { ok: true, updated: true|false }
 *       로그 1건(id 포함)을 work-logs-db에서 해당 행만 찾아 교체한다.
 *   ?action=delete-log&id=...         → { ok: true, deleted: true|false }
 *       work-logs-db에서 해당 id 행만 삭제한다.
 *   ?action=migrate-worklogs[&force=1] → { ok: true, migrated, skipped, overwritten, totalOld }
 *       (1회성) 예전 key-value 시트의 work-logs JSON 블롭을 읽어 work-logs-db로 옮긴다.
 *       id가 이미 work-logs-db에 있으면 건너뛰므로 여러 번 실행해도 안전(멱등)하다.
 *       force=1이면 이미 있는 id도 예전 블롭 값으로 덮어쓴다 — work-logs-db 자체 결함을
 *       백업 원본으로 복구할 때 사용(마이그레이션 이후 새로 생긴 항목은 예전 블롭에 없으므로 건드리지 않음).
 *       예전 블롭 자체는 지우지 않고 백업으로 남겨둔다.
 *   ?action=register-release&value=... → { ok: true, release: {...}, releases: [...] }
 *       value는 { month, type, label?, vehicles, applicableMids, createdBy? } 형태의 JSON.
 *       같은 month+type 조합 내 등록 순번(seq)을 서버에서 계산해 이름을 자동 생성한다
 *       (예: "2026-07 정기펌웨어 2차 등록 (VCU 이슈 대응)"). LockService로 직렬화하므로
 *       PM/관리자가 동시에 등록해도 순번이 겹치지 않는다. meta-releases 배열에 append.
 *
 *   ?action=get-summary                → Summary.gs가 만든 '공수 집계' 탭 값 (읽기 전용, 주간 집계 탭용)
 *
 * 미러 시트 '공수로그':
 *   work-logs-db가 바뀔 때마다 로그를 "한 행 = 기록 1건" 표 형태로 자동 반영한다.
 *   Looker Studio 등 BI 도구 연결용이며 조회 전용 — 이 시트를 직접 수정해도
 *   앱 데이터(work-logs-db)에는 반영되지 않고, 다음 갱신 때 덮어써진다.
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
      if (key === 'work-logs') {
        out = { key: key, value: JSON.stringify(getAllLogEntries_()) };
      } else {
        out = { key: key, value: getValue_(key) };
      }
    } else if (action === 'set') {
      out = withLock_(function () {
        if (key === 'work-logs') return setAllLogEntries_(e.parameter.value);
        setValue_(key, e.parameter.value);
        return { key: key, ok: true };
      });
    } else if (action === 'append') {
      out = withLock_(function () {
        if (key === 'work-logs') return appendLogEntry_(JSON.parse(e.parameter.value));
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
        return updateLogEntryDb_(e.parameter.value);
      });
    } else if (action === 'delete-log') {
      out = withLock_(function () {
        return deleteLogEntryDb_(e.parameter.id);
      });
    } else if (action === 'migrate-worklogs') {
      out = withLock_(function () {
        return migrateWorklogsToDb_(e.parameter.force === '1' || e.parameter.force === 'true');
      });
    } else if (action === 'get-summary') {
      // '공수 집계' 탭 값 조회 (Summary.gs, 읽기 전용) — 주간 집계 탭에서 사용
      out = (typeof summaryJson_ === 'function') ? summaryJson_() : { ok: false, error: 'Summary.gs not installed' };
    } else if (action === 'register-release') {
      out = withLock_(function () {
        return registerRelease_(e.parameter.value);
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
var KNOWN_KEYS = ['key', 'work-logs', 'active-timers', 'meta-members', 'meta-vehicles', 'meta-admin-pin', 'meta-releases', 'meta-categories', 'meta-pm-pin'];
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
           (x.releaseId || null) === (t.releaseId || null);
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
  }
  return { key: key, ok: true, appended: !exists };
}

/* ---------------- work-logs 전용 DB 시트 (행 기반) ---------------- */
// JSON 블롭은 스프레드시트 셀 글자수 제한(약 5만자)에 근접해 폐기.
// 'work-logs-db' 시트에 로그 1건 = 1행으로 저장하며, 이 시트가 원본(source of truth)이다.
// '공수로그' 시트는 계속 사람이 보기 위한 자동 미러(조회 전용)로만 남는다.
var LOGDB_SHEET_NAME = 'work-logs-db';
var LOGDB_HEADER = [
  'id', 'member', 'releaseId', 'month', 'vehicle', 'mid', 'sub', 'group', 'stage',
  'minutes', 'date', 'note', 'source', 'measuredMinutes', 'startedAt', 'savedAt', 'checkStatus', 'round'
];

function logDbSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(LOGDB_SHEET_NAME);
  var isNew = !s;
  if (isNew) s = ss.insertSheet(LOGDB_SHEET_NAME);
  var headerRange = s.getRange(1, 1, 1, LOGDB_HEADER.length);
  var current = s.getLastColumn() > 0 ? s.getRange(1, 1, 1, Math.min(s.getLastColumn(), LOGDB_HEADER.length)).getValues()[0] : [];
  var matches = current.length === LOGDB_HEADER.length && current.every(function (v, i) { return v === LOGDB_HEADER[i]; });
  if (!matches) headerRange.setValues([LOGDB_HEADER]).setFontWeight('bold');
  if (isNew) ensureLogDbTextFormat_(s);
  return s;
}

var LOGDB_MONTH_COL = LOGDB_HEADER.indexOf('month') + 1;
var LOGDB_DATE_COL = LOGDB_HEADER.indexOf('date') + 1;

/**
 * 'month'(YYYY-MM), 'date'(YYYY-MM-DD) 컬럼은 그대로 두면 Sheets가 자동으로 날짜로
 * 인식해 Date 객체로 바꿔버린다(예: "2026-07-15" → 자정 KST가 UTC로 환산되어
 * "2026-07-14T15:00:00.000Z"로 저장/조회됨). 일반 텍스트 서식을 강제해 원본 문자열을 그대로 보존한다.
 * 헤더 아래로 넉넉히(2000행) 미리 서식을 걸어둬서 앞으로 append되는 행도 자동으로 적용받는다.
 * (아래 writeLogDbRow_이 쓰기 시점마다 한 번 더 강제하므로 이건 보조 안전장치)
 */
function ensureLogDbTextFormat_(sheet) {
  var rows = Math.max(sheet.getMaxRows() - 1, 2000);
  sheet.getRange(2, LOGDB_MONTH_COL, rows, 1).setNumberFormat('@');
  sheet.getRange(2, LOGDB_DATE_COL, rows, 1).setNumberFormat('@');
}

/**
 * work-logs-db 특정 행에 로그 1건을 쓴다. month/date 셀은 쓰기 "직전"에 매번 일반 텍스트
 * 서식을 강제한다 — 열 전체에 미리 걸어둔 서식만 믿으면(ensureLogDbTextFormat_) 드물게
 * 새로 추가되는 행에 서식이 안 먹는 경우가 있어(실제로 96건 중 4건에서 발생 확인),
 * 매 쓰기마다 직접 강제하는 쪽이 확실하다.
 */
function writeLogDbRow_(sheet, rowNum, entry) {
  sheet.getRange(rowNum, LOGDB_MONTH_COL).setNumberFormat('@');
  sheet.getRange(rowNum, LOGDB_DATE_COL).setNumberFormat('@');
  sheet.getRange(rowNum, 1, 1, LOGDB_HEADER.length).setValues([logEntryToDbRow_(entry)]);
}

/** Date 객체로 깨진 채 남아있는 과거 셀을 읽을 때도 안전하게 원래 문자열 형태로 되돌린다 */
function normalizeDateCell_(v, pattern) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), pattern);
  }
  return v;
}

/** id가 있는 행 번호(1-base)를 반환, 없으면 -1 */
function findLogDbRow_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function logEntryToDbRow_(e) {
  return [
    e.id || '', e.member || '', e.releaseId || '', e.month || '', e.vehicle || '',
    e.mid || '', e.sub || '', e.group || '', e.stage || '',
    (e.minutes != null) ? e.minutes : '', e.date || '', e.note || '', e.source || '',
    (e.measuredMinutes != null) ? e.measuredMinutes : '',
    (e.startedAt != null) ? e.startedAt : '',
    (e.savedAt != null) ? e.savedAt : '',
    e.checkStatus || '',
    (e.round != null) ? e.round : ''
  ];
}

function dbRowToLogEntry_(row) {
  var e = {
    id: row[0], member: row[1], releaseId: row[2] || null, month: normalizeDateCell_(row[3], 'yyyy-MM'), vehicle: row[4],
    mid: row[5], sub: row[6], group: row[7], stage: row[8] || null,
    minutes: row[9] === '' ? 0 : row[9],
    date: normalizeDateCell_(row[10], 'yyyy-MM-dd'), note: row[11] || '', source: row[12]
  };
  if (row[13] !== '') e.measuredMinutes = row[13];
  if (row[14] !== '') e.startedAt = row[14];
  if (row[15] !== '') e.savedAt = row[15];
  if (row[16] !== '') e.checkStatus = row[16];
  if (row[17] !== '') e.round = row[17];
  return e;
}

function getAllLogEntries_() {
  var sheet = logDbSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, LOGDB_HEADER.length).getValues();
  return rows.filter(function (r) { return r[0] !== ''; }).map(dbRowToLogEntry_);
}

/** 로그 1건 append (멱등: 같은 id가 이미 있으면 건너뜀) */
function appendLogEntry_(item) {
  var sheet = logDbSheet_();
  var exists = item && item.id && findLogDbRow_(sheet, item.id) !== -1;
  if (!exists) {
    writeLogDbRow_(sheet, sheet.getLastRow() + 1, item);
    mirrorAppend_(item);
  }
  return { key: 'work-logs', ok: true, appended: !exists };
}

/** 로그 1건을 id 기준으로 교체 (관리자 수정용). 전체 배열을 URL로 보내지 않아 로그가 많아져도 안전 */
function updateLogEntryDb_(entryJson) {
  var entry = JSON.parse(entryJson);
  if (!entry || !entry.id) return { ok: false, error: 'entry id required' };
  var sheet = logDbSheet_();
  var row = findLogDbRow_(sheet, entry.id);
  var found = row !== -1;
  if (found) {
    writeLogDbRow_(sheet, row, entry);
    rebuildMirrorFromDb_();
  }
  return { ok: true, updated: found };
}

/** 로그 1건을 id 기준으로 삭제 */
function deleteLogEntryDb_(id) {
  var sheet = logDbSheet_();
  var row = findLogDbRow_(sheet, id);
  var deleted = row !== -1;
  if (deleted) {
    sheet.deleteRow(row);
    rebuildMirrorFromDb_();
  }
  return { ok: true, deleted: deleted };
}

/** action=set&key=work-logs 폴백 경로용 — 전체 배열로 DB 시트를 통째로 재작성 */
function setAllLogEntries_(logsJson) {
  var logs = JSON.parse(logsJson);
  if (Object.prototype.toString.call(logs) !== '[object Array]') return { key: 'work-logs', ok: false, error: 'value is not an array' };
  var sheet = logDbSheet_();
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, LOGDB_HEADER.length).clearContent();
  if (logs.length) {
    sheet.getRange(2, LOGDB_MONTH_COL, logs.length, 1).setNumberFormat('@');
    sheet.getRange(2, LOGDB_DATE_COL, logs.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, logs.length, LOGDB_HEADER.length).setValues(logs.map(logEntryToDbRow_));
  }
  rebuildMirrorFromDb_();
  return { key: 'work-logs', ok: true };
}

/**
 * (1회성) 예전 key-value 시트에 JSON 블롭으로 저장돼 있던 work-logs를 읽어
 * work-logs-db 시트로 옮긴다. id가 이미 DB에 있으면 건너뛰므로 여러 번 실행해도 안전(멱등)하다.
 * force=true면 이미 있는 id도 예전 블롭 값으로 덮어쓴다 — DB 시트 자체의 결함(예: 텍스트 서식
 * 누락으로 날짜가 깨진 경우)을 백업 원본으로 복구할 때 사용. 예전 블롭은 지우지 않고 남겨둔다.
 */
function migrateWorklogsToDb_(force) {
  var raw = getValue_('work-logs');
  var oldLogs = raw ? JSON.parse(raw) : [];
  if (Object.prototype.toString.call(oldLogs) !== '[object Array]') {
    return { ok: false, error: 'old work-logs is not an array' };
  }
  var sheet = logDbSheet_();
  ensureLogDbTextFormat_(sheet);
  var migrated = 0, skipped = 0, overwritten = 0;
  oldLogs.forEach(function (e) {
    var row = findLogDbRow_(sheet, e.id);
    if (row === -1) {
      writeLogDbRow_(sheet, sheet.getLastRow() + 1, e);
      migrated++;
    } else if (force) {
      writeLogDbRow_(sheet, row, e);
      overwritten++;
    } else {
      skipped++;
    }
  });
  rebuildMirrorFromDb_();
  return { ok: true, migrated: migrated, skipped: skipped, overwritten: overwritten, totalOld: oldLogs.length };
}

/* ---------------- 배포 건(펌웨어 배포) 등록 ---------------- */

/**
 * 배포 건 신규 등록. 같은 month+type 조합의 기존 등록 건수를 세어 seq를 매기고,
 * "{month} {type} {seq}차 등록 (label)" 형태로 이름을 자동 생성한다.
 * LockService로 감싸져 있어 동시 등록 시에도 seq가 겹치지 않는다.
 */
function registerRelease_(payloadJson) {
  var payload = JSON.parse(payloadJson);
  if (!payload || !payload.month || !payload.type) {
    return { ok: false, error: 'month/type required' };
  }
  var releases = readArray_('meta-releases');
  if (releases === null) return { ok: false, error: 'meta-releases is not an array' };

  var seq = releases.filter(function (r) {
    return r && r.month === payload.month && r.type === payload.type;
  }).length + 1;

  var label = (payload.label || '').trim();
  var name = payload.month + ' ' + payload.type + ' ' + seq + '차 등록' + (label ? ' (' + label + ')' : '');

  var release = {
    id: 'rel-' + Utilities.getUuid(),
    name: name,
    month: payload.month,
    type: payload.type,
    seq: seq,
    label: label,
    vehicles: payload.vehicles || [],
    applicableMids: payload.applicableMids || [],
    createdAt: new Date().toISOString(),
    createdBy: payload.createdBy || ''
  };
  releases.push(release);
  setValue_('meta-releases', JSON.stringify(releases));
  return { ok: true, release: release, releases: releases };
}

/* ---------------- 미러 시트 (Looker Studio 연결용) ---------------- */
var MIRROR_SHEET_NAME = '공수로그';
var MIRROR_HEADER = [
  '날짜', '담당자', '월', '배포 건', '차종', '구분', '중분류', '소분류', '작업단계', '소요시간(분)', '방식', '비고', 'id',
  '측정값(분)', '차이(분)', '시작시각', '저장시각', '확인상태'
];
var MIRROR_DATE_COL = MIRROR_HEADER.indexOf('날짜') + 1;
var MIRROR_MONTH_COL = MIRROR_HEADER.indexOf('월') + 1;
var MIRROR_DIFF_COL = MIRROR_HEADER.indexOf('차이(분)') + 1;
var MIRROR_STARTED_COL = MIRROR_HEADER.indexOf('시작시각') + 1;
var MIRROR_SAVED_COL = MIRROR_HEADER.indexOf('저장시각') + 1;
var MIRROR_DATETIME_FORMAT = 'yyyy-MM-dd HH:mm:ss';

/** { releaseId: name } 조회용 맵. 미러 전체 재구성 시 행마다 다시 읽지 않도록 한 번만 만든다 */
function releasesMap_() {
  var raw = getValue_('meta-releases');
  var map = {};
  if (raw) {
    try {
      JSON.parse(raw).forEach(function (r) { if (r && r.id) map[r.id] = r.name; });
    } catch (err) {}
  }
  return map;
}

function mirrorSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(MIRROR_SHEET_NAME);
  var isNew = !s;
  if (isNew) s = ss.insertSheet(MIRROR_SHEET_NAME);
  // 헤더가 최신 컬럼 구성과 다르면(신규 생성 포함) 항상 맞춰 쓴다 — 기존 데이터 행은 그대로 유지
  var headerRange = s.getRange(1, 1, 1, MIRROR_HEADER.length);
  var current = s.getLastColumn() > 0 ? s.getRange(1, 1, 1, Math.min(s.getLastColumn(), MIRROR_HEADER.length)).getValues()[0] : [];
  var matches = current.length === MIRROR_HEADER.length && current.every(function (v, i) { return v === MIRROR_HEADER[i]; });
  if (!matches) {
    headerRange.setValues([MIRROR_HEADER]).setFontWeight('bold');
    // 컬럼 구성이 바뀌면(예: '차이(분)'을 중간에 끼워넣어 뒤 컬럼들이 밀림) 예전 자리에 남아있던
    // 서식(예: 예전엔 '시작시각' 자리였는데 지금은 숫자 컬럼인 경우 날짜 서식이 남아있음)이
    // 새 데이터를 엉뚱하게 보여줄 수 있어, 헤더가 바뀔 때 데이터 영역 서식을 전부 리셋한다
    resetMirrorDataFormat_(s);
  }
  return s;
}

/** 데이터 영역 서식을 일반으로 리셋한 뒤 날짜/월은 텍스트로, 시작·저장시각은 날짜+시간으로 강제한다 */
function resetMirrorDataFormat_(sheet) {
  var rows = Math.max(sheet.getMaxRows() - 1, 2000);
  sheet.getRange(2, 1, rows, MIRROR_HEADER.length).setNumberFormat('General');
  sheet.getRange(2, MIRROR_DATE_COL, rows, 1).setNumberFormat('@');
  sheet.getRange(2, MIRROR_MONTH_COL, rows, 1).setNumberFormat('@');
  sheet.getRange(2, MIRROR_STARTED_COL, rows, 1).setNumberFormat(MIRROR_DATETIME_FORMAT);
  sheet.getRange(2, MIRROR_SAVED_COL, rows, 1).setNumberFormat(MIRROR_DATETIME_FORMAT);
}

/** work-logs-db와 마찬가지로 미러 시트의 날짜/월/시각 셀도 쓰기 직전마다 서식을 강제한다 */
function writeMirrorRow_(sheet, rowNum, row) {
  sheet.getRange(rowNum, MIRROR_DATE_COL).setNumberFormat('@');
  sheet.getRange(rowNum, MIRROR_MONTH_COL).setNumberFormat('@');
  sheet.getRange(rowNum, MIRROR_DIFF_COL).setNumberFormat('General');
  sheet.getRange(rowNum, MIRROR_STARTED_COL).setNumberFormat(MIRROR_DATETIME_FORMAT);
  sheet.getRange(rowNum, MIRROR_SAVED_COL).setNumberFormat(MIRROR_DATETIME_FORMAT);
  sheet.getRange(rowNum, 1, 1, MIRROR_HEADER.length).setValues([row]);
}

function groupLabel_(g) {
  if (g === 'verify') return '검증수행';
  if (g === 'pm') return 'PM업무';
  return '환경설정';
}

/** 신뢰도 표시: checkStatus가 있으면 그대로, 없으면(과거 데이터) source로 유추 */
function checkStatusLabel_(e) {
  if (e.checkStatus) return e.checkStatus;
  return e.source === 'manual' ? '수동입력' : '확인불가';
}

/**
 * epoch ms 숫자를 Date로 바꾸되, 값이 없거나 유효하지 않으면 빈 문자열을 돌려준다.
 * new Date(값)이 Invalid Date가 된 채로 시트에 써지면 Sheets가 날짜 일련번호 0(=1899-12-30)으로
 * 표시해버리는 함정이 있어, 여기서 미리 걸러 '이상한 옛날 날짜'가 절대 나오지 않게 한다.
 */
function safeDate_(ms) {
  if (typeof ms !== 'number' || !isFinite(ms) || ms <= 0) return '';
  var d = new Date(ms);
  return isNaN(d.getTime()) ? '' : d;
}

function logToRow_(e, relMap) {
  var relName = e.releaseId ? ((relMap && relMap[e.releaseId]) || '(삭제된 배포 건)') : '미지정';
  // 실제 측정값과 저장값이 얼마나 차이 나는지(수정 폭) 바로 보이도록 계산해둔다.
  // 측정값 자체가 없는(과거) 기록은 비교 불가하므로 빈 칸으로 둔다.
  var diff = (e.measuredMinutes != null) ? Math.round((e.minutes - e.measuredMinutes) * 100) / 100 : '';
  return [
    e.date || '', e.member || '', e.month || '', relName, e.vehicle || '',
    groupLabel_(e.group), e.mid || '', e.sub || '', e.stage || '',
    e.minutes || 0, e.source === 'timer' ? '타이머' : '직접입력', e.note || '', e.id || '',
    (e.measuredMinutes != null) ? e.measuredMinutes : '',
    diff,
    safeDate_(e.startedAt),
    safeDate_(e.savedAt),
    checkStatusLabel_(e)
  ];
}

function mirrorAppend_(item) {
  try {
    var s = mirrorSheet_();
    writeMirrorRow_(s, s.getLastRow() + 1, logToRow_(item, releasesMap_()));
  } catch (err) {
    // 미러 실패가 본 저장을 막으면 안 되므로 삼킨다 (원본 work-logs-db가 항상 기준)
  }
}

/** work-logs-db 전체 내용으로 미러 시트를 다시 채운다 */
function rebuildMirror_(logsJson) {
  try {
    var logs = JSON.parse(logsJson);
    if (Object.prototype.toString.call(logs) !== '[object Array]') return;
    var s = mirrorSheet_();
    var last = s.getLastRow();
    if (last > 1) s.getRange(2, 1, last - 1, MIRROR_HEADER.length).clearContent();
    if (logs.length) {
      var relMap = releasesMap_();
      var rows = logs.map(function (e) { return logToRow_(e, relMap); });
      // 전체 재구성 때마다 서식을 일반으로 리셋 — 예전 컬럼 배치 때 남은 서식이
      // (헤더는 이미 최신이라 mirrorSheet_의 리셋을 안 타는 경우에도) 새 데이터를
      // 잘못 표시하지 않도록 매번 확실히 한다
      s.getRange(2, 1, rows.length, MIRROR_HEADER.length).setNumberFormat('General');
      s.getRange(2, MIRROR_DATE_COL, rows.length, 1).setNumberFormat('@');
      s.getRange(2, MIRROR_MONTH_COL, rows.length, 1).setNumberFormat('@');
      s.getRange(2, MIRROR_STARTED_COL, rows.length, 1).setNumberFormat(MIRROR_DATETIME_FORMAT);
      s.getRange(2, MIRROR_SAVED_COL, rows.length, 1).setNumberFormat(MIRROR_DATETIME_FORMAT);
      s.getRange(2, 1, rows.length, MIRROR_HEADER.length).setValues(rows);
    }
  } catch (err) {
    // 미러 실패가 본 저장을 막으면 안 되므로 삼킨다
  }
}

/** work-logs-db 시트를 원본으로 삼아 미러를 재구성 (update-log/delete-log/migrate 등에서 사용) */
function rebuildMirrorFromDb_() {
  rebuildMirror_(JSON.stringify(getAllLogEntries_()));
}
