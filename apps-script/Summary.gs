/**
 * '공수 집계' 탭 생성기 (결과 전용)
 *
 * 공수관리 사이트가 자동으로 채우는 '공수로그' 탭을 읽어서, 담당자 × 구분 × 월 × 주차별
 * 공수(시간)를 '공수 집계' 탭에 수식으로 보여준다. 기존 탭(공수로그, work-logs-db 등)은
 * 읽기만 하고 절대 수정하지 않는다.
 *
 * - 정기/공식: 공수로그의 '배포 건' 이름(정기펌웨어/공식펌웨어)으로 구분해 자동 합산 (실시간 반영)
 * - 실무시간외 검증: 정기/공식 중 확인상태 '수정됨' + 비고에 야간/야건/오버나잇/밤샘/점심시간이 있는 기록
 *   (정기/공식 행에서는 이 기록을 빼므로 두 행을 더하면 정기/공식 전체가 된다)
 * - 변경점, POC, 고객대응, 기타: 노란 칸에 직접 입력 (다시 생성해도 값 유지)
 * - 주차: 각 월의 Windows 달력 주(일~토). 예) 10월 1W 10/1~3, 2W 10/4~10 … 5W 10/25~31 / 11월 1W 11/1~7 … 5W 11/29~30
 * - 월 열은 startMonth부터 "실행한 날의 월"까지 만든다. 새 달이 시작되면 buildSummarySheet를
 *   다시 실행하거나, installSummaryMonthlyTrigger를 한 번 실행해 매월 1일 자동 갱신되게 한다.
 *
 * 사용법: Apps Script 편집기에서 이 파일을 추가 → 함수 'buildSummarySheet' 선택 → 실행.
 *        (웹앱 배포와 무관하므로 재배포할 필요 없음)
 */

var SUM_CFG = {
  sheetName: '공수 집계',
  sourceSheet: '공수로그',
  startMonth: '2026-10',
  members: ['문동권', '장서진', '황승현', '고유빈', '윤현우'],
  cats: ['정기/공식', '실무시간외 검증', '변경점', 'POC', '고객대응', '기타'],
  // 자동 집계 항목: 공수로그 '배포 건'(D열) 이름에 대한 정규식 + 실무시간외 여부
  //   offHours true  → 확인상태(R열)가 '수정됨'이고 비고(L열)에 offHoursPattern이 있는 기록만
  //   offHours false → 그 외 기록만 (두 행을 더하면 정기/공식 전체)
  autoPattern: {
    '정기/공식': { release: '정기펌웨어|공식펌웨어', offHours: false },
    '실무시간외 검증': { release: '정기펌웨어|공식펌웨어', offHours: true }
  },
  offHoursPattern: '야간|야건|오버나잇|밤샘|점심시간'
};

// 기존 엑셀(공수 측정 현황.xlsx)에 직접 입력돼 있던 값 — 탭을 처음 만들 때 한 번만 채운다
// [담당자, 구분, 월, 주차(1부터), 시간]
var SUM_SEED = [
  ['장서진', '고객대응', '2026-10', 4, 23.11],
  ['장서진', '고객대응', '2026-10', 5, 12.96],
  ['장서진', '기타', '2026-10', 4, 8],
  ['장서진', '기타', '2026-10', 5, 5],
  ['장서진', '고객대응', '2026-11', 1, 29.59]
];

var SUM_FIRST_ROW = 5;   // 데이터 시작 행 (1~4행: 제목/안내/월/주차)
var SUM_FIRST_COL = 4;   // D열부터 월 블록

function buildSummarySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getDocumentProperties();
  var oldLayout = JSON.parse(props.getProperty('SUMMARY_LAYOUT') || 'null');
  var sh = ss.getSheetByName(SUM_CFG.sheetName);

  // 1) 직접 입력 값 보존 (처음 생성이면 SEED 사용)
  var manual = {};
  if (sh && oldLayout) {
    manual = summaryReadManual_(sh, oldLayout);
  } else if (!oldLayout) {
    SUM_SEED.forEach(function (s) { manual[summaryKey_(s[0], s[1], s[2], s[3])] = s[4]; });
  }

  // 2) 탭 준비 (있으면 초기화)
  if (!sh) {
    sh = ss.insertSheet(SUM_CFG.sheetName, 0);
  } else {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    sh.clear();
  }

  var months = summaryMonths_();
  var L = summaryLayout_(months, SUM_CFG.members, SUM_CFG.cats);
  var nRows = L.grandRow;
  var nCols = L.lastCol;
  if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
  if (sh.getMaxRows() < nRows + 3) sh.insertRowsAfter(sh.getMaxRows(), nRows + 3 - sh.getMaxRows());

  // 3) 값/수식 그리드 작성
  var grid = [];
  for (var r = 0; r < nRows; r++) { var row = []; for (var c = 0; c < nCols; c++) row.push(''); grid.push(row); }
  function put(r, c, v) { grid[r - 1][c - 1] = v; }

  var tz = ss.getSpreadsheetTimeZone();
  put(1, 2, '공수 집계 (결과 전용 · 자동 계산)');
  put(2, 2, '정기/공식·실무시간외 검증은 \'' + SUM_CFG.sourceSheet + '\' 탭에서 자동 합산(시간). 실무시간외 = 수정됨 + 비고에 야간/오버나잇/밤샘/점심시간. ' +
            '노란 칸(변경점·POC·고객대응·기타)은 직접 입력. ' +
            '주차는 각 월의 Windows 달력 주(일~토). 마지막 생성: ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'));
  put(3, 2, '담당자'); put(3, 3, '구분');

  var src = "'" + SUM_CFG.sourceSheet + "'!";
  L.months.forEach(function (m) {
    put(3, m.col, m.label);
    m.weeks.forEach(function (w, wi) { put(4, m.col + wi, (wi + 1) + 'W\n' + w.label); });
    put(3, m.totalCol, '합산');
  });

  SUM_CFG.members.forEach(function (name, mi) {
    var top = L.memberTop[mi];
    put(top, 2, name);
    SUM_CFG.cats.forEach(function (cat, ci) {
      var r = top + ci;
      put(r, 3, cat);
      var pat = SUM_CFG.autoPattern[cat];
      L.months.forEach(function (m) {
        m.weeks.forEach(function (w, wi) {
          var c = m.col + wi;
          if (pat) {
            var off = '(' + src + '$R$2:$R="수정됨")*REGEXMATCH(' + src + '$L$2:$L&"","' + SUM_CFG.offHoursPattern + '")';
            put(r, c, '=ROUND(ARRAYFORMULA(SUMPRODUCT((' + src + '$B$2:$B=$B$' + top + ')*(' + src + '$A$2:$A>="' + w.from + '")*(' +
                      src + '$A$2:$A<="' + w.to + '")*REGEXMATCH(' + src + '$D$2:$D&"","' + pat.release + '")*' +
                      (pat.offHours ? off : '(1-' + off + ')') + '*' + src + '$J$2:$J))/60,2)');
          } else {
            var v = manual[summaryKey_(name, cat, m.ym, wi + 1)];
            put(r, c, (v === undefined || v === null) ? '' : v);
          }
        });
        put(r, m.totalCol, '=SUM(' + summaryA1_(r, m.col) + ':' + summaryA1_(r, m.col + m.weeks.length - 1) + ')');
      });
    });
    var tr = top + SUM_CFG.cats.length;
    put(tr, 3, '합산');
    for (var c = SUM_FIRST_COL; c <= nCols; c++) {
      put(tr, c, '=SUM(' + summaryA1_(top, c) + ':' + summaryA1_(tr - 1, c) + ')');
    }
  });

  put(L.grandRow, 3, '총 합산');
  for (var c2 = SUM_FIRST_COL; c2 <= nCols; c2++) {
    put(L.grandRow, c2, '=SUM(' + L.memberTop.map(function (t) {
      return summaryA1_(t + SUM_CFG.cats.length, c2);
    }).join(',') + ')');
  }

  sh.getRange(1, 1, nRows, nCols).setValues(grid);

  // 4) 서식
  summaryFormat_(sh, L, nRows, nCols);

  props.setProperty('SUMMARY_LAYOUT', JSON.stringify({ months: months, members: SUM_CFG.members, cats: SUM_CFG.cats }));
}

/** 매월 1일 06시에 탭을 다시 생성해 새 달 열을 자동으로 추가 (한 번만 실행) */
function installSummaryMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'buildSummarySheet') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buildSummarySheet').timeBased().onMonthDay(1).atHour(6).create();
}

/* ---------------- 내부 함수 ---------------- */

function summaryKey_(member, cat, ym, week) { return [member, cat, ym, week].join('|'); }

function summaryA1_(r, c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; }
  return s + r;
}

/** startMonth부터 이번 달까지 'YYYY-MM' 목록 */
function summaryMonths_() {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  var out = [], y = +SUM_CFG.startMonth.slice(0, 4), m = +SUM_CFG.startMonth.slice(5, 7);
  while (true) {
    var ym = y + '-' + (m < 10 ? '0' : '') + m;
    if (ym > now) break;
    out.push(ym);
    if (++m > 12) { m = 1; y++; }
  }
  return out;
}

/** 해당 월의 달력 주(일~토) 목록: [{from:'YYYY-MM-DD', to:'YYYY-MM-DD', label:'10/1~3'}] */
function summaryWeeks_(ym) {
  var y = +ym.slice(0, 4), m = +ym.slice(5, 7);
  var last = new Date(y, m, 0).getDate();
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  var weeks = [], start = 1;
  for (var d = 1; d <= last; d++) {
    var wd = new Date(y, m - 1, d).getDay(); // 0 = 일요일, 6 = 토요일 (Windows 달력처럼 일요일 시작)
    if (wd === 6 || d === last) {
      weeks.push({
        from: ym + '-' + pad(start), to: ym + '-' + pad(d),
        label: m + '/' + start + (d > start ? '~' + d : '')
      });
      start = d + 1;
    }
  }
  return weeks;
}

function summaryLayout_(months, members, cats) {
  var col = SUM_FIRST_COL, ms = [];
  months.forEach(function (ym) {
    var weeks = summaryWeeks_(ym);
    ms.push({ ym: ym, label: (+ym.slice(2, 4)) + '년 ' + (+ym.slice(5, 7)) + '월', col: col, weeks: weeks, totalCol: col + weeks.length });
    col += weeks.length + 1;
  });
  var tops = members.map(function (_, i) { return SUM_FIRST_ROW + i * (cats.length + 1); });
  return { months: ms, memberTop: tops, grandRow: SUM_FIRST_ROW + members.length * (cats.length + 1), lastCol: col - 1 };
}

/** 이전에 생성된 탭에서 직접 입력 칸 값을 읽어온다 */
function summaryReadManual_(sh, old) {
  var L = summaryLayout_(old.months, old.members, old.cats);
  var lastRow = Math.min(sh.getLastRow(), L.grandRow), lastCol = Math.min(sh.getLastColumn(), L.lastCol);
  var out = {};
  if (lastRow < SUM_FIRST_ROW || lastCol < SUM_FIRST_COL) return out;
  var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  old.members.forEach(function (name, mi) {
    old.cats.forEach(function (cat, ci) {
      if (SUM_CFG.autoPattern[cat]) return;
      var r = L.memberTop[mi] + ci;
      if (r > lastRow) return;
      L.months.forEach(function (m) {
        m.weeks.forEach(function (w, wi) {
          var c = m.col + wi;
          if (c > lastCol) return;
          var v = vals[r - 1][c - 1];
          if (v !== '' && v !== null) out[summaryKey_(name, cat, m.ym, wi + 1)] = v;
        });
      });
    });
  });
  return out;
}

function summaryFormat_(sh, L, nRows, nCols) {
  var FONT = '맑은 고딕';
  var all = sh.getRange(3, 2, nRows - 2, nCols - 1);
  sh.getRange(1, 1, nRows, nCols).setFontFamily(FONT).setFontSize(10).setVerticalAlignment('middle');
  sh.getRange(1, 2).setFontSize(13).setFontWeight('bold');
  sh.getRange(2, 2).setFontSize(9).setFontColor('#666666');
  all.setHorizontalAlignment('center').setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(SUM_FIRST_ROW, SUM_FIRST_COL, nRows - SUM_FIRST_ROW + 1, nCols - SUM_FIRST_COL + 1).setNumberFormat('0.00;-0.00;"-"');

  // 머리글
  sh.getRange(3, 2, 2, nCols - 1).setBackground('#d9e1f2').setFontWeight('bold').setWrap(true);
  sh.getRange(3, 2, 2, 1).merge(); sh.getRange(3, 3, 2, 1).merge();
  L.months.forEach(function (m) {
    sh.getRange(3, m.col, 1, m.weeks.length).merge();
    sh.getRange(3, m.totalCol, 2, 1).merge();
    sh.getRange(3, m.col, nRows - 2, m.weeks.length + 1)
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.getRange(SUM_FIRST_ROW, m.totalCol, nRows - SUM_FIRST_ROW + 1, 1).setBackground('#f2f2f2').setFontWeight('bold');
  });

  // 인원 블록
  var n = SUM_CFG.cats.length;
  L.memberTop.forEach(function (top) {
    sh.getRange(top, 2, n, 1).merge().setFontWeight('bold');
    SUM_CFG.cats.forEach(function (cat, ci) {
      if (!SUM_CFG.autoPattern[cat]) {
        L.months.forEach(function (m) { sh.getRange(top + ci, m.col, 1, m.weeks.length).setBackground('#fff2cc'); });
      }
    });
    sh.getRange(top + n, 2, 1, nCols - 1).setBackground('#e7e6e6').setFontWeight('bold');
    sh.getRange(top, 2, n + 1, nCols - 1)
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sh.getRange(L.grandRow, 2, 1, nCols - 1).setBackground('#bdd7ee').setFontWeight('bold')
    .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 크기/고정/탭 표시
  sh.setColumnWidth(1, 20); sh.setColumnWidth(2, 70); sh.setColumnWidth(3, 80);
  sh.setColumnWidths(SUM_FIRST_COL, nCols - SUM_FIRST_COL + 1, 62);
  sh.setRowHeight(4, 36);
  sh.setFrozenRows(4); sh.setFrozenColumns(3);
  sh.setTabColor('#1f6feb');

  // 실수로 수식을 고치지 않도록 경고(직접 입력 칸은 경고 없이 편집 가능)
  var prot = sh.protect().setDescription('공수 집계: 자동 계산 영역').setWarningOnly(true);
  var free = [];
  L.memberTop.forEach(function (top) {
    SUM_CFG.cats.forEach(function (cat, ci) {
      if (!SUM_CFG.autoPattern[cat]) L.months.forEach(function (m) { free.push(sh.getRange(top + ci, m.col, 1, m.weeks.length)); });
    });
  });
  prot.setUnprotectedRanges(free);
}

/* ---------------- 사이트(주간 집계 탭)용 조회 ---------------- */

/**
 * '공수 집계' 탭의 현재 값을 JSON으로 돌려준다 (Code.gs의 ?action=get-summary에서 호출, 읽기 전용).
 * 시간 단위(h) 숫자. 노란 칸(직접 입력) 값도 그대로 포함된다.
 * { ok, cats, autoCats, months:[{ym,label,weeks:['10/1~3',…]}],
 *   members:[{name, rows:[{ym:{weeks:[…],total}}…(cats 순서)], sum:{ym:{…}}}], grand:{ym:{…}} }
 */
function summaryJson_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SUM_CFG.sheetName);
  var layout = JSON.parse(PropertiesService.getDocumentProperties().getProperty('SUMMARY_LAYOUT') || 'null');
  if (!sh || !layout) return { ok: false, error: 'summary sheet not built' };

  var L = summaryLayout_(layout.months, layout.members, layout.cats);
  var vals = sh.getRange(1, 1, L.grandRow, L.lastCol).getValues();
  var num = function (v) { var n = Number(v); return (v === '' || v === null || isNaN(n)) ? 0 : Math.round(n * 100) / 100; };
  var block = function (r) {
    var o = {};
    L.months.forEach(function (m) {
      var w = [];
      for (var i = 0; i < m.weeks.length; i++) w.push(num(vals[r - 1][m.col + i - 1]));
      o[m.ym] = { weeks: w, total: num(vals[r - 1][m.totalCol - 1]) };
    });
    return o;
  };

  return {
    ok: true,
    cats: layout.cats,
    autoCats: Object.keys(SUM_CFG.autoPattern),
    months: L.months.map(function (m) {
      return { ym: m.ym, label: m.label, weeks: m.weeks.map(function (w) { return w.label; }) };
    }),
    members: layout.members.map(function (name, mi) {
      var top = L.memberTop[mi];
      return {
        name: name,
        rows: layout.cats.map(function (_, ci) { return block(top + ci); }),
        sum: block(top + layout.cats.length)
      };
    }),
    grand: block(L.grandRow)
  };
}
