// 출퇴근 기록 시트용 웹 앱. 퇴근 타이머 위젯이 이 주소로 기록을 보내고 받아 간다.
// 확장 프로그램 > Apps Script에 붙여 넣고, 배포 > 새 배포 > 웹 앱
// (실행 사용자: 나, 액세스 권한: 모든 사용자)으로 배포한다.

const SHEET = '기록';
const TYPES = { work: '근무', leave: '연차', half: '반차', holiday: '휴일' };
const TYPE_KEYS = Object.fromEntries(Object.entries(TYPES).map(([k, v]) => [v, k]));

function sheet_() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dates_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, 1).getDisplayValues().map((r) => r[0]) : [];
}

// 오늘 캘린더 일정을 돌려준다 (내가 거절한 일정은 뺀다)
function events_() {
  const events = CalendarApp.getDefaultCalendar().getEventsForDay(new Date())
    .filter((ev) => ev.getMyStatus() !== CalendarApp.GuestStatus.NO)
    .map((ev) => ({
      title: ev.getTitle() || '(제목 없음)',
      start: ev.getStartTime().toISOString(),
      end: ev.getEndTime().toISOString(),
      allDay: ev.isAllDayEvent(),
      location: ev.getLocation() || '',
    }))
    .sort((a, b) => (a.allDay === b.allDay ? a.start.localeCompare(b.start) : a.allDay ? -1 : 1));
  return json_({ ok: true, events });
}

// Claude 세션 대시보드: 로컬 수집 스크립트가 보낸 요약을 'claude-data' 탭 A1에 JSON으로 보관한다
const CLAUDE_TAB = 'claude-data';
const ROUTINE_TAB = '정기 실행';

function saveClaude_(r) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CLAUDE_TAB) || ss.insertSheet(CLAUDE_TAB);
  sh.getRange(1, 1).setValue(JSON.stringify({ generatedAt: r.generatedAt, sessions: r.sessions, usage: r.usage }));
  return json_({ ok: true });
}

// 세션 요약 + '정기 실행' 탭의 루틴별 최근 기록을 함께 돌려준다
function claude_() {
  const ss = SpreadsheetApp.getActive();
  const data = ss.getSheetByName(CLAUDE_TAB);
  let snapshot = null;
  try { snapshot = data ? JSON.parse(data.getRange(1, 1).getValue()) : null; } catch (err) {}
  const log = ss.getSheetByName(ROUTINE_TAB);
  const n = log ? log.getLastRow() - 1 : 0;
  const runs = n > 0 ? log.getRange(2, 1, n, 4).getDisplayValues()
    .filter((x) => x[0]).map((x) => ({ at: x[0], name: x[1], status: x[2], summary: x[3] })) : [];
  return json_({ ok: true, snapshot, runs: runs.slice(-50) });
}

// 기본: 전체 기록을 돌려준다 [{d: "2026-10-13", t: "work", s: "09:30", e: "18:40"}, ...]
// ?action=events: 오늘 캘린더 일정 / ?action=claude: Claude 세션 대시보드 데이터
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'events') return events_();
  if (e && e.parameter && e.parameter.action === 'claude') return claude_();
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, 5).getDisplayValues() : [];
  const records = rows
    .filter((r) => r[0])
    .map((r) => ({ d: r[0], t: TYPE_KEYS[r[2]] || 'work', s: r[3], e: r[4] }));
  return json_({ ok: true, records });
}

// 하루치 기록을 날짜 기준으로 추가하거나 고친다. 출근·퇴근이 모두 비어 있는 근무일은 행을 지운다.
function doPost(e) {
  const r = JSON.parse(e.postData.contents);
  if (r.action === 'claude') return saveClaude_(r);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(r.d || '')) return json_({ ok: false, error: '날짜 형식이 잘못됐어요' });
  const t = TYPES[r.t] ? r.t : 'work';
  const s = /^\d{2}:\d{2}$/.test(r.s || '') ? r.s : '';
  const end = /^\d{2}:\d{2}$/.test(r.e || '') ? r.e : '';
  const empty = t === 'work' && !s && !end;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const dates = dates_(sh);
    let row = dates.indexOf(r.d) + 2;

    if (row < 2) {
      if (empty) return json_({ ok: true });
      // 날짜 순서를 유지하도록 더 늦은 날짜 앞에 끼워 넣는다 (B·F열 등은 1행의 수식이 채운다)
      const after = dates.findIndex((d) => d && d > r.d);
      if (after >= 0) {
        row = after + 2;
        sh.insertRowBefore(row);
      } else {
        row = sh.getLastRow() + 1;
      }
    } else if (empty) {
      sh.deleteRow(row);
      return json_({ ok: true });
    }

    sh.getRange(row, 1).setValue(r.d);
    sh.getRange(row, 3, 1, 3).setValues([[TYPES[t], s, end]]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
