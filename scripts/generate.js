/**
 * ROOM CALENDAR DASHBOARD — static generator
 * -------------------------------------------
 * Fetches a window of events for 3 meeting rooms and writes docs/index.html:
 * an Outlook-style day view with all 3 rooms as columns, and prev/next-day
 * navigation handled entirely client-side (no re-fetch needed to browse
 * days already included in the generated window).
 *
 * Run manually:   node scripts/generate.js
 * Run on a schedule via .github/workflows/update-dashboard.yml
 */

const fs = require('fs');
const path = require('path');

// ==================== CONFIG ====================

const TIMEZONE = 'Europe/Copenhagen'; // change if your office is elsewhere

const ROOMS = [
  { name: 'beC',      url: 'https://outlook.office365.com/owa/calendar/1c5af52f99f848fa89a70fe5a779d8bf@pentia.dk/435fe01696e14750934a0e8c132431c812962649515559264857/calendar.ics', color: '#5B5FC7' },
  { name: 'beJava',   url: 'https://outlook.office365.com/owa/calendar/114040dfd4234b0396507d4204e608aa@pentia.dk/77c45a3e85a846fa99b9fbfd2697d04a10085636735793479152/calendar.ics', color: '#0F8B62' },
  { name: 'beSwift',  url: 'https://outlook.office365.com/owa/calendar/bfd0e9abe15346fa8bef788a90fd084a@pentia.dk/001c669c0199496996f5da8cf22899dc17789031396291427944/calendar.ics', color: '#B0280A' },
];

const OUTPUT_FILE = path.join(__dirname, '..', 'docs', 'index.html');

const DAYS_BEFORE = 35; // how many days back from today to fetch
const DAYS_AFTER = 45;  // how many days forward from today to fetch

// =================================================

// ---------- Date/timezone helpers (no external deps) ----------

function formatInTZ(date, tz) {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
  }).formatToParts(date);
  const p = {};
  parts.forEach(x => { p[x.type] = x.value; });
  return p;
}

function getLocalDateParts(date, tz) {
  const p = formatInTZ(date, tz);
  return { y: +p.year, m: +p.month, d: +p.day, h: +p.hour, mi: +p.minute, s: +p.second };
}

function pureDate(y, m, d) {
  return new Date(Date.UTC(y, m - 1, d));
}

function localWallTimeToUTC(y, m, d, h, mi, s) {
  const guess = new Date(Date.UTC(y, m - 1, d, h, mi, s));
  const f = getLocalDateParts(guess, TIMEZONE);
  const guessAsIfUTC = Date.UTC(f.y, f.m - 1, f.d, f.h, f.mi, f.s);
  const offset = guessAsIfUTC - guess.getTime();
  return new Date(guess.getTime() - offset);
}

// ---------- ICS parsing ----------

function unfoldICS(text) {
  return text.replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, '');
}

function parseEventBlock(block) {
  const lines = block.split(/\r?\n/).filter(l => l && l !== 'BEGIN:VEVENT' && l !== 'END:VEVENT');
  const props = {};
  lines.forEach(line => {
    const match = line.match(/^([A-Za-z\-]+)((?:;[^:]*)?):(.*)$/);
    if (!match) return;
    const name = match[1].toUpperCase();
    const paramsRaw = match[2] || '';
    const value = match[3];
    const params = {};
    paramsRaw.split(';').filter(Boolean).forEach(p => {
      const [k, v] = p.split('=');
      if (k) params[k.toUpperCase()] = v;
    });
    props[name] = { value, params };
  });
  return props;
}

function parseICSDateTime(value, params) {
  params = params || {};
  if (/^\d{8}$/.test(value)) {
    const y = +value.substr(0, 4), m = +value.substr(4, 2), d = +value.substr(6, 2);
    return localWallTimeToUTC(y, m, d, 0, 0, 0);
  }
  if (/^\d{8}T\d{6}Z$/.test(value)) {
    const y = +value.substr(0, 4), m = +value.substr(4, 2), d = +value.substr(6, 2);
    const h = +value.substr(9, 2), mi = +value.substr(11, 2), s = +value.substr(13, 2);
    return new Date(Date.UTC(y, m - 1, d, h, mi, s));
  }
  if (/^\d{8}T\d{6}$/.test(value)) {
    const y = +value.substr(0, 4), m = +value.substr(4, 2), d = +value.substr(6, 2);
    const h = +value.substr(9, 2), mi = +value.substr(11, 2), s = +value.substr(13, 2);
    return localWallTimeToUTC(y, m, d, h, mi, s);
  }
  return new Date(value);
}

// ---------- RRULE evaluation (DAILY / WEEKLY / MONTHLY / YEARLY) ----------

const DAY_CODE_TO_NUM = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };
const NUM_TO_DAY_CODE = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];

function startOfWeek(dateOnly, wkstNum) {
  const wd = dateOnly.getUTCDay();
  const diff = (wd - wkstNum + 7) % 7;
  const d = new Date(dateOnly);
  d.setUTCDate(d.getUTCDate() - diff);
  return d;
}

function isNthWeekdayOfMonth(y, m, d, weekdayNum, ord) {
  const dateOnly = pureDate(y, m, d);
  if (dateOnly.getUTCDay() !== weekdayNum) return false;
  if (ord > 0) {
    return Math.floor((d - 1) / 7) + 1 === ord;
  } else {
    const daysInMonth = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return Math.floor((daysInMonth - d) / 7) + 1 === -ord;
  }
}

function expandRRuleForDate(rruleStr, dtStart, exdateProp, dayStart) {
  const rules = {};
  rruleStr.split(';').forEach(p => {
    const [k, v] = p.split('=');
    if (k) rules[k.toUpperCase()] = v;
  });

  const freq = rules.FREQ;
  const interval = rules.INTERVAL ? parseInt(rules.INTERVAL, 10) : 1;
  const until = rules.UNTIL ? parseICSDateTime(rules.UNTIL, {}) : null;

  const startParts = getLocalDateParts(dtStart, TIMEZONE);
  const startDateOnly = pureDate(startParts.y, startParts.m, startParts.d);
  const startWeekday = startDateOnly.getUTCDay();

  const todayParts = getLocalDateParts(dayStart, TIMEZONE);
  const todayDateOnly = pureDate(todayParts.y, todayParts.m, todayParts.d);

  let matches = false;

  if (freq === 'DAILY') {
    const diffDays = Math.round((todayDateOnly - startDateOnly) / 86400000);
    matches = diffDays >= 0 && diffDays % interval === 0;
  } else if (freq === 'WEEKLY') {
    const byDay = rules.BYDAY ? rules.BYDAY.split(',') : [NUM_TO_DAY_CODE[startWeekday]];
    const todayCode = NUM_TO_DAY_CODE[todayDateOnly.getUTCDay()];
    if (byDay.includes(todayCode)) {
      const wkstNum = rules.WKST ? DAY_CODE_TO_NUM[rules.WKST] : 1;
      const startWeekBegin = startOfWeek(startDateOnly, wkstNum);
      const todayWeekBegin = startOfWeek(todayDateOnly, wkstNum);
      const weekDiff = Math.round((todayWeekBegin - startWeekBegin) / (7 * 86400000));
      matches = weekDiff >= 0 && weekDiff % interval === 0;
    }
  } else if (freq === 'MONTHLY') {
    const monthDiff = (todayParts.y - startParts.y) * 12 + (todayParts.m - startParts.m);
    if (monthDiff >= 0 && monthDiff % interval === 0) {
      if (rules.BYDAY) {
        const ord = parseInt(rules.BYDAY, 10);
        const code = rules.BYDAY.replace(/^-?\d+/, '');
        matches = isNthWeekdayOfMonth(todayParts.y, todayParts.m, todayParts.d, DAY_CODE_TO_NUM[code], ord);
      } else {
        matches = todayParts.d === startParts.d;
      }
    }
  } else if (freq === 'YEARLY') {
    const yearDiff = todayParts.y - startParts.y;
    matches = yearDiff >= 0 && yearDiff % interval === 0
      && todayParts.m === startParts.m && todayParts.d === startParts.d;
  }

  if (!matches) return [];

  const occStart = localWallTimeToUTC(todayParts.y, todayParts.m, todayParts.d, startParts.h, startParts.mi, startParts.s);
  if (until && occStart > until) return [];

  if (exdateProp) {
    const exdates = exdateProp.value.split(',').map(v => parseICSDateTime(v.trim(), exdateProp.params).getTime());
    if (exdates.includes(occStart.getTime())) return [];
  }

  return [occStart];
}

// Returns all occurrences (recurring expanded, single events included once)
// whose time overlaps [rangeStart, rangeEnd]
function getOccurrencesInRange(icsText, rangeStart, rangeEnd) {
  const unfolded = unfoldICS(icsText);
  const blocks = unfolded.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || [];
  const occurrences = [];

  blocks.forEach(block => {
    const props = parseEventBlock(block);
    if (!props.DTSTART) return;

    const dtStart = parseICSDateTime(props.DTSTART.value, props.DTSTART.params);
    const dtEndRaw = props.DTEND
      ? parseICSDateTime(props.DTEND.value, props.DTEND.params)
      : new Date(dtStart.getTime() + 60 * 60 * 1000);
    const duration = dtEndRaw.getTime() - dtStart.getTime();
    const summary = props.SUMMARY ? props.SUMMARY.value : '(no title)';

    if (props.RRULE) {
      let dp = getLocalDateParts(rangeStart, TIMEZONE);
      let dayCursor = localWallTimeToUTC(dp.y, dp.m, dp.d, 0, 0, 0);
      let guard = 0;
      while (dayCursor <= rangeEnd && guard < 400) {
        guard++;
        const occStarts = expandRRuleForDate(props.RRULE.value, dtStart, props.EXDATE, dayCursor);
        occStarts.forEach(start => {
          occurrences.push({ start: start.toISOString(), end: new Date(start.getTime() + duration).toISOString(), summary });
        });
        const next = getLocalDateParts(dayCursor, TIMEZONE);
        dayCursor = localWallTimeToUTC(next.y, next.m, next.d + 1, 0, 0, 0);
      }
    } else {
      if (dtStart <= rangeEnd && dtEndRaw >= rangeStart) {
        occurrences.push({ start: dtStart.toISOString(), end: dtEndRaw.toISOString(), summary });
      }
    }
  });

  return occurrences;
}

async function fetchRoomEvents(room, rangeStart, rangeEnd) {
  try {
    const response = await fetch(room.url);
    if (!response.ok) throw new Error('HTTP ' + response.status);
    const icsText = await response.text();
    return { ok: true, events: getOccurrencesInRange(icsText, rangeStart, rangeEnd) };
  } catch (err) {
    return { ok: false, error: err.message, events: [] };
  }
}

// ---------- HTML template (calendar rendering happens client-side) ----------

function renderPage(roomsData, generatedAtISO) {
  const roomsJson = JSON.stringify(roomsData).replace(/</g, '\\u003c');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Room Calendar</title>
<style>
  :root {
    --grid-line: #e1dfdd;
    --text-primary: #201f1e;
    --text-secondary: #605e5c;
    --header-bg: #faf9f8;
    --accent: #2564cf;
    --now-line: #d13438;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: "Segoe UI", -apple-system, Roboto, Arial, sans-serif;
    background: #ffffff;
    color: var(--text-primary);
  }
  .topbar {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 10px 20px;
    border-bottom: 1px solid var(--grid-line);
    flex-wrap: wrap;
    row-gap: 8px;
  }
  .nav-btn {
    border: 1px solid var(--grid-line);
    background: #fff;
    border-radius: 4px;
    width: 32px;
    height: 32px;
    font-size: 16px;
    cursor: pointer;
    color: var(--text-primary);
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .nav-btn:hover, .view-btn:hover { background: var(--header-bg); }
  .nav-btn:disabled { opacity: 0.35; cursor: default; }
  .date-input {
    border: 1px solid var(--grid-line);
    border-radius: 4px;
    height: 32px;
    padding: 0 8px;
    font-size: 13px;
    color: var(--text-primary);
    background: #fff;
  }
  .date-label {
    font-size: 17px;
    font-weight: 600;
    margin-left: 4px;
    white-space: nowrap;
  }
  .view-toggle {
    display: flex;
    border: 1px solid var(--grid-line);
    border-radius: 4px;
    overflow: hidden;
  }
  .view-btn {
    background: #fff;
    border: none;
    border-right: 1px solid var(--grid-line);
    padding: 0 12px;
    height: 32px;
    font-size: 13px;
    cursor: pointer;
    color: var(--text-primary);
  }
  .view-btn:last-child { border-right: none; }
  .view-btn.active { background: var(--accent); color: #fff; }
  .room-filters {
    display: flex;
    gap: 6px;
    flex-wrap: wrap;
  }
  .room-chip {
    display: flex;
    align-items: center;
    gap: 5px;
    border: 1px solid var(--grid-line);
    border-radius: 999px;
    padding: 4px 10px 4px 8px;
    font-size: 12px;
    cursor: pointer;
    background: #fff;
    color: var(--text-primary);
    user-select: none;
  }
  .room-chip .dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
  .room-chip.off { opacity: 0.4; }
  .room-chip.off .dot { background: transparent !important; border: 1.5px solid var(--text-secondary); }
  .updated {
    margin-left: auto;
    font-size: 12px;
    color: var(--text-secondary);
    white-space: nowrap;
  }
  .calendar-wrap { overflow-x: auto; }

  /* ---- Day / Week time grid ---- */
  .calendar {
    display: grid;
    min-width: 640px;
  }
  .day-group-header {
    grid-row: 1;
    border-bottom: 1px solid var(--grid-line);
    padding: 8px 6px;
    font-size: 13px;
    font-weight: 600;
    text-align: center;
    cursor: pointer;
    white-space: nowrap;
    background: #fff;
  }
  .day-group-header:hover { background: var(--header-bg); }
  .day-group-header .wd { color: var(--text-secondary); font-weight: 500; margin-right: 5px; }
  .day-group-header.is-today { color: var(--accent); }
  .room-header {
    grid-row: 2;
    background: #fff;
    border-bottom: 2px solid var(--grid-line);
    padding: 6px 4px;
    font-size: 11px;
    font-weight: 600;
    text-align: center;
  }
  .room-header .dot { display: inline-block; width: 7px; height: 7px; border-radius: 50%; margin-right: 4px; }
  .gutter-header { grid-row: 1 / span 2; border-bottom: 2px solid var(--grid-line); }
  .gutter { position: relative; border-right: 1px solid var(--grid-line); grid-row: 3; }
  .gutter .hour-label {
    position: absolute;
    right: 8px;
    transform: translateY(-50%);
    font-size: 11px;
    color: var(--text-secondary);
  }
  .room-col {
    position: relative;
    border-right: 1px solid var(--grid-line);
    grid-row: 3;
  }
  .day-divider { border-right: 3px solid #8a8886 !important; }
  .hour-line {
    position: absolute;
    left: 0; right: 0;
    border-top: 1px solid var(--grid-line);
  }
  .now-line {
    position: absolute;
    left: 0; right: 0;
    height: 0;
    border-top: 2px solid var(--now-line);
    z-index: 2;
  }
  .now-dot {
    position: absolute;
    left: -4px; top: -4px;
    width: 8px; height: 8px;
    border-radius: 50%;
    background: var(--now-line);
  }
  .event-block {
    position: absolute;
    left: 3px; right: 3px;
    border-radius: 5px;
    border-left: 3px solid;
    padding: 3px 6px;
    font-size: 11px;
    line-height: 1.3;
    overflow: hidden;
    color: #201f1e;
  }
  .event-title { font-weight: 600; }
  .event-time { color: var(--text-secondary); font-size: 10px; }
  .error-note {
    padding: 8px;
    font-size: 11px;
    color: var(--now-line);
    text-align: center;
  }

  /* ---- Month view ---- */
  .month-grid {
    display: grid;
    grid-template-columns: repeat(7, minmax(120px, 1fr));
    min-width: 840px;
  }
  .month-weekday {
    padding: 8px 10px;
    font-size: 12px;
    font-weight: 600;
    color: var(--text-secondary);
    border-bottom: 2px solid var(--grid-line);
    text-align: left;
  }
  .month-cell {
    border-right: 1px solid var(--grid-line);
    border-bottom: 1px solid var(--grid-line);
    min-height: 110px;
    padding: 6px;
    cursor: pointer;
    vertical-align: top;
  }
  .month-cell:hover { background: var(--header-bg); }
  .month-cell.outside { background: #fbfbfb; color: var(--text-secondary); }
  .month-cell .day-num {
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 4px;
  }
  .month-cell.is-today .day-num {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 22px; height: 22px;
    border-radius: 50%;
    background: var(--accent);
    color: #fff;
  }
  .month-chip {
    font-size: 10.5px;
    border-radius: 3px;
    border-left: 3px solid;
    padding: 1px 4px;
    margin-bottom: 2px;
    overflow: hidden;
    white-space: nowrap;
    text-overflow: ellipsis;
  }
  .month-more { font-size: 10.5px; color: var(--text-secondary); }

  /* ---- Sidebar ---- */
  .sidebar-backdrop {
    position: fixed; inset: 0;
    background: rgba(0,0,0,0.25);
    z-index: 10;
  }
  .sidebar {
    position: fixed;
    top: 0; right: 0;
    bottom: 0;
    width: min(340px, 90vw);
    background: #fff;
    box-shadow: -4px 0 16px rgba(0,0,0,0.15);
    z-index: 11;
    display: flex;
    flex-direction: column;
  }
  .sidebar-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 16px 16px 12px;
    border-bottom: 1px solid var(--grid-line);
  }
  .sidebar-header h2 { font-size: 15px; margin: 0; }
  .sidebar-close {
    border: none;
    background: none;
    font-size: 20px;
    line-height: 1;
    cursor: pointer;
    color: var(--text-secondary);
    padding: 4px;
  }
  .sidebar-body { padding: 12px 16px; overflow-y: auto; }
  .sidebar-empty { color: var(--text-secondary); font-size: 13px; padding: 12px 0; }
  .sidebar-item {
    border-left: 3px solid;
    border-radius: 4px;
    background: var(--header-bg);
    padding: 8px 10px;
    margin-bottom: 8px;
  }
  .sidebar-item .room-tag {
    font-size: 10.5px;
    font-weight: 700;
    text-transform: none;
    color: var(--text-secondary);
    margin-bottom: 2px;
  }
  .sidebar-item .title { font-size: 13px; font-weight: 600; }
  .sidebar-item .time { font-size: 12px; color: var(--text-secondary); margin-top: 2px; }

  /* ---- Single-event detail modal ---- */
  .modal-backdrop {
    position: fixed; inset: 0;
    background: rgba(0,0,0,0.35);
    z-index: 20;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .event-modal {
    background: #fff;
    border-radius: 8px;
    border-left: 5px solid;
    box-shadow: 0 8px 28px rgba(0,0,0,0.25);
    width: min(380px, 90vw);
    padding: 20px 22px;
  }
  .event-modal .modal-top {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 10px;
  }
  .event-modal .room-tag {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    font-size: 12px;
    font-weight: 700;
    color: var(--text-secondary);
  }
  .event-modal .room-tag .dot { width: 8px; height: 8px; border-radius: 50%; }
  .event-modal .modal-close {
    border: none;
    background: none;
    font-size: 20px;
    line-height: 1;
    cursor: pointer;
    color: var(--text-secondary);
    padding: 2px;
  }
  .event-modal .modal-title {
    font-size: 17px;
    font-weight: 700;
    margin-bottom: 10px;
    line-height: 1.35;
  }
  .event-modal .modal-row {
    font-size: 13.5px;
    color: var(--text-primary);
    margin-bottom: 4px;
  }
  .event-modal .modal-row .label { color: var(--text-secondary); }
</style>
</head>
<body>

<div class="topbar">
  <button class="nav-btn" id="prevBtn" aria-label="Previous">&#8249;</button>
  <button class="nav-btn" id="nextBtn" aria-label="Next">&#8250;</button>
  <input type="date" class="date-input" id="dateInput">
  <span class="date-label" id="dateLabel"></span>

  <div class="view-toggle" id="viewToggle">
    <button class="view-btn" data-mode="day">Day</button>
    <button class="view-btn" data-mode="week">Week</button>
    <button class="view-btn" data-mode="month">Month</button>
  </div>

  <div class="room-filters" id="roomFilters"></div>

  <span class="updated">
    <span id="updatedLabel"></span> · refreshes in <span id="refreshCountdown"></span>
  </span>
</div>

<div class="calendar-wrap" id="calendarWrap"></div>

<script>
const TIMEZONE = ${JSON.stringify(TIMEZONE)};
const GENERATED_AT = ${JSON.stringify(generatedAtISO)};
const ROOMS = ${roomsJson};
const RANGE_DAYS_BEFORE = ${DAYS_BEFORE};
const RANGE_DAYS_AFTER = ${DAYS_AFTER};
const ROW_HEIGHT = 44; // px per hour

// ---------- date helpers ----------

function formatInTZ(date, tz) {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
  }).formatToParts(date);
  const p = {};
  parts.forEach(x => { p[x.type] = x.value; });
  return p;
}
function getLocalDateParts(date, tz) {
  const p = formatInTZ(date, tz);
  return { y: +p.year, m: +p.month, d: +p.day, h: +p.hour, mi: +p.minute, s: +p.second };
}
function localWallTimeToUTC(y, m, d, h, mi, s, tz) {
  const guess = new Date(Date.UTC(y, m - 1, d, h, mi, s));
  const f = getLocalDateParts(guess, tz);
  const guessAsIfUTC = Date.UTC(f.y, f.m - 1, f.d, f.h, f.mi, f.s);
  const offset = guessAsIfUTC - guess.getTime();
  return new Date(guess.getTime() - offset);
}
function weekdayName(y, m, d, style) {
  return new Intl.DateTimeFormat('en-GB', { weekday: style || 'long', timeZone: 'UTC' })
    .format(new Date(Date.UTC(y, m - 1, d)));
}
function monthName(y, m, d, style) {
  return new Intl.DateTimeFormat('en-GB', { month: style || 'long', timeZone: 'UTC' })
    .format(new Date(Date.UTC(y, m - 1, d)));
}
function shiftDate(dp, delta) {
  const t = new Date(Date.UTC(dp.y, dp.m - 1, dp.d + delta));
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}
function shiftMonth(dp, delta) {
  const t = new Date(Date.UTC(dp.y, dp.m - 1 + delta, 1));
  const daysInMonth = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0)).getUTCDate();
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: Math.min(dp.d, daysInMonth) };
}
function sameDate(a, b) { return a.y === b.y && a.m === b.m && a.d === b.d; }
function dateToKey(dp) { return dp.y + '-' + String(dp.m).padStart(2,'0') + '-' + String(dp.d).padStart(2,'0'); }
function weekdayNum(dp) { return new Date(Date.UTC(dp.y, dp.m - 1, dp.d)).getUTCDay(); } // 0=Sun
function startOfWeekMon(dp) {
  const wd = weekdayNum(dp);
  const diff = (wd === 0) ? 6 : wd - 1; // days since Monday
  return shiftDate(dp, -diff);
}
function isoWeekNumber(dp) {
  // Standard ISO 8601 rule: week 1 is the week containing the year's first
  // Thursday; weeks start Monday. Using the Thursday of the target week
  // correctly handles weeks that span a year boundary.
  const target = new Date(Date.UTC(dp.y, dp.m - 1, dp.d));
  const dayNum = (target.getUTCDay() + 6) % 7; // Mon=0..Sun=6
  target.setUTCDate(target.getUTCDate() - dayNum + 3);
  const firstThursday = new Date(Date.UTC(target.getUTCFullYear(), 0, 4));
  const firstThursdayDayNum = (firstThursday.getUTCDay() + 6) % 7;
  firstThursday.setUTCDate(firstThursday.getUTCDate() - firstThursdayDayNum + 3);
  return 1 + Math.round((target - firstThursday) / (7 * 86400000));
}
function dateWithinFetchedRange(dp) {
  const t = Date.UTC(dp.y, dp.m - 1, dp.d);
  return t >= Date.UTC(minDate.y, minDate.m - 1, minDate.d) && t <= Date.UTC(maxDate.y, maxDate.m - 1, maxDate.d);
}
function formatHourLabel(h) {
  const period = h < 12 ? 'AM' : 'PM';
  let hh = h % 12;
  if (hh === 0) hh = 12;
  return hh + ' ' + period;
}
function escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ---------- state ----------
//
// View state (mode, date, room filters) is persisted in sessionStorage so
// that the auto-refresh reload (and manual browser refreshes) only update
// the underlying data, not what the person was looking at. sessionStorage
// is scoped to this one browser tab and clears when the tab is closed, so
// a genuinely new visit still gets the intended "Week view, today" default.
const STORAGE_KEY = 'roomCalendarViewState';

function loadSavedState() {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return null;
    return parsed;
  } catch (e) {
    return null; // storage unavailable (e.g. private browsing) — fall back to defaults
  }
}
function saveState() {
  try {
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify({
      viewMode,
      viewDate,
      visibleRooms: [...visibleRooms],
    }));
  } catch (e) {
    // storage unavailable — nothing to do, state just won't persist across reloads
  }
}

const nowParts = getLocalDateParts(new Date(), TIMEZONE);
const savedState = loadSavedState();

const validModes = ['day', 'week', 'month'];
const validRoomNames = new Set(ROOMS.map(r => r.name));

let viewMode = (savedState && validModes.includes(savedState.viewMode))
  ? savedState.viewMode
  : 'week'; // first-ever load in this tab defaults to Week view

let viewDate = (savedState && savedState.viewDate && Number.isInteger(savedState.viewDate.y)
    && Number.isInteger(savedState.viewDate.m) && Number.isInteger(savedState.viewDate.d))
  ? savedState.viewDate
  : { y: nowParts.y, m: nowParts.m, d: nowParts.d };

let visibleRooms = (savedState && Array.isArray(savedState.visibleRooms) && savedState.visibleRooms.some(n => validRoomNames.has(n)))
  ? new Set(savedState.visibleRooms.filter(n => validRoomNames.has(n)))
  : new Set(ROOMS.map(r => r.name));

const minDate = shiftDate({ y: nowParts.y, m: nowParts.m, d: nowParts.d }, -RANGE_DAYS_BEFORE);
const maxDate = shiftDate({ y: nowParts.y, m: nowParts.m, d: nowParts.d }, RANGE_DAYS_AFTER);

// ---------- event lookup ----------

// Returns events (across visible rooms) that overlap the given local date,
// clipped to that date's boundaries, each tagged with its room.
function eventsForDate(dp) {
  const dayStart = localWallTimeToUTC(dp.y, dp.m, dp.d, 0, 0, 0, TIMEZONE);
  const dayEnd = localWallTimeToUTC(dp.y, dp.m, dp.d, 23, 59, 59, TIMEZONE);
  const result = [];
  ROOMS.forEach(room => {
    if (!visibleRooms.has(room.name)) return;
    (room.events || []).forEach(ev => {
      const s = new Date(ev.start), e = new Date(ev.end);
      if (s <= dayEnd && e >= dayStart) {
        const clipStart = s < dayStart ? dayStart : s;
        const clipEnd = e > dayEnd ? dayEnd : e;
        const sp = getLocalDateParts(clipStart, TIMEZONE);
        const ep = getLocalDateParts(clipEnd, TIMEZONE);
        result.push({
          room: room.name,
          color: room.color,
          summary: ev.summary,
          startHour: sp.h + sp.mi / 60,
          endHour: ep.h + ep.mi / 60 + (ep.h === 23 && ep.mi === 59 ? (60 - ep.s / 60) / 60 : 0),
          startLabel: String(sp.h).padStart(2, '0') + ':' + String(sp.mi).padStart(2, '0'),
          endLabel: String(ep.h).padStart(2, '0') + ':' + String(ep.mi).padStart(2, '0'),
        });
      }
    });
  });
  result.sort((a, b) => a.startHour - b.startHour);
  return result;
}

// ---------- sidebar ----------

function openSidebar(dp) {
  closeSidebar();
  const backdrop = document.createElement('div');
  backdrop.className = 'sidebar-backdrop';
  backdrop.id = 'sidebarBackdrop';
  backdrop.addEventListener('click', closeSidebar);

  const panel = document.createElement('div');
  panel.className = 'sidebar';
  panel.id = 'sidebarPanel';

  const header = document.createElement('div');
  header.className = 'sidebar-header';
  const h2 = document.createElement('h2');
  h2.textContent = weekdayName(dp.y, dp.m, dp.d) + ', ' + dp.d + ' ' + monthName(dp.y, dp.m, dp.d) + ' ' + dp.y;
  const closeBtn = document.createElement('button');
  closeBtn.className = 'sidebar-close';
  closeBtn.setAttribute('aria-label', 'Close');
  closeBtn.textContent = '\\u00d7';
  closeBtn.addEventListener('click', closeSidebar);
  header.appendChild(h2);
  header.appendChild(closeBtn);

  const body = document.createElement('div');
  body.className = 'sidebar-body';
  const events = eventsForDate(dp);
  if (events.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'sidebar-empty';
    empty.textContent = 'No meetings booked in any visible room.';
    body.appendChild(empty);
  } else {
    events.forEach(ev => {
      const item = document.createElement('div');
      item.className = 'sidebar-item';
      item.style.borderLeftColor = ev.color;
      item.innerHTML =
        '<div class="room-tag">' + escapeHtml(ev.room) + '</div>' +
        '<div class="title">' + escapeHtml(ev.summary) + '</div>' +
        '<div class="time">' + ev.startLabel + '\\u2013' + ev.endLabel + '</div>';
      body.appendChild(item);
    });
  }

  panel.appendChild(header);
  panel.appendChild(body);
  document.body.appendChild(backdrop);
  document.body.appendChild(panel);
}
function closeSidebar() {
  const b = document.getElementById('sidebarBackdrop');
  const p = document.getElementById('sidebarPanel');
  if (b) b.remove();
  if (p) p.remove();
}

// ---------- single-event detail modal ----------

function openEventDetail(ev, dp) {
  closeEventDetail();
  const backdrop = document.createElement('div');
  backdrop.className = 'modal-backdrop';
  backdrop.id = 'eventModalBackdrop';
  backdrop.addEventListener('click', (e) => {
    if (e.target === backdrop) closeEventDetail();
  });

  const modal = document.createElement('div');
  modal.className = 'event-modal';
  modal.style.borderLeftColor = ev.color;

  const top = document.createElement('div');
  top.className = 'modal-top';
  const roomTag = document.createElement('span');
  roomTag.className = 'room-tag';
  roomTag.innerHTML = '<span class="dot" style="background:' + ev.color + '"></span>' + escapeHtml(ev.room);
  const closeBtn = document.createElement('button');
  closeBtn.className = 'modal-close';
  closeBtn.setAttribute('aria-label', 'Close');
  closeBtn.textContent = '\\u00d7';
  closeBtn.addEventListener('click', closeEventDetail);
  top.appendChild(roomTag);
  top.appendChild(closeBtn);

  const dateStr = weekdayName(dp.y, dp.m, dp.d) + ', ' + dp.d + ' ' + monthName(dp.y, dp.m, dp.d) + ' ' + dp.y;

  const title = document.createElement('div');
  title.className = 'modal-title';
  title.textContent = ev.summary;

  const dateRow = document.createElement('div');
  dateRow.className = 'modal-row';
  dateRow.innerHTML = '<span class="label">Date: </span>' + escapeHtml(dateStr);

  const timeRow = document.createElement('div');
  timeRow.className = 'modal-row';
  timeRow.innerHTML = '<span class="label">Time: </span>' + ev.startLabel + '\\u2013' + ev.endLabel;

  modal.appendChild(top);
  modal.appendChild(title);
  modal.appendChild(dateRow);
  modal.appendChild(timeRow);

  backdrop.appendChild(modal);
  document.body.appendChild(backdrop);
}
function closeEventDetail() {
  const b = document.getElementById('eventModalBackdrop');
  if (b) b.remove();
}

// ---------- time-grid render (day & week share this) ----------

function renderTimeGrid(days) {
  const wrap = document.getElementById('calendarWrap');
  const rooms = ROOMS.filter(r => visibleRooms.has(r.name));
  const cal = document.createElement('div');
  cal.className = 'calendar';
  const colCount = 1 + days.length * Math.max(rooms.length, 1);
  cal.style.gridTemplateColumns = '56px repeat(' + (days.length * Math.max(rooms.length, 1)) + ', minmax(' + (days.length > 1 ? 90 : 160) + 'px, 1fr))';

  // gather all events per day/room to compute dynamic hour range
  const perDay = days.map(dp => {
    const dayStart = localWallTimeToUTC(dp.y, dp.m, dp.d, 0, 0, 0, TIMEZONE);
    const dayEnd = localWallTimeToUTC(dp.y, dp.m, dp.d, 23, 59, 59, TIMEZONE);
    const perRoom = rooms.map(room => {
      const clipped = [];
      (room.events || []).forEach(ev => {
        const s = new Date(ev.start), e = new Date(ev.end);
        if (s <= dayEnd && e >= dayStart) {
          const clipStart = s < dayStart ? dayStart : s;
          const clipEnd = e > dayEnd ? dayEnd : e;
          const sp = getLocalDateParts(clipStart, TIMEZONE);
          const ep = getLocalDateParts(clipEnd, TIMEZONE);
          clipped.push({
            summary: ev.summary,
            startHour: sp.h + sp.mi / 60,
            endHour: ep.h + ep.mi / 60 + (ep.h === 23 && ep.mi === 59 ? (60 - ep.s / 60) / 60 : 0),
            startLabel: String(sp.h).padStart(2, '0') + ':' + String(sp.mi).padStart(2, '0'),
            endLabel: String(ep.h).padStart(2, '0') + ':' + String(ep.mi).padStart(2, '0'),
          });
        }
      });
      return { room, events: clipped, error: room.error };
    });
    return { dp, perRoom };
  });

  let minHour = 7, maxHour = 19;
  perDay.forEach(d => d.perRoom.forEach(r => r.events.forEach(ev => {
    minHour = Math.min(minHour, Math.floor(ev.startHour));
    maxHour = Math.max(maxHour, Math.ceil(ev.endHour));
  })));
  minHour = Math.max(0, minHour);
  maxHour = Math.min(24, maxHour);
  const totalHours = maxHour - minHour;
  const gridHeight = totalHours * ROW_HEIGHT;

  cal.style.gridTemplateRows = '30px 26px ' + gridHeight + 'px';

  // top-left corner
  const corner = document.createElement('div');
  corner.className = 'gutter-header';
  corner.style.gridColumn = '1';
  cal.appendChild(corner);

  // day headers + room mini-headers
  perDay.forEach((d, dayIdx) => {
    const roomSpan = Math.max(rooms.length, 1);
    const startCol = 2 + dayIdx * roomSpan;

    const dayHeader = document.createElement('div');
    dayHeader.className = 'day-group-header day-divider' + (sameDate(d.dp, nowParts) ? ' is-today' : '');
    dayHeader.style.gridColumn = startCol + ' / span ' + roomSpan;
    dayHeader.innerHTML = '<span class="wd">' + weekdayName(d.dp.y, d.dp.m, d.dp.d, 'short') + '</span>' + d.dp.d;
    if (days.length > 1) {
      dayHeader.addEventListener('click', () => openSidebar(d.dp));
    }
    cal.appendChild(dayHeader);

    rooms.forEach((room, i) => {
      const rh = document.createElement('div');
      rh.className = 'room-header' + (i === rooms.length - 1 ? ' day-divider' : '');
      rh.style.gridColumn = String(startCol + i);
      rh.innerHTML = '<span class="dot" style="background:' + room.color + '"></span>' + room.name;
      cal.appendChild(rh);
    });
  });

  // gutter
  const gutter = document.createElement('div');
  gutter.className = 'gutter';
  gutter.style.gridColumn = '1';
  gutter.style.height = gridHeight + 'px';
  for (let h = minHour; h <= maxHour; h++) {
    const label = document.createElement('div');
    label.className = 'hour-label';
    label.style.top = ((h - minHour) * ROW_HEIGHT) + 'px';
    label.textContent = formatHourLabel(h);
    gutter.appendChild(label);
  }
  cal.appendChild(gutter);

  // room columns
  perDay.forEach((d, dayIdx) => {
    const roomSpan = Math.max(rooms.length, 1);
    const startCol = 2 + dayIdx * roomSpan;
    const isLastRoomOfDay = (i) => i === rooms.length - 1;

    if (rooms.length === 0) {
      const col = document.createElement('div');
      col.className = 'room-col day-divider';
      col.style.gridColumn = String(startCol);
      col.style.height = gridHeight + 'px';
      cal.appendChild(col);
      return;
    }

    d.perRoom.forEach((r, i) => {
      const col = document.createElement('div');
      col.className = 'room-col' + (isLastRoomOfDay(i) ? ' day-divider' : '');
      col.style.gridColumn = String(startCol + i);
      col.style.height = gridHeight + 'px';

      if (r.error) {
        const err = document.createElement('div');
        err.className = 'error-note';
        err.textContent = 'Could not load';
        col.appendChild(err);
      }

      for (let h = minHour; h <= maxHour; h++) {
        const line = document.createElement('div');
        line.className = 'hour-line';
        line.style.top = ((h - minHour) * ROW_HEIGHT) + 'px';
        col.appendChild(line);
      }

      r.events.forEach(ev => {
        const block = document.createElement('div');
        block.className = 'event-block';
        const top = (ev.startHour - minHour) * ROW_HEIGHT;
        const height = Math.max(16, (ev.endHour - ev.startHour) * ROW_HEIGHT - 2);
        block.style.top = top + 'px';
        block.style.height = height + 'px';
        block.style.background = r.room.color + '22';
        block.style.borderLeftColor = r.room.color;
        block.style.cursor = 'pointer';
        block.innerHTML = '<div class="event-title">' + escapeHtml(ev.summary) + '</div>' +
          '<div class="event-time">' + ev.startLabel + '\\u2013' + ev.endLabel + '</div>';
        block.addEventListener('click', (e) => {
          e.stopPropagation();
          openEventDetail({ room: r.room.name, color: r.room.color, summary: ev.summary, startLabel: ev.startLabel, endLabel: ev.endLabel }, d.dp);
        });
        col.appendChild(block);
      });

      if (sameDate(d.dp, nowParts)) {
        const nowHour = nowParts.h + nowParts.mi / 60;
        if (nowHour >= minHour && nowHour <= maxHour) {
          const line = document.createElement('div');
          line.className = 'now-line';
          line.style.top = ((nowHour - minHour) * ROW_HEIGHT) + 'px';
          const dot = document.createElement('div');
          dot.className = 'now-dot';
          line.appendChild(dot);
          col.appendChild(line);
        }
      }

      cal.appendChild(col);
    });
  });

  wrap.innerHTML = '';
  wrap.appendChild(cal);
}

// ---------- month render ----------

function renderMonthGrid(anchor) {
  const wrap = document.getElementById('calendarWrap');
  const grid = document.createElement('div');
  grid.className = 'month-grid';

  const weekdays = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
  weekdays.forEach(wd => {
    const h = document.createElement('div');
    h.className = 'month-weekday';
    h.textContent = wd;
    grid.appendChild(h);
  });

  const firstOfMonth = { y: anchor.y, m: anchor.m, d: 1 };
  const gridStart = startOfWeekMon(firstOfMonth);
  const rooms = ROOMS.filter(r => visibleRooms.has(r.name));

  for (let i = 0; i < 42; i++) {
    const dp = shiftDate(gridStart, i);
    const cell = document.createElement('div');
    const outside = dp.m !== anchor.m;
    cell.className = 'month-cell' + (outside ? ' outside' : '') + (sameDate(dp, nowParts) ? ' is-today' : '');

    const num = document.createElement('div');
    num.className = 'day-num';
    num.textContent = dp.d;
    cell.appendChild(num);

    const events = eventsForDate(dp).filter(ev => rooms.some(r => r.name === ev.room));
    const maxChips = 3;
    events.slice(0, maxChips).forEach(ev => {
      const chip = document.createElement('div');
      chip.className = 'month-chip';
      chip.style.borderLeftColor = ev.color;
      chip.style.background = ev.color + '1a';
      chip.style.cursor = 'pointer';
      chip.textContent = ev.startLabel + ' ' + ev.summary;
      chip.addEventListener('click', (e) => {
        e.stopPropagation();
        openEventDetail(ev, dp);
      });
      cell.appendChild(chip);
    });
    if (events.length > maxChips) {
      const more = document.createElement('div');
      more.className = 'month-more';
      more.textContent = '+' + (events.length - maxChips) + ' more';
      cell.appendChild(more);
    }

    cell.addEventListener('click', () => openSidebar(dp));
    grid.appendChild(cell);

    if (i === 41) break;
  }

  wrap.innerHTML = '';
  wrap.appendChild(grid);
}

// ---------- top-level render / controls ----------

function render() {
  closeSidebar();
  closeEventDetail();

  if (viewMode === 'day') {
    renderTimeGrid([viewDate]);
    document.getElementById('dateLabel').textContent =
      weekdayName(viewDate.y, viewDate.m, viewDate.d) + ', ' + viewDate.d + ' ' + monthName(viewDate.y, viewDate.m, viewDate.d) + ' ' + viewDate.y;
  } else if (viewMode === 'week') {
    const weekStart = startOfWeekMon(viewDate);
    const days = [];
    for (let i = 0; i < 7; i++) days.push(shiftDate(weekStart, i));
    renderTimeGrid(days);
    const weekEnd = days[6];
    const sameMonth = weekStart.m === weekEnd.m;
    const rangeLabel = sameMonth
      ? weekStart.d + '\\u2013' + weekEnd.d + ' ' + monthName(weekEnd.y, weekEnd.m, weekEnd.d) + ' ' + weekEnd.y
      : weekStart.d + ' ' + monthName(weekStart.y, weekStart.m, weekStart.d, 'short') + ' \\u2013 ' + weekEnd.d + ' ' + monthName(weekEnd.y, weekEnd.m, weekEnd.d, 'short') + ' ' + weekEnd.y;
    document.getElementById('dateLabel').textContent = 'Week ' + isoWeekNumber(weekStart) + ' \\u00b7 ' + rangeLabel;
  } else if (viewMode === 'month') {
    renderMonthGrid(viewDate);
    document.getElementById('dateLabel').textContent = monthName(viewDate.y, viewDate.m, viewDate.d) + ' ' + viewDate.y;
  }

  const genParts = getLocalDateParts(new Date(GENERATED_AT), TIMEZONE);
  document.getElementById('updatedLabel').textContent =
    'Updated ' + String(genParts.h).padStart(2,'0') + ':' + String(genParts.mi).padStart(2,'0');

  document.getElementById('dateInput').value = dateToKey(viewDate);

  const prevTarget = viewMode === 'month' ? shiftMonth(viewDate, -1) : shiftDate(viewDate, viewMode === 'week' ? -7 : -1);
  const nextTarget = viewMode === 'month' ? shiftMonth(viewDate, 1) : shiftDate(viewDate, viewMode === 'week' ? 7 : 1);
  document.getElementById('prevBtn').disabled = viewMode !== 'month' && !dateWithinFetchedRange(prevTarget);
  document.getElementById('nextBtn').disabled = viewMode !== 'month' && !dateWithinFetchedRange(nextTarget);

  document.querySelectorAll('.view-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.mode === viewMode);
  });

  saveState();
}

// ---------- wire up controls ----------

document.getElementById('prevBtn').addEventListener('click', () => {
  viewDate = viewMode === 'month' ? shiftMonth(viewDate, -1) : shiftDate(viewDate, viewMode === 'week' ? -7 : -1);
  render();
});
document.getElementById('nextBtn').addEventListener('click', () => {
  viewDate = viewMode === 'month' ? shiftMonth(viewDate, 1) : shiftDate(viewDate, viewMode === 'week' ? 7 : 1);
  render();
});
document.getElementById('dateInput').addEventListener('change', (e) => {
  const [y, m, d] = e.target.value.split('-').map(Number);
  if (y && m && d) { viewDate = { y, m, d }; render(); }
});
function hardReload() {
  // Cache-busting: a plain location.reload() can be served a cached copy
  // by the browser or GitHub Pages' CDN. Appending a unique query string
  // forces a genuine network fetch of the latest committed file.
  location.href = location.pathname + '?_=' + Date.now();
}
document.getElementById('viewToggle').addEventListener('click', (e) => {
  const btn = e.target.closest('.view-btn');
  if (!btn) return;
  viewMode = btn.dataset.mode;
  render();
});

const filterWrap = document.getElementById('roomFilters');
ROOMS.forEach(room => {
  const chip = document.createElement('div');
  chip.className = 'room-chip';
  chip.innerHTML = '<span class="dot" style="background:' + room.color + '"></span>' + room.name;
  chip.addEventListener('click', () => {
    if (visibleRooms.has(room.name)) {
      if (visibleRooms.size > 1) visibleRooms.delete(room.name); // keep at least one room visible
    } else {
      visibleRooms.add(room.name);
    }
    chip.classList.toggle('off', !visibleRooms.has(room.name));
    render();
  });
  filterWrap.appendChild(chip);
});

// Auto-refresh: reload the page periodically to pick up whatever the
// GitHub Action last committed. A visible countdown makes it clear this
// is actually running, rather than the page silently going stale.
//
// The countdown is anchored to GENERATED_AT (embedded at build time), not
// to when this particular page load happened — so a manual browser
// refresh, or the auto-refresh itself, doesn't reset the clock back to
// a full cycle. It always reflects time remaining in the real cycle.
//
// This matches the external cron-job.org schedule (every 1 minute) that
// triggers the Action via workflow_dispatch, not GitHub's own built-in
// schedule trigger.
const AUTO_REFRESH_MINUTES = 1;
const RETRY_SECONDS_WHEN_OVERDUE = 15; // polling interval if the Action is running late
const generatedAtMs = new Date(GENERATED_AT).getTime();

function computeSecondsUntilRefresh() {
  const elapsedSeconds = (Date.now() - generatedAtMs) / 1000;
  const remaining = AUTO_REFRESH_MINUTES * 60 - elapsedSeconds;
  // If the Action is running behind schedule, don't hammer reloads —
  // just retry at a fixed cadence until fresh content actually shows up.
  return remaining > 0 ? Math.ceil(remaining) : RETRY_SECONDS_WHEN_OVERDUE;
}

let secondsUntilRefresh = computeSecondsUntilRefresh();
function tickCountdown() {
  const mm = Math.floor(secondsUntilRefresh / 60);
  const ss = secondsUntilRefresh % 60;
  const el = document.getElementById('refreshCountdown');
  if (el) el.textContent = mm + ':' + String(ss).padStart(2, '0');
  if (secondsUntilRefresh <= 0) {
    hardReload();
    return;
  }
  secondsUntilRefresh--;
}
tickCountdown();
setInterval(tickCountdown, 1000);

render();
</script>
</body>
</html>`;
}

// ---------- Main ----------

async function main() {
  const now = new Date();
  const nowParts = getLocalDateParts(now, TIMEZONE);
  const rangeStart = localWallTimeToUTC(nowParts.y, nowParts.m, nowParts.d - DAYS_BEFORE, 0, 0, 0);
  const rangeEnd = localWallTimeToUTC(nowParts.y, nowParts.m, nowParts.d + DAYS_AFTER, 23, 59, 59);

  const roomsData = [];
  for (const room of ROOMS) {
    const result = await fetchRoomEvents(room, rangeStart, rangeEnd);
    roomsData.push({
      name: room.name,
      color: room.color,
      events: result.events,
      error: result.ok ? null : result.error,
    });
  }

  const html = renderPage(roomsData, now.toISOString());
  fs.mkdirSync(path.dirname(OUTPUT_FILE), { recursive: true });
  fs.writeFileSync(OUTPUT_FILE, html, 'utf8');
  console.log('Wrote', OUTPUT_FILE);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
