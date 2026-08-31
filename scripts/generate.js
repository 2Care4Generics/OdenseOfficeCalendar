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

const DAYS_BEFORE = 3;  // how many days back from today to fetch
const DAYS_AFTER = 13;  // how many days forward from today to fetch

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
    gap: 12px;
    padding: 12px 20px;
    border-bottom: 1px solid var(--grid-line);
    flex-wrap: wrap;
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
  }
  .nav-btn:hover { background: var(--header-bg); }
  .nav-btn:disabled { opacity: 0.35; cursor: default; }
  .today-btn {
    border: 1px solid var(--grid-line);
    background: #fff;
    border-radius: 4px;
    padding: 0 14px;
    height: 32px;
    font-size: 14px;
    cursor: pointer;
    color: var(--text-primary);
  }
  .today-btn:hover { background: var(--header-bg); }
  .date-label {
    font-size: 17px;
    font-weight: 600;
    margin-left: 4px;
  }
  .updated {
    margin-left: auto;
    font-size: 12px;
    color: var(--text-secondary);
  }
  .calendar-wrap { overflow-x: auto; }
  .calendar {
    display: grid;
    grid-template-columns: 56px repeat(3, minmax(160px, 1fr));
    min-width: 640px;
  }
  .room-header {
    position: sticky;
    top: 0;
    background: #fff;
    z-index: 3;
    border-bottom: 2px solid var(--grid-line);
    padding: 10px 8px;
    font-size: 14px;
    font-weight: 600;
    text-align: center;
  }
  .room-header .dot {
    display: inline-block;
    width: 8px; height: 8px;
    border-radius: 50%;
    margin-right: 6px;
  }
  .gutter-header { border-bottom: 2px solid var(--grid-line); }
  .gutter { position: relative; border-right: 1px solid var(--grid-line); }
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
  }
  .room-col:last-child { border-right: none; }
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
    font-size: 12px;
    line-height: 1.3;
    overflow: hidden;
    color: #201f1e;
  }
  .event-title { font-weight: 600; }
  .event-time { color: var(--text-secondary); font-size: 11px; }
  .empty-state {
    padding: 40px 20px;
    text-align: center;
    color: var(--text-secondary);
    font-size: 13px;
  }
  .error-note {
    padding: 8px;
    font-size: 11px;
    color: var(--now-line);
    text-align: center;
  }
</style>
</head>
<body>

<div class="topbar">
  <button class="nav-btn" id="prevBtn" aria-label="Previous day">&#8249;</button>
  <button class="nav-btn" id="nextBtn" aria-label="Next day">&#8250;</button>
  <button class="today-btn" id="todayBtn">Today</button>
  <span class="date-label" id="dateLabel"></span>
  <span class="updated" id="updatedLabel"></span>
</div>

<div class="calendar-wrap">
  <div class="calendar" id="calendar"></div>
</div>

<script>
const TIMEZONE = ${JSON.stringify(TIMEZONE)};
const GENERATED_AT = ${JSON.stringify(generatedAtISO)};
const ROOMS = ${roomsJson};
const RANGE_DAYS_BEFORE = ${DAYS_BEFORE};
const RANGE_DAYS_AFTER = ${DAYS_AFTER};
const ROW_HEIGHT = 48; // px per hour

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
function weekdayName(y, m, d) {
  return new Intl.DateTimeFormat('en-GB', { weekday: 'long', timeZone: 'UTC' })
    .format(new Date(Date.UTC(y, m - 1, d)));
}
function monthName(y, m, d) {
  return new Intl.DateTimeFormat('en-GB', { month: 'long', timeZone: 'UTC' })
    .format(new Date(Date.UTC(y, m - 1, d)));
}

const nowParts = getLocalDateParts(new Date(), TIMEZONE);
let viewDate = { y: nowParts.y, m: nowParts.m, d: nowParts.d };

const minDate = shiftDate(viewDate, -RANGE_DAYS_BEFORE);
const maxDate = shiftDate(viewDate, RANGE_DAYS_AFTER);

function shiftDate(dp, delta) {
  const t = new Date(Date.UTC(dp.y, dp.m - 1, dp.d + delta));
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}
function sameDate(a, b) { return a.y === b.y && a.m === b.m && a.d === b.d; }
function dateWithinRange(dp) {
  const t = Date.UTC(dp.y, dp.m - 1, dp.d);
  return t >= Date.UTC(minDate.y, minDate.m - 1, minDate.d) && t <= Date.UTC(maxDate.y, maxDate.m - 1, maxDate.d);
}

function formatHourLabel(h) {
  const period = h < 12 ? 'AM' : 'PM';
  let hh = h % 12;
  if (hh === 0) hh = 12;
  return hh + ' ' + period;
}

function render() {
  const dayStart = localWallTimeToUTC(viewDate.y, viewDate.m, viewDate.d, 0, 0, 0, TIMEZONE);
  const dayEnd = localWallTimeToUTC(viewDate.y, viewDate.m, viewDate.d, 23, 59, 59, TIMEZONE);

  // clip each room's events to this day, in local wall-clock hours
  const perRoom = ROOMS.map(room => {
    const clipped = [];
    room.events.forEach(ev => {
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

  // dynamic hour range: default 7-19, expand to fit events
  let minHour = 7, maxHour = 19;
  perRoom.forEach(r => r.events.forEach(ev => {
    minHour = Math.min(minHour, Math.floor(ev.startHour));
    maxHour = Math.max(maxHour, Math.ceil(ev.endHour));
  }));
  minHour = Math.max(0, minHour);
  maxHour = Math.min(24, maxHour);
  const totalHours = maxHour - minHour;
  const gridHeight = totalHours * ROW_HEIGHT;

  const cal = document.getElementById('calendar');
  cal.style.gridTemplateRows = '44px ' + gridHeight + 'px';
  cal.innerHTML = '';

  // top-left empty corner
  const corner = document.createElement('div');
  corner.className = 'gutter-header';
  cal.appendChild(corner);

  // room headers
  perRoom.forEach(r => {
    const h = document.createElement('div');
    h.className = 'room-header';
    h.innerHTML = '<span class="dot" style="background:' + r.room.color + '"></span>' + r.room.name;
    cal.appendChild(h);
  });

  // gutter (hour labels)
  const gutter = document.createElement('div');
  gutter.className = 'gutter';
  gutter.style.height = gridHeight + 'px';
  for (let h = minHour; h <= maxHour; h++) {
    const label = document.createElement('div');
    label.className = 'hour-label';
    label.style.top = ((h - minHour) * ROW_HEIGHT) + 'px';
    label.textContent = formatHourLabel(h);
    gutter.appendChild(label);
  }
  cal.appendChild(gutter);

  const isToday = sameDate(viewDate, nowParts);

  perRoom.forEach(r => {
    const col = document.createElement('div');
    col.className = 'room-col';
    col.style.height = gridHeight + 'px';

    if (r.error) {
      const err = document.createElement('div');
      err.className = 'error-note';
      err.textContent = 'Could not load this calendar';
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
      const height = Math.max(18, (ev.endHour - ev.startHour) * ROW_HEIGHT - 2);
      block.style.top = top + 'px';
      block.style.height = height + 'px';
      block.style.background = r.room.color + '22';
      block.style.borderLeftColor = r.room.color;
      block.innerHTML = '<div class="event-title">' + escapeHtml(ev.summary) + '</div>' +
        '<div class="event-time">' + ev.startLabel + '\u2013' + ev.endLabel + '</div>';
      col.appendChild(block);
    });

    if (isToday) {
      const nowP = getLocalDateParts(new Date(), TIMEZONE);
      const nowHour = nowP.h + nowP.mi / 60;
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

  document.getElementById('dateLabel').textContent =
    weekdayName(viewDate.y, viewDate.m, viewDate.d) + ', ' + viewDate.d + ' ' + monthName(viewDate.y, viewDate.m, viewDate.d) + ' ' + viewDate.y;

  const genParts = getLocalDateParts(new Date(GENERATED_AT), TIMEZONE);
  document.getElementById('updatedLabel').textContent =
    'Updated ' + String(genParts.h).padStart(2,'0') + ':' + String(genParts.mi).padStart(2,'0');

  document.getElementById('prevBtn').disabled = !dateWithinRange(shiftDate(viewDate, -1));
  document.getElementById('nextBtn').disabled = !dateWithinRange(shiftDate(viewDate, 1));
}

function escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

document.getElementById('prevBtn').addEventListener('click', () => {
  const d = shiftDate(viewDate, -1);
  if (dateWithinRange(d)) { viewDate = d; render(); }
});
document.getElementById('nextBtn').addEventListener('click', () => {
  const d = shiftDate(viewDate, 1);
  if (dateWithinRange(d)) { viewDate = d; render(); }
});
document.getElementById('todayBtn').addEventListener('click', () => {
  viewDate = { y: nowParts.y, m: nowParts.m, d: nowParts.d };
  render();
});

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
