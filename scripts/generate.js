/**
 * ROOM STATUS DASHBOARD — static generator
 * -----------------------------------------
 * Fetches live ICS feeds for 3 meeting rooms, computes free/busy status,
 * and writes docs/index.html for GitHub Pages to serve.
 *
 * Run manually:   node scripts/generate.js
 * Run on a schedule via .github/workflows/update-dashboard.yml
 */

const fs = require('fs');
const path = require('path');

// ==================== CONFIG ====================

const TIMEZONE = 'Europe/Copenhagen'; // change if your office is elsewhere

const ROOMS = [
  { name: 'beC',      url: 'https://outlook.office365.com/owa/calendar/1c5af52f99f848fa89a70fe5a779d8bf@pentia.dk/435fe01696e14750934a0e8c132431c812962649515559264857/calendar.ics', color: '#4F46E5' },
  { name: 'beJava',   url: 'https://outlook.office365.com/owa/calendar/114040dfd4234b0396507d4204e608aa@pentia.dk/77c45a3e85a846fa99b9fbfd2697d04a10085636735793479152/calendar.ics', color: '#059669' },
  { name: 'beSwift',  url: 'https://outlook.office365.com/owa/calendar/bfd0e9abe15346fa8bef788a90fd084a@pentia.dk/001c669c0199496996f5da8cf22899dc17789031396291427944/calendar.ics', color: '#DC2626' },
];

const OUTPUT_FILE = path.join(__dirname, '..', 'docs', 'index.html');

// How often the *page* itself checks for a fresh copy from GitHub Pages
// (the underlying data only actually changes as often as the Action runs)
const PAGE_REFRESH_SECONDS = 60;

// =================================================

// ---------- Date/timezone helpers (no external deps) ----------

function formatInTZ(date, tz) {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
  }).formatToParts(date);
  const p = {};
  parts.forEach(x => { p[x.type] = x.value; });
  return p; // { year, month, day, hour, minute, second }
}

function getLocalDateParts(date, tz) {
  const p = formatInTZ(date, tz);
  return {
    y: +p.year, m: +p.month, d: +p.day,
    h: +p.hour, mi: +p.minute, s: +p.second,
  };
}

function pureDate(y, m, d) {
  return new Date(Date.UTC(y, m - 1, d));
}

// Converts a wall-clock time in TIMEZONE into the correct UTC instant
function localWallTimeToUTC(y, m, d, h, mi, s) {
  const guess = new Date(Date.UTC(y, m - 1, d, h, mi, s));
  const f = getLocalDateParts(guess, TIMEZONE);
  const guessAsIfUTC = Date.UTC(f.y, f.m - 1, f.d, f.h, f.mi, f.s);
  const offset = guessAsIfUTC - guess.getTime();
  return new Date(guess.getTime() - offset);
}

function formatTime(date, tz) {
  const p = getLocalDateParts(date, tz);
  return String(p.h).padStart(2, '0') + ':' + String(p.mi).padStart(2, '0');
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

function getTodaysOccurrences(icsText, now) {
  const unfolded = unfoldICS(icsText);
  const blocks = unfolded.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || [];

  const nowParts = getLocalDateParts(now, TIMEZONE);
  const dayStart = localWallTimeToUTC(nowParts.y, nowParts.m, nowParts.d, 0, 0, 0);
  const dayEnd = localWallTimeToUTC(nowParts.y, nowParts.m, nowParts.d, 23, 59, 59);

  let occurrences = [];

  blocks.forEach(block => {
    const props = parseEventBlock(block);
    if (!props.DTSTART) return;

    const dtStart = parseICSDateTime(props.DTSTART.value, props.DTSTART.params);
    const dtEndRaw = props.DTEND
      ? parseICSDateTime(props.DTEND.value, props.DTEND.params)
      : new Date(dtStart.getTime() + 60 * 60 * 1000);
    const duration = dtEndRaw.getTime() - dtStart.getTime();
    const summary = props.SUMMARY ? props.SUMMARY.value : '';

    if (props.RRULE) {
      const occStarts = expandRRuleForDate(props.RRULE.value, dtStart, props.EXDATE, dayStart);
      occStarts.forEach(start => {
        occurrences.push({ start, end: new Date(start.getTime() + duration), summary });
      });
    } else {
      if (dtStart <= dayEnd && dtEndRaw >= dayStart) {
        occurrences.push({ start: dtStart, end: dtEndRaw, summary });
      }
    }
  });

  return occurrences;
}

// ---------- Fetch + compute status ----------

async function getRoomStatus(room, now) {
  let events = [];
  let error = null;
  try {
    const response = await fetch(room.url);
    if (!response.ok) throw new Error('HTTP ' + response.status);
    const icsText = await response.text();
    events = getTodaysOccurrences(icsText, now);
  } catch (err) {
    error = err.message;
  }

  events.sort((a, b) => a.start - b.start);

  const current = events.find(ev => ev.start <= now && now < ev.end);
  let status, detail;

  if (error) {
    status = 'error';
    detail = 'Could not load calendar (' + error + ')';
  } else if (current) {
    status = 'busy';
    detail = (current.summary || 'Busy') + ' — until ' + formatTime(current.end, TIMEZONE);
  } else {
    const next = events.find(ev => ev.start > now);
    status = 'free';
    detail = next ? 'Free until ' + formatTime(next.start, TIMEZONE) : 'Free for the rest of the day';
  }

  return { name: room.name, color: room.color, status, detail, events };
}

function escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function renderDashboard(roomStatuses, now) {
  const parts = formatInTZ(now, TIMEZONE);
  const nowStr = `${parts.day}/${parts.month}/${parts.year} ${parts.hour}:${parts.minute}`;

  const cards = roomStatuses.map(r => {
    const bg = r.status === 'busy' ? '#FEE2E2' : r.status === 'error' ? '#F3F4F6' : '#DCFCE7';
    const badgeColor = r.status === 'busy' ? '#DC2626' : r.status === 'error' ? '#6B7280' : '#16A34A';
    const badgeText = r.status === 'busy' ? 'BUSY' : r.status === 'error' ? 'UNKNOWN' : 'FREE';

    const upcoming = r.events
      .filter(ev => ev.end > now)
      .slice(0, 4)
      .map(ev => `<li>${formatTime(ev.start, TIMEZONE)}–${formatTime(ev.end, TIMEZONE)} ${escapeHtml(ev.summary || '')}</li>`)
      .join('');

    return `
      <div class="card" style="background:${bg}; border-top: 6px solid ${r.color};">
        <div class="card-header">
          <span class="room-name">${escapeHtml(r.name)}</span>
          <span class="badge" style="background:${badgeColor}">${badgeText}</span>
        </div>
        <div class="detail">${escapeHtml(r.detail)}</div>
        ${upcoming ? `<ul class="upcoming">${upcoming}</ul>` : ''}
      </div>`;
  }).join('');

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta http-equiv="refresh" content="${PAGE_REFRESH_SECONDS}">
  <title>Room Status</title>
  <style>
    body { font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; background:#F9FAFB; margin:0; padding:24px; }
    h1 { font-size: 20px; color:#111827; margin-bottom: 4px; }
    .timestamp { color:#6B7280; font-size: 13px; margin-bottom: 24px; }
    .grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap:16px; max-width: 900px; }
    .card { border-radius: 12px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
    .card-header { display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; }
    .room-name { font-size: 18px; font-weight:600; color:#111827; }
    .badge { color:white; font-size:11px; font-weight:700; padding:3px 8px; border-radius:999px; letter-spacing:0.5px; }
    .detail { font-size: 14px; color:#374151; margin-bottom: 8px; }
    .upcoming { margin:0; padding-left:18px; font-size:12px; color:#6B7280; }
    .upcoming li { margin-bottom:2px; }
  </style>
</head>
<body>
  <h1>Meeting Room Status</h1>
  <div class="timestamp">Updated ${nowStr} (Europe/Copenhagen) · page checks for a refresh every ${PAGE_REFRESH_SECONDS}s</div>
  <div class="grid">
    ${cards}
  </div>
</body>
</html>`;
}

// ---------- Main ----------

async function main() {
  const now = new Date();
  const roomStatuses = [];
  for (const room of ROOMS) {
    roomStatuses.push(await getRoomStatus(room, now));
  }
  const html = renderDashboard(roomStatuses, now);
  fs.mkdirSync(path.dirname(OUTPUT_FILE), { recursive: true });
  fs.writeFileSync(OUTPUT_FILE, html, 'utf8');
  console.log('Wrote', OUTPUT_FILE);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
