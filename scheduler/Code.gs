/**
 * Large Martian — practice scheduler engine (Google Apps Script)
 *
 * What it does
 *  - Stores settings, members, availability answers and the chosen schedule in a Google Sheet it creates.
 *  - Serves a tiny JSON API that the band room page (largemartian.com/band/) talks to.
 *  - Once a day (8am) checks the date: on the poll day it emails the band to mark unavailable days;
 *    on the deadline day it builds next month's practice schedule, writes it to the Google Calendar,
 *    and emails everyone.
 *
 * One-time setup
 *  1. script.google.com → New project → paste this whole file over Code.gs → save.
 *  2. Run the function `setup` once (choose it in the toolbar, press Run) and approve the permissions.
 *  3. Deploy → New deployment → type "Web app" → Execute as: Me → Who has access: Anyone → Deploy.
 *     Copy the Web app URL. Put it in the band room: band/lock.html → scheduler.url.
 *  4. Done. Everything else is controlled from the band room page.
 */

var TOKEN = 'lm-7f3a9c2e4b1d8e6f';            // must match scheduler.token in band.json
var CALENDAR_ID = 'e03e08705b6d1b2023312919c095d62a5825e9cd38d098cc627939c5643aa3e2@group.calendar.google.com';
var SITE = 'https://largemartian.com/band/';
var TZ = 'America/Los_Angeles';
var TAG = '[LM-auto]';

var DEFAULTS = {
  practicesPerWeek: 1,
  dayOrder: ['Sun', 'Wed', 'Thu', 'Tue', 'Fri', 'Mon', 'Sat'],
  slots: {
    weekday: { start: '18:00', end: '21:00' },
    // weekend sessions, in order of preference; people mark which of these they can't make
    weekend: [
      { id: 'am', label: 'Morning', start: '09:30', end: '12:00' },
      { id: 'pm', label: 'Afternoon', start: '14:00', end: '17:00' },
      { id: 'eve', label: 'Evening', start: '17:00', end: '20:00' }
    ]
  },
  pollDay: 1,
  deadlineDay: 7,
  noAnswer: 'available',          // 'available' | 'skip' (a non-responder blocks nothing / blocks everything)
  minFree: 'all',                 // 'all' = everyone free; or a number = at least that many free
  members: [ { name: 'Ethan', email: 'brooke.ethan@gmail.com' } ],
  notify: true
};

// ---------- storage ----------
function sheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Large Martian practice scheduler');
    props.setProperty('SHEET_ID', ss.getId());
  }
  ['Settings', 'Availability', 'Schedule', 'Log'].forEach(function (n) { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  var s = ss.getSheetByName('Sheet1'); if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  return ss;
}
function getSettings_() {
  var sh = sheet_().getSheetByName('Settings');
  var raw = sh.getRange('A1').getValue();
  var s = {};
  try { s = raw ? JSON.parse(raw) : {}; } catch (e) { s = {}; }
  var out = JSON.parse(JSON.stringify(DEFAULTS));
  Object.keys(s).forEach(function (k) { out[k] = s[k]; });
  if (!out.slots || !out.slots.weekday) out.slots = JSON.parse(JSON.stringify(DEFAULTS.slots));
  if (!Array.isArray(out.slots.weekend)) out.slots.weekend = JSON.parse(JSON.stringify(DEFAULTS.slots.weekend));
  return out;
}
function saveSettings_(s) {
  var cur = getSettings_();
  Object.keys(s).forEach(function (k) { if (k in DEFAULTS) cur[k] = s[k]; });
  sheet_().getSheetByName('Settings').getRange('A1').setValue(JSON.stringify(cur));
  return cur;
}
function log_(msg) { sheet_().getSheetByName('Log').appendRow([new Date(), msg]); }

// Sheets likes to turn '2026-11', '2026-11-08' and '18:00' into dates/times. Store those columns as
// plain text, and normalise anything that was already converted when reading back.
function cellMonth_(v) { return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM') : String(v || '').trim().slice(0, 7); }
function cellDate_(v) { return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '').trim().slice(0, 10); }
function cellTime_(v) { return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'HH:mm') : String(v || '').trim(); }
function appendText_(sh, values) {
  var row = sh.getLastRow() + 1;
  var rng = sh.getRange(row, 1, 1, values.length);
  rng.setNumberFormat('@');
  rng.setValues([values]);
}

// availability rows: month (YYYY-MM), member, JSON dates, updated
function getAvailability_(month) {
  var sh = sheet_().getSheetByName('Availability');
  var rows = sh.getDataRange().getValues();
  var out = {};
  rows.forEach(function (r) { if (cellMonth_(r[0]) === month && r[1]) { try { out[String(r[1])] = JSON.parse(r[2] || '[]'); } catch (e) { out[String(r[1])] = []; } } });
  return out;
}
function saveAvailability_(month, member, dates) {
  var sh = sheet_().getSheetByName('Availability');
  var rows = sh.getDataRange().getValues();
  for (var i = 0; i < rows.length; i++) {
    if (cellMonth_(rows[i][0]) === month && String(rows[i][1]) === member) {
      sh.getRange(i + 1, 3, 1, 2).setValues([[JSON.stringify(dates), new Date().toISOString()]]);
      return;
    }
  }
  appendText_(sh, [month, member, JSON.stringify(dates), new Date().toISOString()]);
}
function getSchedule_(month) {
  var sh = sheet_().getSheetByName('Schedule');
  return sh.getDataRange().getValues().filter(function (r) { return cellMonth_(r[0]) === month; })
    .map(function (r) { return { date: cellDate_(r[1]), start: cellTime_(r[2]), end: cellTime_(r[3]), eventId: String(r[4] || ''), label: String(r[5] || '') }; });
}
function setSchedule_(month, items) {
  var sh = sheet_().getSheetByName('Schedule');
  var rows = sh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 0; i--) if (cellMonth_(rows[i][0]) === month) sh.deleteRow(i + 1);
  items.forEach(function (it) { appendText_(sh, [month, it.date, it.start, it.end, it.eventId || '', it.label || '']); });
}

// ---------- dates ----------
var DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function ymd_(d) { return d.getFullYear() + '-' + pad_(d.getMonth() + 1) + '-' + pad_(d.getDate()); }
function monthDays_(month) { // month 'YYYY-MM' -> array of {date:'YYYY-MM-DD', dow:'Sun', weekend:bool}
  var y = +month.slice(0, 4), m = +month.slice(5, 7);
  var out = [], d = new Date(y, m - 1, 1);
  while (d.getMonth() === m - 1) {
    out.push({ date: ymd_(d), dow: DAYS[d.getDay()], weekend: d.getDay() === 0 || d.getDay() === 6 });
    d.setDate(d.getDate() + 1);
  }
  return out;
}
function nextMonth_(from) { var d = from || new Date(); var y = d.getFullYear(), m = d.getMonth() + 2; if (m > 12) { m = 1; y++; } return y + '-' + pad_(m); }
function weekKey_(dateStr) { // weeks run Mon..Sun, so Sunday is the end of its week
  var p = dateStr.split('-').map(Number); var d = new Date(p[0], p[1] - 1, p[2]);
  var dow = (d.getDay() + 6) % 7; d.setDate(d.getDate() - dow); return ymd_(d);
}

// ---------- the picker ----------
// availability entries are 'YYYY-MM-DD' (whole day) or 'YYYY-MM-DD#slotId' (one weekend session)
function busyFor_(availability, members, settings, date, slotId) {
  var busy = [];
  members.forEach(function (m) {
    var a = availability[m];
    if (!a) { if (settings.noAnswer === 'skip') busy.push(m + ' (no answer)'); return; }
    if (a.indexOf(date) >= 0 || (slotId && a.indexOf(date + '#' + slotId) >= 0)) busy.push(m);
  });
  return busy;
}
function buildSchedule_(month, settings, availability) {
  var members = settings.members.map(function (m) { return m.name; });
  var answered = Object.keys(availability).filter(function (n) { return members.indexOf(n) >= 0; });
  var days = monthDays_(month);
  var need = members.length;
  var rank = {}; settings.dayOrder.forEach(function (d, i) { rank[d] = i; });
  var okWith = function (busy) { return settings.minFree === 'all' ? busy.length === 0 : (need - busy.length) >= Number(settings.minFree); };

  // every candidate (day, session): weekdays have one session, weekends have each configured one
  var scored = days.map(function (d) {
    var options = [];
    if (d.weekend) {
      settings.slots.weekend.forEach(function (sl, i) {
        var busy = busyFor_(availability, members, settings, d.date, sl.id);
        options.push({ slot: sl.id, label: sl.label, start: sl.start, end: sl.end, busy: busy, ok: okWith(busy), slotRank: i });
      });
    } else {
      var busy = busyFor_(availability, members, settings, d.date, null);
      options.push({ slot: null, label: '', start: settings.slots.weekday.start, end: settings.slots.weekday.end, busy: busy, ok: okWith(busy), slotRank: 0 });
    }
    var best = options.filter(function (o) { return o.ok; }).sort(function (a, b) { return a.slotRank - b.slotRank; })[0] || null;
    return { date: d.date, dow: d.dow, weekend: d.weekend, options: options, best: best, ok: !!best, rank: (d.dow in rank) ? rank[d.dow] : 99 };
  });

  var byWeek = {};
  scored.forEach(function (d) { var k = weekKey_(d.date); (byWeek[k] = byWeek[k] || []).push(d); });
  var picked = [];
  Object.keys(byWeek).sort().forEach(function (k) {
    // a week belongs to the month its Sunday falls in, so month edges aren't double-booked or orphaned
    var wk = k.split('-').map(Number); var sun = new Date(wk[0], wk[1] - 1, wk[2] + 6);
    if (ymd_(sun).slice(0, 7) !== month) return;
    var cands = byWeek[k].filter(function (d) { return d.ok && d.rank < 99; }).sort(function (a, b) { return a.rank - b.rank || a.date.localeCompare(b.date); });
    var n = Math.min(Number(settings.practicesPerWeek) || 1, cands.length);
    var chosen = [];
    for (var i = 0; i < cands.length && chosen.length < n; i++) {
      var c = cands[i];
      // avoid back-to-back days when we can
      var adjacent = chosen.some(function (x) { return Math.abs(new Date(x.date) - new Date(c.date)) < 36 * 3600 * 1000; });
      if (adjacent && cands.length - i > n - chosen.length) continue;
      chosen.push(c);
    }
    chosen.forEach(function (c) {
      picked.push({ date: c.date, dow: c.dow, slot: c.best.slot, label: c.best.label, start: c.best.start, end: c.best.end });
    });
  });
  picked.sort(function (a, b) { return a.date.localeCompare(b.date); });
  return { picked: picked, days: scored, answered: answered, members: members };
}

function writeCalendar_(month, picked) {
  var cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error('Calendar not found: ' + CALENDAR_ID);
  var y = +month.slice(0, 4), m = +month.slice(5, 7);
  var from = new Date(y, m - 1, 1), to = new Date(y, m, 1);
  cal.getEvents(from, to).forEach(function (ev) { if ((ev.getDescription() || '').indexOf(TAG) >= 0) ev.deleteEvent(); });
  return picked.map(function (p) {
    var d = p.date.split('-').map(Number);
    var s = p.start.split(':').map(Number), e = p.end.split(':').map(Number);
    var ev = cal.createEvent('Band practice' + (p.label ? ' (' + p.label.toLowerCase() + ')' : ''), new Date(d[0], d[1] - 1, d[2], s[0], s[1]), new Date(d[0], d[1] - 1, d[2], e[0], e[1]),
      { description: 'Large Martian practice. Set automatically from everyone\'s availability. ' + TAG + '\n' + SITE });
    p.eventId = ev.getId();
    return p;
  });
}

function emails_(settings) { return settings.members.map(function (m) { return m.email; }).filter(Boolean); }
function monthName_(month) { var y = +month.slice(0, 4), m = +month.slice(5, 7); return Utilities.formatDate(new Date(y, m - 1, 1), TZ, 'MMMM yyyy'); }
function fmtDay_(dateStr) { var p = dateStr.split('-').map(Number); return Utilities.formatDate(new Date(p[0], p[1] - 1, p[2]), TZ, 'EEE MMM d'); }
function fmtTime_(hm) { var p = hm.split(':').map(Number); var h = p[0] % 12 || 12; return h + (p[1] ? ':' + pad_(p[1]) : '') + (p[0] < 12 ? 'am' : 'pm'); }

function openPoll_(month, settings) {
  if (!settings.notify) return;
  var to = emails_(settings); if (!to.length) return;
  MailApp.sendEmail({
    to: to.join(','),
    subject: 'Large Martian — which ' + monthName_(month) + ' days can\'t you make?',
    htmlBody: '<p>Practice planning for <b>' + monthName_(month) + '</b>.</p>' +
      '<p>Go to the band room and tap the days you <b>can\'t</b> make: <a href="' + SITE + '#scheduler">' + SITE + '#scheduler</a></p>' +
      '<p>Deadline: the ' + settings.deadlineDay + 'th. After that the schedule is set from everyone\'s answers and goes on the calendar.</p>'
  });
  log_('Poll opened for ' + month);
}

function runSchedule_(month, settings) {
  var avail = getAvailability_(month);
  var res = buildSchedule_(month, settings, avail);
  var withEvents = writeCalendar_(month, res.picked);
  setSchedule_(month, withEvents);
  if (settings.notify && emails_(settings).length) {
    var lines = withEvents.length ? withEvents.map(function (p) { return '<li><b>' + fmtDay_(p.date) + '</b>, ' + fmtTime_(p.start) + '–' + fmtTime_(p.end) + (p.label ? ' (' + p.label.toLowerCase() + ')' : '') + '</li>'; }).join('') : '<li>No day worked for everyone — talk it out in the band room.</li>';
    var missing = res.members.filter(function (m) { return res.answered.indexOf(m) < 0; });
    MailApp.sendEmail({
      to: emails_(settings).join(','),
      subject: 'Large Martian — ' + monthName_(month) + ' practice schedule',
      htmlBody: '<p>Practices for <b>' + monthName_(month) + '</b> (on the calendar now):</p><ul>' + lines + '</ul>' +
        (missing.length ? '<p>No answer from: ' + missing.join(', ') + '.</p>' : '') +
        '<p>Band room: <a href="' + SITE + '">' + SITE + '</a></p>'
    });
  }
  log_('Schedule built for ' + month + ': ' + withEvents.map(function (p) { return p.date; }).join(', '));
  return { picked: withEvents, days: res.days, answered: res.answered };
}

// ---------- daily tick (time trigger) ----------
function dailyTick() {
  var s = getSettings_();
  var today = new Date();
  var dom = Number(Utilities.formatDate(today, TZ, 'd'));
  var month = nextMonth_(today);
  var props = PropertiesService.getScriptProperties();
  if (dom === Number(s.pollDay) && props.getProperty('polled-' + month) !== '1') { openPoll_(month, s); props.setProperty('polled-' + month, '1'); }
  if (dom === Number(s.deadlineDay) && props.getProperty('built-' + month) !== '1') { runSchedule_(month, s); props.setProperty('built-' + month, '1'); }
}

function setup() {
  sheet_(); getSettings_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyTick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyTick').timeBased().atHour(8).everyDays(1).inTimezone(TZ).create();
  // touch the calendar so the permission is granted now
  CalendarApp.getCalendarById(CALENDAR_ID);
  log_('setup ran');
  return 'ok — sheet: ' + sheet_().getUrl();
}

// ---------- web API ----------
function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function doGet(e) { return handle_(e && e.parameter ? e.parameter : {}); }
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  return handle_(body);
}
function handle_(q) {
  try {
    if (q.token !== TOKEN) return out_({ ok: false, error: 'bad token' });
    var s = getSettings_();
    var month = q.month || nextMonth_();
    switch (q.action) {
      case 'state': {
        var avail = getAvailability_(month);
        return out_({ ok: true, month: month, settings: s, availability: avail, schedule: getSchedule_(month),
          preview: buildSchedule_(month, s, avail).picked, sheet: sheet_().getUrl() });
      }
      case 'saveSettings': { var ns = saveSettings_(q.settings || {}); return out_({ ok: true, settings: ns }); }
      case 'saveAvailability': { saveAvailability_(month, String(q.member), (q.unavailable || []).map(String)); return out_({ ok: true, availability: getAvailability_(month) }); }
      case 'preview': { var r = buildSchedule_(month, s, getAvailability_(month)); return out_({ ok: true, picked: r.picked, days: r.days, answered: r.answered }); }
      case 'run': { var rr = runSchedule_(month, s); return out_({ ok: true, picked: rr.picked, answered: rr.answered }); }
      case 'openPoll': { openPoll_(month, s); return out_({ ok: true }); }
      default: return out_({ ok: false, error: 'unknown action' });
    }
  } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
}
