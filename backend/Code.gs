/**
 * 9 Tap Tour data service (Google Apps Script web app).
 *
 * Lives inside the "9 Tap Tour Data" Google Sheet (Extensions > Apps Script).
 * The website reads Tour data from it and sends check-ins, scores, setup and results to it.
 *
 * Reading (GET):   <web app URL>?tabs=Tournaments,Pots        -> { ok, data: { Tournaments: [...], Pots: [...] } }
 *                  <web app URL>?tabs=Check-ins&date=2026-11-29  (only that day's rows)
 * Writing (POST):  body is JSON text: { action, key, ... }      -> { ok, ... }
 *                  The site posts with Content-Type text/plain so browsers send it without a CORS preflight.
 *
 * Who can do what
 *   - Anyone can read the public tabs. Phone and email columns are never sent out.
 *   - Anyone can add their own check-in (action "checkin"), because bowlers check in from their phones.
 *   - Everything else needs the staff key (Script property STAFF_KEY, made by running setup() once).
 *     Later the staff key will sit behind the website's staff sign-in, so it never reaches a browser.
 */

const PUBLIC_TABS = ['Tournaments', 'Results', 'Members', 'Jackpot Winners', 'Pots', 'BOTY', 'Recaps', 'Venues', 'Check-ins', 'Scores', 'Settings', 'Schedule'];
const PRIVATE_COLS = ['Phone', 'Email'];           // never sent out by doGet
const SETUP_JSON = 'Setup (site data)';           // full Setup object, so the site gets back exactly what it saved
const VENUE_JSON = 'Format (site data)';

/* ---------- one-time setup ---------- */

// Run this once from the editor (Run > setup). It makes the staff key and shows it in the log.
function setup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('STAFF_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, '').slice(0, 20); props.setProperty('STAFF_KEY', key); }
  Logger.log('Staff key: ' + key);
  Logger.log('Tabs found: ' + SpreadsheetApp.getActive().getSheets().map(s => s.getName()).join(', '));
}

// Run this if the staff key ever leaks. The old key stops working right away.
function newStaffKey() {
  PropertiesService.getScriptProperties().deleteProperty('STAFF_KEY');
  setup();
}

/* ---------- web app entry points ---------- */

function doGet(e) {
  try {
    runFixes();
    const p = (e && e.parameter) || {};
    if (p.ping) return out({ ok: true, now: new Date().toISOString() });
    const want = (p.tabs ? String(p.tabs).split(',') : PUBLIC_TABS).map(s => s.trim()).filter(t => PUBLIC_TABS.indexOf(t) >= 0);
    const data = {};
    want.forEach(t => {
      let rows = readTab(t);
      if (p.date && (t === 'Check-ins' || t === 'Scores' || t === 'Results')) rows = rows.filter(r => r['Date'] === p.date);
      if (p.since && (t === 'Tournaments' || t === 'Schedule')) rows = rows.filter(r => r['Date'] >= p.since);
      data[t] = rows;
    });
    return out({ ok: true, data });
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    runFixes();
    const body = JSON.parse(e.postData.contents || '{}');
    const act = body.action;
    if (act !== 'checkin' && !isStaff(body.key)) return out({ ok: false, error: 'Staff key needed' });
    lock.waitLock(20000);
    const fn = ACTIONS[act];
    if (!fn) return out({ ok: false, error: 'Unknown action: ' + act });
    const res = fn(body) || {};
    return out(Object.assign({ ok: true }, res));
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ---------- actions ---------- */

const ACTIONS = {
  // A bowler checks in from their phone (public). One row per bowler per squad; a repeat replaces the old row.
  checkin(b) {
    const r = b.row || {};
    if (!/^\d{4}-\d\d-\d\d$/.test(r['Date'] || '')) throw new Error('Check-in needs a Date');
    if (!r['Member #'] && !(r['First name'] && r['Last name'])) throw new Error('Check-in needs a member number or a name');
    r['Checked in at'] = r['Checked in at'] || new Date().toISOString();
    const same = x => x['Date'] === r['Date'] && String(x['Squad']) === String(r['Squad']) &&
      (r['Member #'] ? String(x['Member #']) === String(r['Member #']) : (x['First name'] + ' ' + x['Last name']).toLowerCase() === (r['First name'] + ' ' + r['Last name']).toLowerCase());
    // Only the Desk (staff key) marks a bowler paid, and lanes are given out only once paid.
    // A repeat check-in from a phone never undoes a payment the Desk already took.
    const prev = readTab('Check-ins').filter(same).find(x => String(x['Paid'] || '').trim());
    r['Paid'] = prev ? prev['Paid'] : '';
    r['Lane'] = prev ? prev['Lane'] : '';
    deleteWhere('Check-ins', same);
    appendRows('Check-ins', [r]);
    return {};
  },
  // Staff save changes to a day's check-ins (Desk changes lanes, removes bowlers, marks paid...).
  // Only the bowlers sent are replaced and only those in "removed" are deleted, so check-ins from phones are kept.
  saveCheckins(b) {
    const key = r => String(r['Member #']) + ':' + String(r['Squad']);
    const touch = new Set((b.rows || []).concat(b.removed || []).map(key));
    replaceWhere('Check-ins', x => x['Date'] === b.date && touch.has(key(x)), rowsFor(b.rows, b.date));
    return {};
  },
  // Rovers save one squad's scores. Rows: { Squad, Lane, Member #, Bowler, Game 1..4 }.
  saveScores(b) {
    const sq = String(b.squad);
    replaceWhere('Scores', x => x['Date'] === b.date && String(x['Squad']) === sq,
      rowsFor(b.rows, b.date).map(r => Object.assign(r, { 'Squad': b.squad, 'Squad time': b.squadTime || '', 'Updated at': new Date().toISOString(), 'Entered by': b.by || r['Entered by'] || 'staff' })));
    return {};
  },
  // Staff Setup for one tournament day. b.setup is the site's Setup object.
  saveSetup(b) {
    const s = b.setup || {};
    upsert('Tournaments', 'Date', b.date, {
      'Date': b.date, 'Center': b.center, 'City': b.city,
      'First lane': s.first, 'Last lane': s.last, 'Lane order': s.dir, 'Bowlers per lane': s.max,
      'Max per lane when full': s.overflow ? s.overflowMax : '', 'Games bowled': s.games, 'Games that count': s.count,
      'Squad times': (s.squads || []).join(', '), 'Check-in opens (min before squad)': s.openMins,
      'Check-in open': s.open ? 'Yes' : '', 'Closed squads': (s.closed || []).map(i => i + 1).join(', '),
      'Entry fee': s.fee == null ? '' : s.fee, 'Re-entry fee': s.fee2 == null ? '' : s.fee2, 'How to fill': s.fill === 'spread' ? 'One per lane, then wrap' : 'Fill each lane first', 'Setup saved at': new Date().toISOString(),
      [SETUP_JSON]: JSON.stringify(s)
    });
    return {};
  },
  // Saved venues (Setup's venue box). b.venues: [{ name, city, fmt }]
  saveVenues(b) {
    (b.venues || []).forEach(v => {
      const f = v.fmt || {};
      upsert('Venues', 'Venue', v.name, {
        'Venue': v.name, 'City': v.city, 'Lanes (first)': f.first, 'Lanes (last)': f.last, 'Lane order': f.dir,
        'Bowlers per lane': f.max, 'Max per lane when full': f.overflow ? f.overflowMax : '', 'Games bowled': f.games,
        'Games that count': f.count, 'Usual squad times': (f.squads || []).join(', '), 'Check-in opens (min before squad)': f.openMins,
        [VENUE_JSON]: JSON.stringify(f)
      });
    });
    if (b.removed) b.removed.forEach(n => deleteWhere('Venues', x => x['Venue'] === n));
    return {};
  },
  // Staff edit a bowler (Bowlers page) or add a new member. Matched on Member #, then Full name.
  // Staff delete a member on the Bowlers page (a mistake or a duplicate). Only the Members row goes; results keep the name.
  deleteMember(b) { const k = rawKey(b.name); if (k) deleteWhere('Members', r => rawKey(r['Full name']) === k); return {}; },
  saveMember(b) {
    const m = b.member || {};
    const full = m['Full name'] || [m['First name'], m['Last name']].filter(Boolean).join(' ');
    const rows = readTab('Members');
    const was = String(b.was || '').trim().toLowerCase(); // the row's name before a rename on the site
    const hit = (was && rows.find(r => nameKey(r['Full name']) === nameKey(was))) || rows.find(r => m['Member #'] && String(r['Member #']) === String(m['Member #'])) || rows.find(r => r['Full name'] && nameKey(r['Full name']) === nameKey(full)); // "Jr." and "Jr" are the same row
    m['Full name'] = full;
    if (hit) upsert('Members', 'Full name', hit['Full name'], m); else appendRows('Members', [m]);
    sortMembers();
    return {};
  },
  // Setup's Schedule tab: add or change an upcoming tournament date. A moved date drops its old row (never one with results).
  // The schedule lives on the Schedule tab (one row per tournament date). Past dates are never changed from the site.
  saveEvent(b) {
    if (!b.date || b.date < today()) return { ok: false, error: 'Only upcoming dates can be changed.' };
    if (b.was && b.was !== b.date && b.was >= today()) deleteWhere('Schedule', x => x['Date'] === b.was);
    upsert('Schedule', 'Date', b.date, { 'Date': b.date, 'Venue': b.center, 'City': b.city || '', 'Squad times': b.squads || '', 'Event / note': b.note || '', 'Invite only': b.invite ? 'Yes' : '' });
    return {};
  },
  removeEvent(b) { if (b.date && b.date >= today()) deleteWhere('Schedule', x => x['Date'] === b.date); return {}; },
  // Setup's Tour tab (same for every tournament): one row per setting in the Settings tab.
  saveSettings(b) {
    const ss = SpreadsheetApp.getActive(); if (!ss.getSheetByName('Settings')) ss.insertSheet('Settings').getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    Object.keys(b.settings || {}).forEach(k => upsert('Settings', 'Setting', k, { 'Setting': k, 'Value': b.settings[k] }));
    return {};
  },
  // Staff change the pots on Setup. Adds a dated row; the "Current" formulas pick the newest one.
  savePots(b) {
    appendRows('Pots', [{ 'Email sent': b.date || today(), 'For tournament': '', '9-Tap Jackpot': b.jackpot, 'Progressive Pot': b.progressive, 'Email subject': 'Set on the website by staff' }]);
    return {};
  },
  // Finalize day: the day's standings become Results rows, and the winner goes on the Tournaments row.
  finalize(b) {
    const r = b.result || {};
    replaceWhere('Results', x => x['Date'] === b.date, (r.rows || []).map(x => ({
      'Date': b.date, 'Center': r.where, 'Place': String(x[0]), 'Bowler': x[1], 'Handicap': x[2], 'Bonus pins': x[3], 'Total': x[4], 'Prize': x[5], 'Source': 'Website (Finalize day)'
    })));
    upsert('Tournaments', 'Date', b.date, { 'Date': b.date, 'Center': r.where, 'City': r.city, 'Winner': r.name, 'Winning total': r.total, '1st prize': r.won, 'Entries': r.entries, 'Total paid': r.paid, 'Cut (last cash)': r.lowCash, 'Source': 'Website (Finalize day)' });
    return {};
  },
  // Staff undo a Finalize.
  reopen(b) {
    deleteWhere('Results', x => x['Date'] === b.date && x['Source'] === 'Website (Finalize day)');
    upsert('Tournaments', 'Date', b.date, { 'Winner': '', 'Winning total': '', '1st prize': '', 'Entries': '', 'Total paid': '', 'Cut (last cash)': '' });
    return {};
  },
  // Lets the site check a staff key before it shows staff pages.
  checkKey() { return {}; }
};

/* ---------- sheet helpers ---------- */

function isStaff(key) {
  const real = PropertiesService.getScriptProperties().getProperty('STAFF_KEY');
  return !!real && key === real;
}
function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function tz() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function today() { return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd'); }
// Tidy of the Members tab (one time): the City column is removed, header "Nickname / other names" becomes "Nicknames", and each cell keeps only
// names that aren't already the bowler's first or last name ("Steve Nolan" for Steven Nolan -> "Steve").
function nickClean(first, last, str) {
  const k = w => w.toLowerCase().replace(/[^a-z0-9']/g, ''), own = new Set((first + ' ' + last).split(/\s+/).map(k).filter(Boolean).concat(['jr', 'sr', 'ii', 'iii'])), seen = new Set();
  return String(str || '').split(/[,;()\/]+/).map(p => p.split(/\s+/).filter(w => k(w) && !own.has(k(w))).join(' ').trim())
    .filter(p => p && !seen.has(p.toLowerCase()) && seen.add(p.toLowerCase())).join(', ');
}
// One-time: the schedule moves from the Tournaments tab to its own Schedule tab (Greg, Nov 27). Tournaments keeps each
// tournament's setup and results; Schedule is the list of dates the site shows.
function fixSchedule() {
  const ss = SpreadsheetApp.getActive(); if (ss.getSheetByName('Schedule')) return;
  const sh = ss.insertSheet('Schedule'), H = ['Date', 'Venue', 'City', 'Squad times', 'Event / note', 'Invite only'];
  sh.getRange(1, 1, 1, H.length).setValues([H]);
  const rows = readTab('Tournaments').filter(r => r['Date'] && String(r['Squad times'] || '').trim()).sort((a, b) => String(a['Date']).localeCompare(String(b['Date'])));
  if (rows.length) sh.getRange(2, 1, rows.length, H.length).setValues(rows.map(r => [r['Date'], r['Center'] || '', r['City'] || '', r['Squad times'], r['Event / note'] || '', r['Invite only'] || '']));
}
// One-time: past schedule (real dates and venues 2016-2021, 2023-2026 from the old site and the Tour's emails; 2022 isn't known).
// Rows: [Date, Venue, City, Squad times, Event / note, Invite only]. Dates already on the tab are left alone.
const SCHED_PAST = [["2016-01-09","Bowlero Lanes","Lakewood","","Doubles",""],["2016-01-16","Pacific Lanes","","","",""],["2016-01-23","Secoma Lanes","","","",""],["2016-02-06","West Seattle Bowl","","","",""],["2016-02-13","Kent Bowl","","","",""],["2016-02-20","Daffodil Bowl","","","",""],["2016-02-27","Bowlero Lanes","Lakewood","","",""],["2016-03-05","HiLine Lanes","","","",""],["2016-03-12","Skyway Park Bowl","","","",""],["2016-03-19","Narrows Plaza","","","1st Quarter Championship",""],["2016-04-02","Roxbury Lanes","","","",""],["2016-04-09","Secoma Lanes","","","",""],["2016-04-16","Paradise Bowl","","","",""],["2016-04-23","Kent Bowl","","","",""],["2016-04-30","West Seattle Bowl","","","",""],["2016-05-07","Bowlero Lanes","Lakewood","","",""],["2016-05-14","HiLine Lanes","","","",""],["2016-05-28","Secoma Lanes","","","2nd Quarter Championship",""],["2016-06-04","Daffodil Bowl","","","",""],["2016-06-11","Skyway Park Bowl","","","",""],["2016-06-18","Pacific Lanes","","","Doubles",""],["2016-07-09","Roxbury Lanes","","","",""],["2016-07-16","Narrows Plaza","","","",""],["2016-07-23","West Seattle Bowl","","","",""],["2016-07-30","Paradise Bowl","","","",""],["2016-08-06","Bowlero Lanes","Lakewood","","",""],["2016-08-13","HiLine Lanes","","","",""],["2016-08-20","Secoma Lanes","","","",""],["2016-08-27","Skyway Park Bowl","","","",""],["2016-09-03","Kent Bowl","","","3rd Quarter Championship",""],["2016-09-10","Narrows Plaza","","","",""],["2016-09-17","Roxbury Lanes","","","",""],["2016-10-01","Daffodil Bowl","","","",""],["2016-10-08","West Seattle Bowl","","","",""],["2016-10-15","Kenmore Lanes","Kenmore","","",""],["2016-10-22","Bowlero Lanes","Lakewood","","",""],["2016-10-29","HiLine Lanes","","","",""],["2016-11-05","Skyway Park Bowl","","","",""],["2016-11-12","Paradise Bowl","","","",""],["2016-11-19","HiLine Lanes","","","",""],["2016-12-03","Narrows Plaza","","","",""],["2016-12-10","Secoma Lanes","","","",""],["2016-12-17","Bowlero Lanes","Lakewood","","Year End Championship",""],["2017-01-07","Roxbury Lanes","","","Doubles",""],["2017-01-14","Pacific Lanes","","","",""],["2017-01-21","Narrows Plaza","","","",""],["2017-01-27","Gold Coast","","","3-Region Tournament",""],["2017-02-11","Skyway Park Bowl","","","",""],["2017-02-18","Daffodil Bowl","","","",""],["2017-02-25","Secoma Lanes","","","",""],["2017-03-04","HiLine Lanes","","","",""],["2017-03-11","Pacific Lanes","","","",""],["2017-03-18","Kenmore Lanes","Kenmore","","",""],["2017-04-02","Secoma Lanes","","","1st Quarter Championship",""],["2017-04-08","Paradise Bowl","","","",""],["2017-04-15","Skyway Park Bowl","","","",""],["2017-04-22","HiLine Lanes","","","",""],["2017-04-29","Narrows Plaza","","","",""],["2017-05-13","Daffodil Bowl","","","",""],["2017-05-20","Roxbury Lanes","","","",""],["2017-06-03","Bowlero Lanes","Lakewood","","",""],["2017-06-10","HiLine Lanes","","","",""],["2017-06-17","Secoma Lanes","","","Doubles",""],["2017-06-24","Narrows Plaza","","","",""],["2017-07-15","Kenmore Lanes","Kenmore","","",""],["2017-07-22","Pacific Lanes","","","",""],["2017-07-29","Skyway Park Bowl","","","",""],["2017-08-05","Daffodil Bowl","","","",""],["2017-08-12","Paradise Bowl","","","",""],["2017-08-19","Secoma Lanes","","","",""],["2017-08-26","Roxbury Lanes","","","",""],["2017-09-02","Narrows Plaza","","","3rd Quarter Championship",""],["2017-09-16","HiLine Lanes","","","",""],["2017-09-23","Bowlero Lanes","Lakewood","","",""],["2017-09-30","Pacific Lanes","","","",""],["2017-10-07","Acme Bowl","","","",""],["2017-10-14","Bowlero Lanes","Lakewood","","",""],["2017-10-28","Daffodil Bowl","","","",""],["2017-11-04","Skyway Park Bowl","","","",""],["2017-11-11","Paradise Bowl","","","",""],["2017-11-18","Secoma Lanes","","","",""],["2017-11-25","Narrows Plaza","","","",""],["2017-12-02","Bowlero Lanes","Lakewood","","",""],["2017-12-16","HiLine Lanes","","","Year End Championship",""],["2018-01-06","Secoma Lanes","","","Doubles",""],["2018-01-13","Pacific Lanes","","","",""],["2018-01-20","Skyway Park Bowl","","","",""],["2018-01-28","Gold Coast","","","3-Region Tournament",""],["2018-02-03","Bowlero Lanes","Lakewood","","",""],["2018-02-17","Kenmore Lanes","Kenmore","","",""],["2018-02-24","HiLine Lanes","","","",""],["2018-03-03","Hi Joy Bowl","Port Orchard","","",""],["2018-03-10","Secoma Lanes","","","",""],["2018-03-17","Kent Bowl","","","1st Quarter Championship",""],["2018-03-31","Bowlero Lanes","Lakewood","","",""],["2018-04-07","Roxbury Lanes","","","",""],["2018-04-14","Daffodil Bowl","","","",""],["2018-04-21","Pacific Lanes","","","",""],["2018-04-28","Paradise Bowl","","","",""],["2018-05-12","HiLine Lanes","","","",""],["2018-05-19","Skyway Park Bowl","","","",""],["2018-05-26","Secoma Lanes","","","2nd Quarter Championship",""],["2018-06-02","Narrows Plaza","","","",""],["2018-06-09","Kenmore Lanes","Kenmore","","",""],["2018-06-16","Pacific Lanes","","","Doubles",""],["2018-06-23","Daffodil Bowl","","","",""],["2018-07-14","Roxbury Lanes","","","",""],["2018-07-21","Hi Joy Bowl","Port Orchard","","",""],["2018-07-28","Paradise Bowl","","","",""],["2018-08-04","Pacific Lanes","","","",""],["2018-08-11","Secoma Lanes","","","",""],["2018-08-18","Kent Bowl","","","",""],["2018-08-26","HiLine Lanes","","","",""],["2018-09-01","Daffodil Bowl","","","3rd Quarter Championship",""],["2018-09-08","Skyway Park Bowl","","","",""],["2018-09-15","Nob Hill Lanes","","","",""],["2018-09-29","Narrows Plaza","","","",""],["2018-10-06","HiLine Lanes","","","",""],["2018-10-13","Pacific Lanes","","","",""],["2018-10-20","Daffodil Bowl","","","",""],["2018-10-28","Kenmore Lanes","Kenmore","","",""],["2018-11-03","Kent Bowl","","","",""],["2018-11-10","Paradise Bowl","","","",""],["2018-11-17","Secoma Lanes","","","",""],["2018-11-24","Narrows Plaza","","","",""],["2018-12-01","Bowlero Lanes","Lakewood","","",""],["2018-12-08","Pacific Lanes","","","",""],["2018-12-15","Narrows Plaza","","","Year End Championship",""],["2019-01-05","Secoma Lanes","","","Doubles",""],["2019-01-12","Skyway Park Bowl","","","",""],["2019-01-19","Bowlero Lanes","Lakewood","","",""],["2019-01-25","Gold Coast","","","3-Region Tournament",""],["2019-02-02","HiLine Lanes","","","",""],["2019-02-16","Kenmore Lanes","Kenmore","","",""],["2019-02-23","Daffodil Bowl","","","",""],["2019-03-09","Secoma Lanes","","","",""],["2019-03-16","Narrows Plaza","","","",""],["2019-03-23","Kent Bowl","","","1st Quarter Championship",""],["2019-03-30","Pacific Lanes","","","",""],["2019-04-06","Roxbury Lanes","","","",""],["2019-04-13","Bowlero Lanes","Lakewood","","",""],["2019-04-27","HiLine Lanes","","","",""],["2019-05-04","Skyway Park Bowl","","","",""],["2019-05-11","Narrows Plaza","","","",""],["2019-05-18","Daffodil Bowl","","","",""],["2019-05-25","Secoma Lanes","","","2nd Quarter Championship",""],["2019-06-01","Kenmore Lanes","Kenmore","","",""],["2019-06-08","Bowlero Lanes","Lakewood","","",""],["2019-06-15","Kent Bowl","","","Dad's Day Doubles",""],["2019-06-22","Bowlero Lynnwood","","","",""],["2019-07-13","Narrows Plaza","","","",""],["2019-07-20","Secoma Lanes","","","",""],["2019-07-27","HiLine Lanes","","","",""],["2019-08-03","Bowlero South Seattle","","","",""],["2019-08-10","Secoma Lanes","","","",""],["2019-08-24","Skyway Park Bowl","","","",""],["2019-09-01","Secoma Lanes","","","3rd Quarter Championship",""],["2019-09-07","Narrows Plaza","","","",""],["2019-09-14","Kenmore Lanes","Kenmore","","",""],["2019-09-28","Kent Bowl","","","",""],["2019-10-05","Bowlero Lanes","Lakewood","","",""],["2019-10-12","Bowlero Lynnwood","","","",""],["2019-10-20","Bowlero Acme","","","",""],["2019-10-26","Narrows Plaza","","","",""],["2019-11-02","Angel of the Winds","Arlington","","",""],["2019-11-09","Paradise Bowl","","","",""],["2019-11-16","Secoma Lanes","","","",""],["2019-11-23","HiLine Lanes","","","",""],["2019-12-01","Bowlero Lanes","Lakewood","","",""],["2019-12-07","Skyway Park Bowl","","","",""],["2019-12-14","Secoma Lanes","","","Year End Championship",""],["2020-01-04","HiLine Lanes","","","Doubles",""],["2020-01-11","Bowlero Lynnwood","","","",""],["2020-01-18","Secoma Lanes","","","",""],["2020-01-26","Gold Coast","","","3-Region Tournament",""],["2020-02-01","Kenmore Lanes","Kenmore","","",""],["2020-02-15","Narrows Plaza","","","",""],["2020-02-22","Bowlero Acme","","","",""],["2020-02-29","Bowlero Lynnwood","","","",""],["2020-03-07","Skyway Park Bowl","","","",""],["2020-09-20","Narrows Plaza","","","",""],["2020-10-25","Narrows Plaza","","","3rd Quarter Championship",""],["2021-02-21","Narrows Plaza","","","2020 Year End Championship",""],["2021-03-28","Narrows Plaza","","","",""],["2021-04-10","HiLine Lanes","","","",""],["2021-04-17","Kenmore Lanes","Kenmore","","",""],["2021-04-24","Skyway Park Bowl","","","",""],["2021-05-01","Bowlero South Seattle","","","",""],["2021-05-08","HiLine Lanes","","","",""],["2021-05-15","Secoma Lanes","","","",""],["2021-05-29","Secoma Lanes","","","",""],["2021-06-05","Narrows Plaza","","","",""],["2021-06-12","Daffodil Bowl","","","",""],["2021-06-19","Kent Bowl","","","",""],["2021-06-26","HiLine Lanes","","","",""],["2021-07-10","Bowlero Lynnwood","","","",""],["2021-07-17","Bowlero South Seattle","","","",""],["2021-07-24","Kenmore Lanes","Kenmore","","",""],["2021-07-31","Daffodil Bowl","","","",""],["2021-08-07","Narrows Plaza","","","",""],["2021-08-14","Secoma Lanes","","","",""],["2021-08-21","Bowlero Lanes","Lakewood","","",""],["2021-08-28","HiLine Lanes","","","",""],["2021-09-05","Daffodil Bowl","","","",""],["2021-09-18","Bowlero South Seattle","","","",""],["2021-09-25","Narrows Plaza","","","",""],["2021-10-02","Skyway Park Bowl","","","",""],["2021-10-09","Bowlero Lanes","Lakewood","","",""],["2021-10-16","Bowlero Lynnwood","","","",""],["2021-10-23","Kent Bowl","","","",""],["2021-10-30","Kenmore Lanes","Kenmore","","",""],["2021-11-13","Bowlero Lanes","Lakewood","","",""],["2021-11-20","Narrows Plaza","","","",""],["2021-11-27","Daffodil Bowl","","","",""],["2021-12-04","Secoma Lanes","","","",""],["2021-12-11","Skyway Park Bowl","","","",""],["2021-12-18","Bowlero Lanes","Lakewood","","Year End Championship",""],["2023-01-07","Hi Joy Bowl","Port Orchard","","Kick-Off",""],["2023-01-21","Bowlero Lanes","Lakewood","","",""],["2023-01-28","Bowlero Lynnwood","","","",""],["2023-02-05","Kenmore Lanes","Kenmore","","",""],["2023-02-12","Daffodil Bowl","","","",""],["2023-02-18","Skyway Park Bowl","","","",""],["2023-02-25","Bowlero Lanes","Lakewood","","",""],["2023-03-12","Bowlero Lynnwood","","","",""],["2023-03-19","Bowlero Lanes","Lakewood","","",""],["2023-03-26","Kent Bowl","","","",""],["2023-04-02","Secoma Lanes","","","",""],["2023-04-15","Skyway Park Bowl","","","",""],["2023-04-22","Bowlero Lanes","Lakewood","","",""],["2023-04-30","Roxbury Lanes","","","",""],["2023-05-07","Bowlero Lynnwood","","","",""],["2023-05-13","Bowlero Acme","","","",""],["2023-05-21","Daffodil Bowl","","","",""],["2023-05-27","Bowlero Lanes","Lakewood","","",""],["2023-06-03","Skyway Park Bowl","","","",""],["2023-06-11","Secoma Lanes","","","2nd Quarter Championship",""],["2023-06-17","Kenmore Lanes","Kenmore","","",""],["2023-06-24","Hi Joy Bowl","Port Orchard","","",""],["2023-07-16","Bowlero Lynnwood","","","",""],["2023-07-29","Bowlero Lanes","Lakewood","","",""],["2023-08-06","Bowlero Acme","","","",""],["2023-08-12","Secoma Lanes","","","",""],["2023-08-20","Paradise Bowl","","","",""],["2023-08-26","Angel of the Winds","Arlington","","",""],["2023-09-03","Daffodil Bowl","","","",""],["2023-09-17","Kenmore Lanes","Kenmore","","",""],["2023-09-23","Bowlero Lanes","Lakewood","","",""],["2023-09-30","Skyway Park Bowl","","","",""],["2023-10-08","Bowlero Lynnwood","","","",""],["2023-10-15","Bowlero Acme","","","",""],["2023-10-22","Hi Joy Bowl","Port Orchard","","",""],["2023-10-29","Bowlero Paradise","Parkland","","",""],["2023-11-11","Bowlero Lanes","Lakewood","","",""],["2023-11-18","Skyway Park Bowl","","","",""],["2023-11-26","Daffodil Bowl","","","",""],["2023-12-03","Secoma Lanes","","","",""],["2023-12-10","Kenmore Lanes","Kenmore","","",""],["2023-12-17","Bowlero Lanes","Lakewood","","Year End Championship",""],["2024-01-06","Bowlero Lanes","Lakewood","","Kick-Off",""],["2024-01-14","Tech City Bowl","","","",""],["2024-01-21","Paradise Bowl","","","",""],["2024-01-28","Kent Bowl","","","",""],["2024-02-03","Hi Joy Bowl","Port Orchard","","",""],["2024-02-11","Daffodil Bowl","","","",""],["2024-02-18","Bowlero Lynnwood","","","",""],["2024-02-25","Bowlero Acme","","","",""],["2024-03-10","Angel of the Winds","Arlington","","",""],["2024-03-17","Bowlero Paradise","Parkland","","",""],["2024-03-24","Secoma Lanes","","","",""],["2024-03-31","Bowlero Lanes","Lakewood","","1st Quarter Championship",""],["2024-04-07","Tech City Bowl","","","",""],["2024-04-14","Bowlero Lanes","Lakewood","","",""],["2024-04-21","Bowlero Acme","","","",""],["2024-04-27","Kenmore Lanes","Kenmore","","",""],["2024-05-12","Bowlero Lanes","Lakewood","","",""],["2024-05-18","Bowlero Paradise","Parkland","","",""],["2024-05-26","Bowlero Lanes","Lakewood","","",""],["2024-06-02","Daffodil Bowl","","","",""],["2024-06-08","Hi Joy Bowl","Port Orchard","","",""],["2024-06-15","Bowlero Paradise","Parkland","","Father's Day Doubles",""],["2024-06-23","Secoma Lanes","","","2nd Quarter Championship",""],["2024-06-30","Gold Coast","","","Multi-Region Tournament",""],["2024-07-21","Daffodil Bowl","","","",""],["2024-07-27","Bowlero Lynnwood","","","",""],["2024-08-03","Bowlero Lanes","Lakewood","","",""],["2024-08-11","Bowlero Acme","","","",""],["2024-08-17","Secoma Lanes","","","",""],["2024-08-24","Kenmore Lanes","Kenmore","","",""],["2024-09-01","Daffodil Bowl","","","",""],["2024-09-15","Angel of the Winds","Arlington","","",""],["2024-09-21","Bowlero Paradise","Parkland","","",""],["2024-09-29","Hi Joy Bowl","Port Orchard","","",""],["2024-10-06","Bowlero Lanes","Lakewood","","",""],["2024-10-13","Bowlero Lynnwood","","","",""],["2024-10-20","Bowlero Acme","","","",""],["2024-10-27","Kent Bowl","","","",""],["2024-11-09","Bowlero Paradise","Parkland","","",""],["2024-11-17","Kenmore Lanes","Kenmore","","",""],["2024-11-23","Hi Joy Bowl","Port Orchard","","",""],["2024-12-01","Daffodil Bowl","","","",""],["2024-12-08","Secoma Lanes","","","",""],["2024-12-15","Angel of the Winds","Arlington","","",""],["2024-12-21","Bowlero Lanes","Lakewood","","",""],["2025-01-12","Bowlero Paradise","Parkland","","",""],["2025-01-19","Bowlero Acme","","","",""],["2025-01-26","Bowlero Lynnwood","","","",""],["2025-02-01","Bowlero Acme","","","",""],["2025-02-09","Daffodil Bowl","","","",""],["2025-02-16","Bowlero Lanes","Lakewood","","",""],["2025-02-22","Kenmore Lanes","Kenmore","","",""],["2025-03-01","Bowlero Paradise","Parkland","","",""],["2025-03-16","Secoma Lanes","","","",""],["2025-03-22","Hi Joy Bowl","Port Orchard","","",""],["2025-03-29","Bowlero Acme","","","",""],["2025-04-05","Kent Bowl","","","",""],["2025-04-12","Bowlero Lanes","Lakewood","","",""],["2025-04-20","Bowlero Lynnwood","","","",""],["2025-04-27","West Seattle Bowl","","","",""],["2025-05-10","Angel of the Winds","Arlington","","",""],["2025-05-17","Bowlero Paradise","Parkland","","",""],["2025-05-24","Secoma Lanes","","","",""],["2025-06-01","Daffodil Bowl","","","",""],["2025-06-07","Kent Bowl","","","",""],["2025-06-14","Bowlero Lanes","Lakewood","","Doubles",""],["2025-06-22","Bowlero Acme","","","",""],["2025-06-29","Hi Joy Bowl","Port Orchard","","",""],["2025-07-19","Kenmore Lanes","Kenmore","","",""],["2025-07-26","Bowlero Paradise","Parkland","","",""],["2025-08-03","West Seattle Bowl","","","",""],["2025-08-10","Secoma Lanes","","","",""],["2025-08-17","Bowlero Lynnwood","","","",""],["2025-08-24","Kent Bowl","","","",""],["2025-08-30","Daffodil Bowl","","","",""],["2025-09-14","Bowlero Lanes","Lakewood","","",""],["2025-09-21","Bowlero Acme","","","",""],["2025-09-28","Hi Joy Bowl","Port Orchard","","",""],["2025-10-05","Kenmore Lanes","Kenmore","","",""],["2025-10-11","Bowlero Paradise","Parkland","","",""],["2025-10-19","West Seattle Bowl","","","",""],["2025-10-26","Angel of the Winds","Arlington","","",""],["2025-11-08","Bowlero Lanes","Lakewood","","",""],["2025-11-16","Bowlero Acme","","","",""],["2025-11-22","Bowlero Paradise","Parkland","","",""],["2025-11-30","Secoma Lanes","","","",""],["2025-12-07","Bowlero Lynnwood","","","",""],["2025-12-14","Bowlero Lanes","Lakewood","","",""],["2026-01-11","Bowlero Lanes","Lakewood","","",""],["2026-01-17","Kent Bowl","","","",""],["2026-01-25","All-Star Lanes","","","",""],["2026-02-01","Kenmore Lanes","Kenmore","","",""],["2026-02-08","Daffodil Bowl","","","",""],["2026-02-15","Lucky Strike","Lynnwood","","",""],["2026-02-21","Hi Joy Bowl","Port Orchard","","",""],["2026-03-01","Bowlero Acme","","","",""],["2026-03-08","Secoma Lanes","","","",""],["2026-03-15","Bowlero Paradise","Parkland","","",""],["2026-03-29","West Seattle Bowl","","","",""],["2026-04-04","Tower Lanes","","","",""],["2026-04-12","Bowlero Lanes","Lakewood","","",""],["2026-04-25","Hi Joy Bowl","Port Orchard","","",""],["2026-05-02","Angel of the Winds","Arlington","","",""],["2026-05-09","Bowlero Paradise","Parkland","","",""],["2026-05-24","Secoma Lanes","","","",""],["2026-05-30","Bowlero Lynnwood","","","",""],["2026-06-06","Bowlero Lanes","Lakewood","","",""],["2026-06-14","Kenmore Lanes","Kenmore","","",""],["2026-07-19","Bowlero Acme","","","",""],["2026-07-25","Bowlero Lanes","Lakewood","","",""],["2026-08-01","Secoma Lanes","","","",""],["2026-08-02","Secoma Lanes","","","",""],["2026-08-08","West Seattle Bowl","","","",""],["2026-08-15","Lucky Strike","Lynnwood","","",""],["2026-08-22","Westside Lanes","Olympia","","",""],["2026-08-29","Hi Joy Bowl","Port Orchard","","",""],["2026-09-06","Kent Bowl","","","",""],["2026-09-13","Bowlero Acme","","","",""],["2026-09-19","Bowlero Paradise","Parkland","","",""],["2026-10-04","Secoma Lanes","","","",""],["2026-10-11","Bowlero Lanes","Lakewood","","",""],["2026-10-18","West Seattle Bowl","","","",""],["2026-10-25","Kent Bowl","","","",""],["2026-11-01","Kenmore Lanes","Kenmore","10:00 AM, 12:00 PM, 2:00 PM","",""],["2026-11-07","Westside Lanes","Olympia","11:00 AM, 1:30 PM, 4:00 PM","Best 3 of 4",""],["2026-11-15","Bowlero Paradise","Parkland","10:00 AM, 12:30 PM, 3:00 PM","Veteran's Day Special",""],["2026-11-22","Angel of the Winds","Arlington","11:00 AM, 1:30 PM, 4:00 PM","",""],["2026-11-29","Lucky Strike","Lynnwood","11:00 AM, 1:00 PM, 3:00 PM","",""],["2026-12-06","Hi Joy Bowl","Port Orchard","10:00 AM, 12:00 PM, 2:00 PM","",""],["2026-12-13","Bowlero Lanes","Lakewood","9:00 AM","Year End Championship","Yes"]];
// One-time tidy of every tab (Greg, Nov 28): date columns become real dates shown as M/D/YYYY, "... at" columns real
// date-and-times (M/D/YYYY h:mm AM/PM), squad times plain text (so "9:00 AM" isn't turned into a time value).
// The formats are set on the whole column, so rows added later look the same. The Schedule tab is then sorted by date.
function fixDates() {
  const P = PropertiesService.getScriptProperties(); if (P.getProperty('dateFmt1')) return;
  const ymd = x => { const [y, m, d] = x.split('-').map(Number); return new Date(y, m - 1, d); };
  SpreadsheetApp.getActive().getSheets().forEach(s => {
    if (s.getLastRow() < 1 || s.getLastColumn() < 1) return;
    const h = headers(s), n = s.getLastRow() - 1, rows = Math.max(1, s.getMaxRows() - 1);
    const cols = h.map((k, i) => [k, i]).filter(([k]) => DATE_COL.test(k) || STAMP_COL.test(k) || TIME_COL.test(k));
    if (!cols.length) return;
    cols.forEach(([k, i]) => s.getRange(2, i + 1, rows, 1).setNumberFormat(TIME_COL.test(k) ? '@' : DATE_COL.test(k) ? 'M/d/yyyy' : 'M/d/yyyy h:mm AM/PM'));
    if (n < 1) return;
    cols.forEach(([k, i]) => { const rg = s.getRange(2, i + 1, n, 1), f = rg.getFormulas(), v = rg.getValues(); let ch = false;
      const out = v.map((r, j) => { let x = r[0]; if (f[j] && f[j][0]) return [f[j][0]];
        if (DATE_COL.test(k) && typeof x === 'string' && /^\d{4}-\d\d-\d\d$/.test(x.trim())) { x = ymd(x.trim()); ch = true; }
        else if (STAMP_COL.test(k) && typeof x === 'string' && /^\d{4}-\d\d-\d\dT/.test(x.trim()) && !isNaN(Date.parse(x))) { x = new Date(x.trim()); ch = true; }
        else if (TIME_COL.test(k) && x instanceof Date) { x = Utilities.formatDate(x, tz(), 'h:mm a'); ch = true; }
        else if (TIME_COL.test(k) && typeof x === 'number') { x = String(x); ch = true; }
        return [x]; });
      if (ch) rg.setValues(out); });
    if (s.getName() === 'Schedule' && n > 1) { const di = h.indexOf('Date'); if (di >= 0) s.getRange(2, 1, n, h.length).sort([{ column: di + 1, ascending: true }]); }
  });
  P.setProperty('dateFmt1', '1');
}
// One-time tidy-ups. Each runs on its own, so one failing never stops the site from reading or saving.
function runFixes() { [fixMembers, fixSchedule, fixSchedulePast, fixBoty, fixDates].forEach(f => { try { f(); } catch (e) { console.error(f.name + ': ' + e); } }); }
function fixSchedulePast() {
  const s = SpreadsheetApp.getActive().getSheetByName('Schedule'); if (!s) return;
  const have = readTab('Schedule'); if (have.some(r => String(r['Date']) < '2025-01-01')) return;
  const got = new Set(have.map(r => String(r['Date']))), add = SCHED_PAST.filter(r => !got.has(r[0]));
  if (!add.length) return;
  appendRows('Schedule', add.map(r => ({ 'Date': r[0], 'Venue': r[1], 'City': r[2], 'Squad times': r[3], 'Event / note': r[4], 'Invite only': r[5] })));
  const n = s.getLastRow() - 1; if (n > 1) s.getRange(2, 1, n, s.getLastColumn()).sort([{ column: 1, ascending: true }]);
}
// One-time: the real 2026 BOTY standings (as of Aug 4, 2026, top 50 from the Tour's posted image) go on the BOTY tab.
const BOTY_2026 = [["Yanis Green", 482], ["Johnny Seng", 468], ["Dale Roberts Jr", 456], ["Samson Ibale", 437], ["Andy Snow", 431], ["Manuel Mason", 405], ["Paul Kelly", 381], ["Walt Dalrymple", 370], ["Stephanie Lanouette", 369], ["Alex Cady", 363], ["Earl Kleist", 352], ["Thang Duc Thai", 349], ["Steven Nolan", 340], ["Andrew Darr", 325], ["James Black", 319], ["Tina Tran", 304], ["Bill Hunter", 302], ["Keliikuewa Kalima-Marion", 301], ["Jay Pannette", 295], ["Derek Chan", 287], ["Jeffrey Price", 287], ["Joe Fincher Jr", 270], ["Oscar Gestoso", 269], ["Brett Harrison", 265], ["Jeremy Lanouette", 254], ["Red Mateo", 253], ["Rolando Liscano", 232], ["Robert Anderson", 232], ["Kris Koepke", 232], ["Skip Shear Jr", 231], ["Collin Dionne", 230], ["Jeremy Price", 228], ["Rene Malicdan", 227], ["Andrey Williams", 227], ["Terry Farr", 224], ["Marilyn Mateo", 222], ["James Mitchell", 221], ["L.C. Cofield", 216], ["Leroy Cofield", 4], ["Crisanta Williams", 217], ["Tong Nguyen", 212], ["Mark Clark", 210], ["Young Pettengill", 205], ["Thomas Donahue", 203], ["Joel Brown", 200], ["Ashley Kochel", 199], ["Austin Forrest", 197], ["Nicholas Pagdilao", 193], ["Jon Schneider", 192], ["Joshua Brown", 192]];
function fixBoty() {
  const s = SpreadsheetApp.getActive().getSheetByName('BOTY'); if (!s) return;
  if (readTab('BOTY').some(r => String(r['Season']) === '2026')) return;
  appendRows('BOTY', BOTY_2026.slice().sort((a, b) => b[1] - a[1]).map((x, i, A) => ({ 'Season': 2026, 'Rank': A.findIndex(y => y[1] === x[1]) + 1, 'Bowler': x[0], 'Points': x[1], 'As of': '2026-08-04', 'Source': 'Tour BOTY standings image (8/4/2026)' })));
}
function fixMembers() {
  const s = SpreadsheetApp.getActive().getSheetByName('Members'); if (!s) return;
  let h = headers(s); const c = h.indexOf('Nickname / other names');
  if (c >= 0) {
    const n = s.getLastRow() - 1, fi = h.indexOf('First name'), li = h.indexOf('Last name');
    if (n > 0) { const v = s.getRange(2, 1, n, h.length).getValues(); s.getRange(2, c + 1, n, 1).setValues(v.map(r => [nickClean(r[fi], r[li], r[c])])); }
    s.getRange(1, c + 1).setValue('Nicknames');
  }
  const ci = headers(s).indexOf('City'); if (ci >= 0) s.deleteColumn(ci + 1); // the Tour doesn't keep a member's city
  dedupeMembers(s);
  drivePhotos(s);
  sortMembers(s);
}
// Same person twice ("Charles Daughery Jr" / "Charles Daughery Jr."): names that match once case, punctuation and spacing
// are ignored. Keeps one row (the one with a Member #), fills its blanks from the others, joins Nicknames / Found in / Notes,
// and renames the dropped spelling on the other tabs so the title counts still add up. Jr and Sr stay different people.
// Two different Member #s are never merged.
function rawKey(n) { return String(n || '').toLowerCase().replace(/[^a-z0-9 ]/g, '').replace(/\s+/g, ' ').trim(); }
function nameKey(n) { const k = rawKey(n); return SAME_PERSON[k] || k; }
// Different spellings Greg confirmed are one person (name key -> the key of the name to keep).
const SAME_PERSON = { 'jay panette': 'jay pannette' };
// Member photos live in Greg's Drive folder 9taptour/images/members (copied there Nov 27), one per bowler, named
// <first-last>.jpg (a member # is added on the end if two members share a name). Photo links that still point
// at the website's img/ folder are switched to the Drive file.
const PHOTO_FOLDER = 'https://drive.google.com/drive/folders/1FUGpwZD3BHdoAtrH9wQgizEeJ4WF8CY5';
const DRIVE_PHOTOS = {
  'alex-cady': '18wuPSCjEHbKmhzRhN1IYFl_CRlSEUSTs',
  'ashley-kochel': '10CDIVYUVqBMzzsL_zc9RqKhAQhv1PpJM',
  'austin-forrest': '1YUNMWSr7g-Dfw3AdHtMPta5TAhMtxiuV',
  'brandon-parker': '1GqX0aaPLw-APiShBwWFsDFfaHsvOEiyd',
  'brian-ward': '18vWWYuPVq6gbmrApEs_Jor_QdD_CgmVy',
  'chandra-martinez': '15bv51Ff8s0o1h5tQFpl58TCjjUts1Fmp',
  'chanyn-pauole': '1J3Iwx5yVBHVVvSw1Csr8bXcYJb5ObKb0',
  'charles-daughery-jr': '1r-94mCEEIAs9SE08Iejf7RXe0mV84UOr',
  'david-pettengill': '1Sou0YmaY2-470M3Cx0CNlhtiaiOPe_EE',
  'david-sims': '1BY-MjhSJiJl3yn2zIcXYTu_X4nNHJQv0',
  'david-suom': '1oGerKUlJ6TqSuFBbgwLc9j10gEkwIoNz',
  'dean-hale-jr': '1QQAb8YQVt5MdVoxEZQOY5WaOHJQ_u5Dv',
  'dennis-latham': '1ZwF9pwtTm0PDH_HhAWKoF6Td2Md-dItE',
  'eddie-burdsall': '1H_1qCQYFy7fq_LsoqqxkWxMUbgYAdGXs',
  'eric-arkoncel': '1CYpcf_z6QwSLq1wDei1zAtAxwaGWpDjz',
  'james-mitchell': '1GGw91TlOAfb0reC80CbOWOeUVVWB4K0P',
  'joe-fincher-jr': '1mWUynxcPFUC77U7LTN9tbTA-_gXInLcu',
  'joel-brown': '1g5ne9bGSB-1Uqw0T2jakQ-es64IO3UkU',
  'kim-pohlot': '1uZtycYkg0Onuxl4Sa2PO0kv3r9hRJEZi',
  'lii-kalima-marion': '1VEgu8_QgrOnoxFu5F6UqEWFsxHSfe2pF',
  'marilyn-mateo': '1tF7Gn6O5-6GccBcxgVpekulC7-4ZOdoH',
  'michael-miller': '1g57-ufvr1cw6rUA4erk69MBroYGpTOVN',
  'miles-nerona': '1uZMtawq7eH60KVgN74IzMIiNxtQR1PLJ',
  'montana-xaysana': '1tfXk7G-sNSn2raWRFjpmg5rQzXXZHj72',
  'paul-kelly': '1oHvo-C_qHqOvT1XCaj9-BwgyPbf-ixJR',
  'remi-serrano': '1RANvyQcQQxlffHaHydXudDIsuQ66lj-_',
  'rene-malicdan': '107-uXAsmhKQK5MShrNNoaMkd0NBOD4We',
  'rob-hess': '1lYK1dnvQoot32lB0iUs8phpe3KReiCeO',
  'ruta-hert': '1wv7-cRTut5WWV-pqSAzXzKEoBlvdcmCL',
  'samson-ibale': '1NIs3o-RmdjiYsR8lgpwuDo63mdn1LSJU',
  'serina-reese': '1oxGQ--sG-vTKK_AK3QnLnsm5AEjuezYL',
  'skip-shear-jr': '1hr6ll023R0ZUNRj1fV0tzCqDm9r0Gwt0',
  'stephanie-perrine': '1ZiUoaQq7y85fgQ3ddIJA55unuS9_JNRW',
  'steve-nolan': '1CLEdAWfgUI1ANCqh1kg62raitmxtnHat',
  'william-poole': '1OJ-42NzoUvy0k7Gphzlvz0sIivAX0fcF',
  'yoel-green': '1DhEnxj3zHtccdzKmPDnsScIlICBqa-EP'
};
function drivePhotos(s) {
  const h = headers(s), c = h.indexOf('Photo link'), n = s.getLastRow() - 1; if (c < 0 || n < 1) return;
  const r = s.getRange(2, c + 1, n, 1), v = r.getValues(); let ch = false;
  v.forEach(x => { const m = String(x[0] || '').match(/\/img\/(?:members\/)?(?:champ-)?([a-z0-9-]+?)(?:-\d{4}-\d\d-\d\d)?\.jpg/i), id = m && (DRIVE_PHOTOS[m[1].toLowerCase()] || (m[1].toLowerCase() === 'skip-shear' && DRIVE_PHOTOS['skip-shear-jr']));
    if (id) { x[0] = 'https://drive.google.com/file/d/' + id + '/view'; ch = true; } });
  if (ch) r.setValues(v);
}
function dedupeMembers(s) {
  s = s || sheet('Members'); const h = headers(s), n = s.getLastRow() - 1, fu = h.indexOf('Full name'), mi = h.indexOf('Member #');
  if (n < 2 || fu < 0) return;
  const rng = s.getRange(2, 1, n, h.length), vals = rng.getValues(), forms = rng.getFormulas(), groups = {}, kill = [], ren = {};
  const num = i => mi < 0 ? '' : String(vals[i][mi] || '').trim(), blank = v => v === '' || v == null;
  vals.forEach((r, i) => { const k = nameKey(r[fu]); if (k) (groups[k] = groups[k] || []).push(i); });
  Object.keys(groups).forEach(k => {
    const g = groups[k]; if (g.length < 2 || new Set(g.map(num).filter(Boolean)).size > 1) return;
    g.sort((a, b) => (num(b) ? 1 : 0) - (num(a) ? 1 : 0));
    const keep = g[0], row = vals[keep].slice(), names = g.map(i => String(vals[i][fu]).trim()), best = g[[x => rawKey(x) === k && x.indexOf('.') < 0, x => rawKey(x) === k, x => x.indexOf('.') < 0, () => true].map(f => names.findIndex(f)).find(x => x >= 0)];
    g.slice(1).forEach(i => h.forEach((col, c) => {
      if (forms[keep][c] || forms[i][c] || blank(vals[i][c])) return;
      if (blank(row[c])) row[c] = vals[i][c];
      else if (/^(Nicknames|Nickname \/ other names|Found in|Notes)$/.test(col)) row[c] = [...new Set(String(row[c]).split(/\s*,\s*/).concat(String(vals[i][c]).split(/\s*,\s*/)).filter(Boolean))].join(', ');
    }));
    ['First name', 'Last name', 'Full name'].forEach(col => { const c = h.indexOf(col); if (c >= 0) row[c] = vals[best][c]; });
    const nc = h.indexOf('Nicknames') >= 0 ? h.indexOf('Nicknames') : h.indexOf('Nickname / other names'), other = names.filter(x => rawKey(x) !== rawKey(vals[best][fu])); // a different spelling stays findable
    if (nc >= 0 && other.length && !forms[keep][nc]) row[nc] = nickClean(row[h.indexOf('First name')], row[h.indexOf('Last name')], [row[nc]].concat(other).filter(Boolean).join(', '));
    h.forEach((col, c) => { if (!forms[keep][c] && row[c] !== vals[keep][c]) s.getRange(keep + 2, c + 1).setValue(row[c]); });
    names.forEach(x => { if (x !== String(vals[best][fu]).trim()) ren[x] = String(vals[best][fu]).trim(); });
    kill.push(...g.slice(1));
  });
  if (!kill.length) return;
  kill.sort((a, b) => b - a).forEach(i => s.deleteRow(i + 2));
  ['Results', 'Jackpot Winners', 'BOTY', 'Tournaments', 'Scores', 'Check-ins'].forEach(t => {
    const ts = SpreadsheetApp.getActive().getSheetByName(t); if (!ts || ts.getLastRow() < 2) return;
    headers(ts).forEach((col, c) => {
      if (!/^(Bowler|Name|Winner|Full name)$/.test(col)) return;
      const r = ts.getRange(2, c + 1, ts.getLastRow() - 1, 1), v = r.getValues();
      if (r.getFormulas().some(f => f[0])) return;
      let ch = false; v.forEach(x => { const nm = String(x[0]).trim(); if (ren[nm]) { x[0] = ren[nm]; ch = true; } }); if (ch) r.setValues(v);
    });
  });
}
// Members stay A-Z by last name, then first name (only re-sorted when something is out of order).
function sortMembers(s) {
  s = s || sheet('Members'); const h = headers(s), li = h.indexOf('Last name'), fi = h.indexOf('First name'), n = s.getLastRow() - 1;
  if (li < 0 || fi < 0 || n < 2) return;
  const key = r => (String(r[li] || '').trim() + '\u0000' + String(r[fi] || '').trim()).toLowerCase(), rows = s.getRange(2, 1, n, h.length).getValues();
  if (rows.every((r, i) => !i || key(rows[i - 1]) <= key(r))) return;
  s.getRange(2, 1, n, h.length).sort([{ column: li + 1, ascending: true }, { column: fi + 1, ascending: true }]);
}
function sheet(name) { const s = SpreadsheetApp.getActive().getSheetByName(name); if (!s) throw new Error('No tab named ' + name); return s; }
function headers(s) { return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String); }

// Dates come out as yyyy-MM-dd; everything else as the cell shows it (numbers stay numbers).
// Column kinds, by header: dates (M/D/YYYY), date-and-time stamps, and squad times kept as plain text.
const DATE_COL = /^(Date|Tournament date|Email sent|For tournament|As of|Last used)$/, STAMP_COL = / at$/, TIME_COL = /^(Squad times?|Usual squad times)$/;
function cellOut(v, k) {
  if (v instanceof Date && k && STAMP_COL.test(k)) return v.toISOString();
  if (v instanceof Date) return v.getFullYear() < 1901 ? Utilities.formatDate(v, tz(), 'h:mm a') : Utilities.formatDate(v, tz(), 'yyyy-MM-dd'); // a time-only cell (squad time) has no real date
  return v;
}
function readTab(name) {
  const s = sheet(name), h = headers(s), n = s.getLastRow() - 1;
  if (n < 1) return [];
  const vals = s.getRange(2, 1, n, h.length).getValues();
  return vals.filter(r => r.some(v => v !== '' && v !== null)).map(r => {
    const o = {};
    h.forEach((k, i) => { if (k && PRIVATE_COLS.indexOf(k) < 0) o[k] = cellOut(r[i], k); });
    ['Setup (site data)', 'Format (site data)'].forEach(k => { if (o[k]) try { o[k] = JSON.parse(o[k]); } catch (_) {} });
    return o;
  });
}
// Adds any header the site sends that the tab doesn't have yet (at the right end).
function ensureCols(s, keys) {
  const h = headers(s), add = keys.filter(k => k && h.indexOf(k) < 0);
  if (add.length) { s.getRange(1, h.length + 1, 1, add.length).setValues([add]).setFontWeight('bold'); }
  return headers(s);
}
function toRow(h, o, old) {
  return h.map((k, i) => {
    if (!(k in o)) return old ? old[i] : '';
    const v = o[k];
    if (v === null || v === undefined) return '';
    if (typeof v === 'string' && /^\d{4}-\d\d-\d\d$/.test(v) && DATE_COL.test(k)) { const [y, m, d] = v.split('-').map(Number); return new Date(y, m - 1, d); }
    if (typeof v === 'string' && STAMP_COL.test(k) && /^\d{4}-\d\d-\d\dT/.test(v)) return new Date(v);
    return v;
  });
}
function appendRows(name, objs) {
  if (!objs.length) return;
  const s = sheet(name), h = ensureCols(s, Object.keys(Object.assign({}, ...objs)));
  s.getRange(s.getLastRow() + 1, 1, objs.length, h.length).setValues(objs.map(o => toRow(h, o)));
}
// Rows whose formulas must survive (Members, Venues) are only ever updated column by column, never rewritten whole.
function upsert(name, keyCol, keyVal, obj) {
  obj = Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined)); // leave cells the site didn't send alone
  const s = sheet(name), h = ensureCols(s, Object.keys(obj)), kc = h.indexOf(keyCol), n = s.getLastRow() - 1;
  const keys = n > 0 ? s.getRange(2, kc + 1, n, 1).getValues().map(r => cellOut(r[0])) : [];
  const i = keys.findIndex(k => String(k) === String(keyVal));
  if (i < 0) { appendRows(name, [Object.assign({ [keyCol]: keyVal }, obj)]); return; }
  Object.keys(obj).forEach(k => { const c = h.indexOf(k); if (c >= 0) s.getRange(i + 2, c + 1).setValue(toRow([k], { [k]: obj[k] })[0]); });
}
function deleteWhere(name, test) {
  const s = sheet(name), rows = readTabRaw(s);
  for (let i = rows.length - 1; i >= 0; i--) if (test(rows[i])) s.deleteRow(i + 2);
}
function replaceWhere(name, test, objs) { deleteWhere(name, test); appendRows(name, objs); }
function readTabRaw(s) {
  const h = headers(s), n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, h.length).getValues().map(r => { const o = {}; h.forEach((k, i) => o[k] = cellOut(r[i], k)); return o; });
}
function rowsFor(rows, date) { return (rows || []).map(r => Object.assign({}, r, { 'Date': date })); }
