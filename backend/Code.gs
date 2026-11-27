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

const PUBLIC_TABS = ['Tournaments', 'Results', 'Members', 'Jackpot Winners', 'Pots', 'BOTY', 'Recaps', 'Venues', 'Check-ins', 'Scores', 'Settings'];
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
    fixMembers();
    const p = (e && e.parameter) || {};
    if (p.ping) return out({ ok: true, now: new Date().toISOString() });
    const want = (p.tabs ? String(p.tabs).split(',') : PUBLIC_TABS).map(s => s.trim()).filter(t => PUBLIC_TABS.indexOf(t) >= 0);
    const data = {};
    want.forEach(t => {
      let rows = readTab(t);
      if (p.date && (t === 'Check-ins' || t === 'Scores' || t === 'Results')) rows = rows.filter(r => r['Date'] === p.date);
      if (p.since && t === 'Tournaments') rows = rows.filter(r => r['Date'] >= p.since);
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
    fixMembers();
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
  saveMember(b) {
    const m = b.member || {};
    const full = m['Full name'] || [m['First name'], m['Last name']].filter(Boolean).join(' ');
    const rows = readTab('Members');
    const was = String(b.was || '').trim().toLowerCase(); // the row's name before a rename on the site
    const hit = (was && rows.find(r => String(r['Full name'] || '').trim().toLowerCase() === was)) || rows.find(r => m['Member #'] && String(r['Member #']) === String(m['Member #'])) || rows.find(r => r['Full name'] && r['Full name'].toLowerCase() === full.toLowerCase());
    m['Full name'] = full;
    if (hit) upsert('Members', 'Full name', hit['Full name'], m); else appendRows('Members', [m]);
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
// One-time tidy of the Members tab: header "Nickname / other names" becomes "Nicknames", and each cell keeps only
// names that aren't already the bowler's first or last name ("Steve Nolan" for Steven Nolan -> "Steve").
function nickClean(first, last, str) {
  const k = w => w.toLowerCase().replace(/[^a-z0-9']/g, ''), own = new Set((first + ' ' + last).split(/\s+/).map(k).filter(Boolean).concat(['jr', 'sr', 'ii', 'iii'])), seen = new Set();
  return String(str || '').split(/[,;()\/]+/).map(p => p.split(/\s+/).filter(w => k(w) && !own.has(k(w))).join(' ').trim())
    .filter(p => p && !seen.has(p.toLowerCase()) && seen.add(p.toLowerCase())).join(', ');
}
function fixMembers() {
  const s = SpreadsheetApp.getActive().getSheetByName('Members'); if (!s) return;
  const h = headers(s), c = h.indexOf('Nickname / other names'); if (c < 0) return;
  const n = s.getLastRow() - 1, fi = h.indexOf('First name'), li = h.indexOf('Last name');
  if (n > 0) { const v = s.getRange(2, 1, n, h.length).getValues(); s.getRange(2, c + 1, n, 1).setValues(v.map(r => [nickClean(r[fi], r[li], r[c])])); }
  s.getRange(1, c + 1).setValue('Nicknames');
}
function sheet(name) { const s = SpreadsheetApp.getActive().getSheetByName(name); if (!s) throw new Error('No tab named ' + name); return s; }
function headers(s) { return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String); }

// Dates come out as yyyy-MM-dd; everything else as the cell shows it (numbers stay numbers).
function cellOut(v) {
  if (v instanceof Date) return v.getFullYear() < 1901 ? Utilities.formatDate(v, tz(), 'h:mm a') : Utilities.formatDate(v, tz(), 'yyyy-MM-dd'); // a time-only cell (squad time) has no real date
  return v;
}
function readTab(name) {
  const s = sheet(name), h = headers(s), n = s.getLastRow() - 1;
  if (n < 1) return [];
  const vals = s.getRange(2, 1, n, h.length).getValues();
  return vals.filter(r => r.some(v => v !== '' && v !== null)).map(r => {
    const o = {};
    h.forEach((k, i) => { if (k && PRIVATE_COLS.indexOf(k) < 0) o[k] = cellOut(r[i]); });
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
    if (typeof v === 'string' && /^\d{4}-\d\d-\d\d$/.test(v) && /Date|date|sent|For tournament|As of|Last used/.test(k)) { const [y, m, d] = v.split('-').map(Number); return new Date(y, m - 1, d); }
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
  return s.getRange(2, 1, n, h.length).getValues().map(r => { const o = {}; h.forEach((k, i) => o[k] = cellOut(r[i])); return o; });
}
function rowsFor(rows, date) { return (rows || []).map(r => Object.assign({}, r, { 'Date': date })); }
