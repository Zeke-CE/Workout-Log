/**
 * Workout Log database: a Google Apps Script web app that stores every profile's data in the Google Sheet
 * this script is bound to. The page (index.html, served by GitHub Pages) talks to it with POST requests.
 * Setup steps are in README.md. Run `setup` once from the editor before deploying.
 *
 * Tabs keep the column names of the original Workout Log sheet, plus a `profile` column so several people
 * can share one sheet. You can read, sort, chart and hand-edit the tabs like any other sheet.
 */

// cols = header row. text = [first column, how many] runs that hold text or dates and are kept as plain text.
const TABS = {
  Profiles: { cols: ['profile', 'salt', 'passcode_hash', 'created'], text: [[1, 4]] },
  Exercises: { cols: ['Exercise', 'Primary', 'Secondary', 'Equipment', 'Mechanic', 'added_by', 'Primary_muscles', 'Secondary_muscles'], text: [[1, 8]] },
  Workouts: { cols: ['profile', 'workout', 'position', 'exercise', 'sets', 'techniques'], text: [[1, 2], [4, 1], [6, 1]] },
  // A row with `technique` filled in is an intensity-technique sub-set of that set; `sub` is its place under the set (2.1 = first one nested under the second).
  Lifts: { cols: ['profile', 'date', 'workout', 'exercise', 'set', 'reps', 'weight_lb', 'notes', 'sub', 'technique', 'rest_s', 'rom'], text: [[1, 4], [8, 3]] },
  Body_Weight: { cols: ['profile', 'date', 'body_weight_lb'], text: [[1, 2]] },
  Config: { cols: ['key', 'value'], text: [[1, 2]] },
};
const MUSCLES = ['abdominals', 'abductors', 'adductors', 'biceps', 'calves', 'chest', 'forearms', 'glutes', 'hamstrings', 'lats', 'lower back', 'middle back', 'neck', 'quadriceps', 'shoulders', 'traps', 'triceps'];
// The exact muscles an added exercise can list: the keys of XWALK in index.html. Primary_muscles/Secondary_muscles hold these; Primary/Secondary hold their groups.
const FINE = ['abductors', 'adductors', 'neck', 'anterior_deltoid', 'biceps_brachii', 'brachialis', 'brachioradialis', 'erector_spinae', 'forearm_extensors', 'forearm_flexors', 'forearms',
  'gastrocnemius', 'gluteus_maximus', 'gluteus_medius', 'hamstrings', 'hip_flexors', 'infraspinatus_teres_minor', 'lateral_deltoid', 'latissimus_dorsi', 'obliques', 'pectoralis_major',
  'posterior_deltoid', 'quadratus_lumborum', 'quadriceps', 'rectus_abdominis', 'rhomboids', 'serratus_anterior', 'soleus', 'supraspinatus', 'tibialis_anterior', 'transverse_abdominis',
  'trapezius', 'triceps_brachii'];
const EQUIP = ['bands', 'barbell', 'body only', 'cable', 'dumbbell', 'e-z curl bar', 'exercise ball', 'foam roll', 'kettlebells', 'machine', 'medicine ball', 'none listed', 'other'];
const MECHANICS = ['', 'compound', 'isolation'];
const TECHNIQUES = ['partial', 'myo', 'rest-pause', 'drop'], SUB = /^[1-9](\.[1-9]){0,2}$/;
const FORMULA_START = /^[=+\-@]/; // Sheets would run these as formulas, so they are stored behind an apostrophe
const MAX = { ops: 200, rows: 20000, name: 80, notes: 500, workouts: 60, items: 40, sets: 10 };

// ---------- Web app entry points ----------

function doGet() {
  return reply({ ok: true, app: 'Workout Log database' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let out;
  try {
    lock.waitLock(25000); // one request at a time, so two phones cannot interleave writes
    const req = JSON.parse(e.postData.contents);
    const handler = { signup: signup, login: login, sync: sync }[req.action];
    if (!handler) throw fail('Unknown request.');
    out = handler(req);
    out.ok = true;
  } catch (err) {
    out = err.userMessage
      ? { error: err.userMessage, code: err.code }
      : { error: 'The database hit a problem (' + (err && err.message ? err.message : err) + '). Nothing was lost; it will retry.', code: 'server' };
  } finally {
    try { lock.releaseLock(); } catch (err) { /* the lock was never taken */ }
  }
  return reply(out);
}

function reply(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** An error whose message is meant for the person using the page. code: 'invalid' | 'auth' | 'server'. */
function fail(message, code) {
  const e = new Error(message);
  e.userMessage = message;
  e.code = code || 'invalid';
  return e;
}

// ---------- Run once from the editor ----------

/** Creates the tabs, a join code for new profiles and the signing secret. Safe to run again. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(name => {
    const def = TABS[name], sh = ss.getSheetByName(name) || ss.insertSheet(name), short = def.cols.length - sh.getMaxColumns();
    if (short > 0) sh.insertColumnsAfter(sh.getMaxColumns(), short);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]);
    sh.setFrozenRows(1);
    def.text.forEach(([col, n]) => sh.getRange(2, col, sh.getMaxRows() - 1, n).setNumberFormat('@'));
  });
  if (!config('join_code')) {
    const code = Utilities.getUuid().replace(/-/g, '').slice(0, 8).replace(/[0-9]/g, d => 'ghjkmnpqrs'[d]);
    appendRows('Config', [['join_code', code]]);
  }
  secret();
}

// ---------- Sheet access ----------

function tab(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw fail('The database is not set up yet: its owner needs to run "setup" in the Apps Script editor.', 'server');
  return sh;
}

function readRows(name) {
  const sh = tab(name), last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, TABS[name].cols.length).getValues();
}

function appendRows(name, rows) {
  if (!rows.length) return;
  const sh = tab(name), def = TABS[name], start = sh.getLastRow() + 1, end = start + rows.length - 1, max = sh.getMaxRows();
  if (end >= max) sh.insertRowsAfter(max, end - max + 200); // always leaves blank rows, so deleting data rows never empties the grid
  def.text.forEach(([col, n]) => sh.getRange(start, col, rows.length, n).setNumberFormat('@'));
  sh.getRange(start, 1, rows.length, def.cols.length)
    .setValues(rows.map(r => r.map(v => (typeof v === 'string' && FORMULA_START.test(v) ? "'" + v : v))));
}

/** Deletes every data row the test matches, bottom-up in contiguous blocks. Returns how many. */
function removeRows(name, match) {
  const sh = tab(name), rows = readRows(name);
  let removed = 0;
  for (let i = rows.length - 1; i >= 0; i--) {
    if (!match(rows[i])) continue;
    let j = i;
    while (j > 0 && match(rows[j - 1])) j--;
    sh.deleteRows(j + 2, i - j + 1);
    removed += i - j + 1;
    i = j;
  }
  return removed;
}

let zone = null;
/** A cell as trimmed text. Date cells (typed by hand in the sheet) become yyyy-mm-dd. */
function str(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    zone = zone || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(v, zone, 'yyyy-MM-dd');
  }
  const s = String(v === null || v === undefined ? '' : v).trim();
  return /^'[=+\-@]/.test(s) ? s.slice(1) : s;
}
const key = v => str(v).toLowerCase();
function num(v) {
  if (typeof v === 'number') return v;
  const s = str(v);
  return s === '' ? NaN : Number(s);
}
function isDate(s) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const d = new Date(s + 'T00:00:00Z');
  return !isNaN(d) && d.toISOString().slice(0, 10) === s;
}
const whole = (n, lo, hi) => typeof n === 'number' && Number.isInteger(n) && n >= lo && n <= hi;
const list = v => str(v).toLowerCase().split(',').map(s => s.trim()).filter(Boolean);

function config(k) {
  const row = readRows('Config').find(r => str(r[0]) === k);
  return row ? str(row[1]) : '';
}

// ---------- Profiles and sign-in ----------

function secret() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SECRET', s);
  }
  return s;
}

function digest(text) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8));
}
/** Salted and stretched, so the sheet never holds a passcode. */
function passHash(salt, passcode) {
  let h = passcode;
  for (let i = 0; i < 200; i++) h = digest(salt + h);
  return 'h' + h;
}

function findProfile(name) {
  const k = key(name);
  const row = k && readRows('Profiles').find(r => key(r[0]) === k);
  return row ? { name: str(row[0]), salt: str(row[1]), hash: str(row[2]) } : null;
}

/** The token is signed with the script's secret and the passcode hash, so it cannot be forged or outlive a reset. */
function tokenFor(p) {
  const id = p.name.toLowerCase();
  return id + '.' + Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(id + '.' + p.hash, secret()));
}
function whoIs(token) {
  const t = String(token || ''), dot = t.indexOf('.');
  const p = dot > 0 ? findProfile(t.slice(0, dot)) : null;
  if (!p || tokenFor(p) !== t) throw fail('Please sign in again.', 'auth');
  return p;
}

/** Stops guessing: after `max` misses on one counter, refuses for 10 minutes. Returns a function that records a miss. */
function throttle(counter, max) {
  const cache = CacheService.getScriptCache(), k = counter.slice(0, 100), n = Number(cache.get(k)) || 0;
  if (n >= max) throw fail('Too many wrong tries. Wait 10 minutes and try again.');
  return () => cache.put(k, String(n + 1), 600);
}

function signup(req) {
  const name = String(req.profile || '').trim().replace(/\s+/g, ' '), pass = String(req.passcode || '');
  if (!/^[A-Za-z0-9][A-Za-z0-9 _-]{0,23}$/.test(name)) {
    throw fail('Profile names are 1 to 24 letters, numbers, spaces, - or _, and start with a letter or number.');
  }
  if (pass.length < 4 || pass.length > 100) throw fail('The passcode needs at least 4 characters.');
  const code = config('join_code'), miss = throttle('join', 20);
  if (!code) throw fail('New profiles are turned off for this log.');
  if (String(req.join || '').trim().toLowerCase() !== code.toLowerCase()) {
    miss();
    throw fail('That join code is not right. Ask the person who runs this log for it.');
  }
  if (findProfile(name)) throw fail('The profile name "' + name + '" is taken. Sign in instead, or pick another name.');
  const salt = 's' + Utilities.getUuid(), p = { name: name, salt: salt, hash: passHash(salt, pass) };
  appendRows('Profiles', [[name, salt, p.hash, new Date().toISOString()]]);
  return { profile: name, token: tokenFor(p) };
}

function login(req) {
  const p = findProfile(req.profile), miss = throttle('fail:' + key(req.profile), 8);
  if (!p || passHash(p.salt, String(req.passcode || '')) !== p.hash) {
    miss();
    throw fail('That profile name or passcode is not right.');
  }
  return { profile: p.name, token: tokenFor(p) };
}

// ---------- Rows: checking what comes in, tidying what is read ----------

function label(v, what, max, required) {
  if (typeof v !== 'string') throw fail(what + ' is missing.');
  const s = v.trim();
  if (required && !s) throw fail(what + ' is missing.');
  if (s.length > max) throw fail(what + ' is longer than ' + max + ' characters.');
  return s;
}

/** Client lift row: [date, workout, exercise, set, reps, weight, notes], and for a technique sub-set also [sub, technique, rest_s, rom]. */
function cleanLift(r) {
  if (!Array.isArray(r) || (r.length !== 7 && r.length !== 11)) throw fail('A set row does not have 7 or 11 values.');
  if (!isDate(r[0])) throw fail('"' + r[0] + '" is not a yyyy-mm-dd date.');
  if (!whole(r[3], 1, 99)) throw fail('Set "' + r[3] + '" is not a whole number, 1 or more.');
  if (!whole(r[4], 1, 9999)) throw fail('Reps "' + r[4] + '" is not a whole number, 1 or more.');
  if (typeof r[5] !== 'number' || !(r[5] >= 0) || r[5] > 99999) throw fail('Weight "' + r[5] + '" is not a number, 0 or more.');
  const row = [r[0], label(r[1], 'Workout name', MAX.name), label(r[2], 'Exercise', MAX.name, true), r[3], r[4], r[5], label(r[6], 'Notes', MAX.notes)];
  if (r.length === 7) return row;
  const blankOr = (v, lo, hi) => v === '' || (typeof v === 'number' && v >= lo && v <= hi);
  if (TECHNIQUES.indexOf(r[8]) < 0 || typeof r[7] !== 'string' || !SUB.test(r[7])) throw fail('Technique "' + r[8] + '" at "' + r[7] + '" is not a sub-set this log knows.');
  if (!blankOr(r[9], 0, 3600)) throw fail('Rest "' + r[9] + '" is not a number of seconds.');
  if (!blankOr(r[10], 0.01, 1)) throw fail('Range of motion "' + r[10] + '" is not a share of the range, up to 1.');
  return row.concat([r[7], r[8], r[9], r[10]]);
}
/** A saved workout's planned sub-sets for an exercise's last set: [{ t, kids }], up to 3 levels deep and 9 per level. */
function cleanTech(list, depth) {
  if (!Array.isArray(list) || list.length > 9 || (depth > 3 && list.length)) throw fail('A planned technique is not valid.');
  return list.map(n => {
    if (!n || TECHNIQUES.indexOf(n.t) < 0) throw fail('A planned technique is not one this log knows.');
    return { t: n.t, kids: cleanTech(n.kids || [], depth + 1) };
  });
}
function cleanBw(r) {
  if (!Array.isArray(r) || r.length !== 2) throw fail('A body-weight row does not have 2 values.');
  if (!isDate(r[0])) throw fail('"' + r[0] + '" is not a yyyy-mm-dd date.');
  if (typeof r[1] !== 'number' || !(r[1] > 0) || r[1] > 9999) throw fail('Body weight "' + r[1] + '" is not a positive number.');
  return [r[0], r[1]];
}
function cleanRows(rows, clean) {
  if (!Array.isArray(rows) || rows.length > MAX.rows) throw fail('Too many rows in one change.');
  return rows.map(clean);
}
function cleanWorkouts(ws) {
  if (!Array.isArray(ws) || ws.length > MAX.workouts) throw fail('Too many workouts.');
  const seen = {};
  return ws.map(w => {
    const name = label(w && w.name, 'Workout name', MAX.name, true);
    if (seen[name.toLowerCase()]) throw fail('Two workouts are named "' + name + '".');
    seen[name.toLowerCase()] = true;
    if (!Array.isArray(w.items) || !w.items.length || w.items.length > MAX.items) throw fail('"' + name + '" needs 1 to ' + MAX.items + ' exercises.');
    return { name: name, items: w.items.map(it => {
      if (!whole(it && it.sets, 1, MAX.sets)) throw fail('"' + name + '" has an exercise without 1 to ' + MAX.sets + ' sets.');
      return { ex: label(it.ex, 'Exercise', MAX.name, true), sets: it.sets, tech: cleanTech(it.tech || [], 1) };
    }) };
  });
}
function cleanExercise(x) {
  const name = label(x && x.name, 'Exercise name', MAX.name, true).replace(/\s+/g, ' ');
  const muscles = (v, what) => {
    if (!Array.isArray(v) || v.some(m => MUSCLES.indexOf(m) < 0)) throw fail(what + ' muscles for "' + name + '" are not from the muscle list.');
    return v.filter((m, i) => v.indexOf(m) === i);
  };
  const primary = muscles(x.primary, 'Primary');
  if (!primary.length) throw fail('"' + name + '" needs at least one primary muscle.');
  if (EQUIP.indexOf(x.equipment) < 0) throw fail('Equipment for "' + name + '" is not from the equipment list.');
  if (MECHANICS.indexOf(x.mechanic) < 0) throw fail('Mechanic for "' + name + '" must be compound, isolation or blank.');
  const out = { name: name, primary: primary, secondary: muscles(x.secondary, 'Secondary').filter(m => primary.indexOf(m) < 0), equipment: x.equipment, mechanic: x.mechanic };
  if (x.fine == null) return out; // an older client, or an import without exact muscles
  const f = x.fine, ok = Array.isArray(f) && f.length === 2 && f.every(v => Array.isArray(v) && v.length <= FINE.length && v.every(m => FINE.indexOf(m) >= 0));
  if (!ok || !f[0].length) throw fail('The exact muscles for "' + name + '" must be from the muscle list, with at least one primary.');
  const pf = f[0].filter((m, i) => f[0].indexOf(m) === i);
  out.fine = [pf, f[1].filter((m, i) => f[1].indexOf(m) === i && pf.indexOf(m) < 0)];
  return out;
}

/** A sheet row as the client's row, or null when it is not usable (blank line, half-typed row). */
function liftFromSheet(r) {
  try {
    const row = [str(r[1]), str(r[2]), str(r[3]), num(r[4]), num(r[5]), num(r[6]), str(r[7])], opt = v => (str(v) === '' ? '' : num(v));
    return cleanLift(key(r[9]) ? row.concat([str(r[8]), key(r[9]), opt(r[10]), opt(r[11])]) : row);
  } catch (err) {
    return null;
  }
}
function techFromSheet(v) {
  try {
    return cleanTech(JSON.parse(str(v) || '[]'), 1);
  } catch (err) {
    return [];
  }
}
function bwFromSheet(r) {
  try {
    return cleanBw([str(r[1]), num(r[2])]);
  } catch (err) {
    return null;
  }
}
function exerciseFromSheet(r, me) {
  const primary = list(r[1]).filter(m => MUSCLES.indexOf(m) >= 0), equipment = key(r[3]), mechanic = key(r[4]);
  if (!str(r[0]) || !primary.length) return null;
  const pf = list(r[6]).filter(m => FINE.indexOf(m) >= 0), sf = list(r[7]).filter(m => FINE.indexOf(m) >= 0 && pf.indexOf(m) < 0);
  return Object.assign(pf.length ? { fine: [pf, sf] } : {}, {
    name: str(r[0]), primary: primary,
    secondary: list(r[2]).filter(m => MUSCLES.indexOf(m) >= 0 && primary.indexOf(m) < 0),
    equipment: EQUIP.indexOf(equipment) >= 0 ? equipment : 'none listed', // a blank cell means "none listed", as in the original sheet
    mechanic: MECHANICS.indexOf(mechanic) >= 0 ? mechanic : '',
    mine: key(r[5]) === me,
  });
}
const byDate = (a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0); // sort is stable, so set order within a day is kept

// ---------- Sync: apply this profile's queued changes, then send back its whole state ----------

function sync(req) {
  const profile = whoIs(req.token), me = profile.name.toLowerCase(), mine = r => key(r[0]) === me;
  const ops = Array.isArray(req.ops) ? req.ops : [];
  if (ops.length > MAX.ops) throw fail('Too many changes in one request.');

  let lifts = readRows('Lifts').filter(mine).map(liftFromSheet).filter(Boolean);
  let bw = readRows('Body_Weight').filter(mine).map(bwFromSheet).filter(Boolean);
  const seen = { lifts: new Set(lifts.map(r => JSON.stringify(r))), bw: new Set(bw.map(r => JSON.stringify(r))) };

  /** Appends the rows that are not already there, so a retried request never doubles anything. */
  const addRows = (kind, tabName, have, rows) => {
    const fresh = rows.filter(r => {
      const k = JSON.stringify(r);
      if (seen[kind].has(k)) return false;
      seen[kind].add(k);
      return true;
    });
    appendRows(tabName, fresh.map(r => [profile.name].concat(r, '', '', '', '').slice(0, TABS[tabName].cols.length))); // a straight set leaves the technique columns blank
    fresh.forEach(r => have.push(r));
  };
  const ownerOf = name => {
    const row = readRows('Exercises').find(r => key(r[0]) === name.toLowerCase());
    return row ? key(row[5]) : null;
  };
  const apply = {
    lifts: op => addRows('lifts', 'Lifts', lifts, cleanRows(op.rows, cleanLift)),
    bw: op => addRows('bw', 'Body_Weight', bw, cleanRows(op.rows, cleanBw)),
    workouts: op => {
      const ws = cleanWorkouts(op.workouts);
      removeRows('Workouts', mine);
      appendRows('Workouts', [].concat.apply([], ws.map(w => w.items.map((it, i) => [profile.name, w.name, i + 1, it.ex, it.sets, it.tech.length ? JSON.stringify(it.tech) : '']))));
    },
    exercise: op => {
      const x = cleanExercise(op.ex), owner = ownerOf(x.name);
      if (owner !== null && owner !== me) throw fail('"' + x.name + '" was added by someone else, so it cannot be changed from this profile.');
      removeRows('Exercises', r => key(r[0]) === x.name.toLowerCase());
      appendRows('Exercises', [[x.name, x.primary.join(', '), x.secondary.join(', '), x.equipment === 'none listed' ? '' : x.equipment, x.mechanic, profile.name,
        x.fine ? x.fine[0].join(', ') : '', x.fine ? x.fine[1].join(', ') : '']]);
    },
    delExercise: op => {
      const name = label(op.name, 'Exercise name', MAX.name, true), owner = ownerOf(name);
      if (owner !== null && owner !== me) throw fail('"' + name + '" was added by someone else, so it cannot be removed from this profile.');
      removeRows('Exercises', r => key(r[0]) === name.toLowerCase());
    },
    delDay: op => {
      if (!isDate(op.date)) throw fail('"' + op.date + '" is not a yyyy-mm-dd date.');
      removeRows('Lifts', r => mine(r) && str(r[1]) === op.date);
      removeRows('Body_Weight', r => mine(r) && str(r[1]) === op.date);
      lifts = lifts.filter(r => r[0] !== op.date);
      bw = bw.filter(r => r[0] !== op.date);
      seen.lifts = new Set(lifts.map(r => JSON.stringify(r)));
      seen.bw = new Set(bw.map(r => JSON.stringify(r)));
    },
  };

  // A change that fails its checks is reported and skipped, so it cannot block the changes queued behind it.
  const rejected = [];
  ops.forEach((op, i) => {
    try {
      if (!op || !apply.hasOwnProperty(op.t)) throw fail('Unknown kind of change.');
      apply[op.t](op);
    } catch (err) {
      if (!err.userMessage) throw err;
      rejected.push({ i: i, why: err.userMessage });
    }
  });

  const workouts = [];
  readRows('Workouts').filter(mine)
    .map(r => ({ name: str(r[1]), pos: num(r[2]), ex: str(r[3]), sets: num(r[4]), tech: techFromSheet(r[5]) }))
    .filter(r => r.name && r.ex && whole(r.sets, 1, MAX.sets))
    .forEach(r => {
      let w = workouts.filter(x => x.name.toLowerCase() === r.name.toLowerCase())[0];
      if (!w) workouts.push(w = { name: r.name, items: [] });
      w.items.push(r);
    });
  workouts.forEach(w => {
    w.items = w.items.sort((a, b) => (a.pos || 0) - (b.pos || 0)).map(r => ({ ex: r.ex, sets: r.sets, tech: r.tech }));
  });

  return {
    profile: profile.name,
    exercises: readRows('Exercises').map(r => exerciseFromSheet(r, me)).filter(Boolean),
    workouts: workouts,
    lifts: lifts.sort(byDate),
    bw: bw.sort(byDate),
    rejected: rejected,
  };
}
