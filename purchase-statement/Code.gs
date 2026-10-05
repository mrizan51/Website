/**
 * Premier Exports International — Purchase Statement, live on Google Sheets.
 *
 * SET-UP — once, by the owner of the sheet
 *   1. Open the purchase statement in Google Sheets.
 *   2. Extensions ▸ Apps Script. Delete what is there, paste this file, click Save.
 *   3. Reload the sheet. A "PEI Purchase" menu appears.
 *   4. PEI Purchase ▸ Set up statement… and allow access when Google asks.
 *
 * WHAT IT DOES
 *   Party names   Every known spelling becomes one name (PARTY_NAMES below), and the
 *                 Party column then takes only names on the Parties sheet.
 *   Lot No.       Always "PUR-" and a number. A new row gets the next number by itself.
 *   Formatting    One font, one set of number formats, widths and colours. Employees
 *                 type in the entry cells; any formatting change is put back at once.
 *   Amount/Swipe  Swipe fills red when Amount is greater; Amount fills green when Swipe
 *                 is greater.
 *   Protection    Title, headings, the Month column and everything outside the entry
 *                 cells are owner-only.
 *   Cleanup log   Every change set-up makes, and every entry worth a second look.
 *
 * Re-running set-up is safe: it changes only what is not already right, and adds
 * to the log rather than replacing it. File ▸ Version history undoes a run.
 */

const SHEET = 'Purchase';
const PARTIES = 'Parties';
const LOG = 'Cleanup log';
const TAG = 'PEI purchase statement';     // marks the protections this script owns
const TITLE = 'PREMIER EXPORTS INTERNATIONAL  ·  PURCHASE STATEMENT';
const HEAD_ROW = 2;
const FIRST = 3;          // first purchase row
const ROWS = 5000;        // purchase rows kept formatted and open for entry
const HEADROOM = 200;     // when fewer empty rows than this remain, 1,000 more are added
const COL = { month: 1, lot: 2, date: 3, party: 4, weight: 5, amount: 6, swipe: 7, payment: 8, notes: 9 };
const LAST_COL = 9;
const LOT_PREFIX = 'PUR-';
const LOT_DIGITS = 3;

const STYLE = {
  font: 'Arial', size: 10,
  ink: '#0b2c3a', muted: '#5e7c8b', title: '#065e7a',
  head: '#065e7a', headInk: '#ffffff', rule: '#d5e3ea',
  red: { fill: '#f4c7c3', ink: '#a50e0e' },
  green: { fill: '#b7e1cd', ink: '#0d652d' },
  amber: '#fce8b2',
};

// Columns A to I: heading, width in pixels, alignment, number format.
const COLUMNS = [
  ['Month', 70, 'center', null],
  ['Lot No.', 95, 'center', '@'],
  ['Date', 95, 'center', 'dd/mm/yyyy'],
  ['Party', 230, 'left', '@'],
  ['Weight (kg)', 105, 'right', '#,##0.00'],
  ['Amount (₹)', 125, 'right', '[$₹]#,##0.00'],
  ['Swipe (₹)', 125, 'right', '[$₹]#,##0.00'],
  ['Payment', 100, 'center', 'dd/mm/yyyy'],
  ['Notes', 260, 'left', '@'],
];

// Spelling as typed → the one name kept. Names already right are not listed;
// any capitalisation of a name below, or of a kept name, is matched too.
const PARTY_NAMES = {
  "ABS FISHERIES": "ABS Fisheries",
  "AFSAL ARF": "Afsal ARF",
  "AFZAL ARF": "Afsal ARF",
  "ANI MARINE": "ANI Marine Foods",
  "ANI MARINE FOODS": "ANI Marine Foods",
  "ANSAR": "Ansar",
  "ARF FISHERIES": "ARF Fisheries Asim",
  "ARF FISHERIES ASIM": "ARF Fisheries Asim",
  "ARSHA MARINE": "Arsha Marine",
  "ASR FISHERIES": "Suresh N ASR",
  "ASR GROUP": "ASR Group",
  "BAHAN": "Bahan",
  "CRH": "Hari CRH",
  "CRH HARI": "Hari CRH",
  "DSS": "DSS Marine",
  "EFADH": "Arfa Efadth",
  "FRIENDS FISHERIES": "Friends Fisheries",
  "HARBOUR ATLANTIC EKS": "Harbour Atlantic EKS",
  "IBRAHIM PP": "Ibrahim PP",
  "KAIRALI SEAFOOD": "Kairali Seafood",
  "KCM": "Abdulla KCM",
  "KCM ABDULLA": "Abdulla KCM",
  "KKH": "KKH Fisheries",
  "KKH FISHERIES": "KKH Fisheries",
  "MRP FISHERIES": "MRP Fisheries",
  "NOUSHAD K A": "Noushad K A AHD",
  "NOUSHAD KA": "Noushad K A AHD",
  "NOUSHYAD K A": "Noushad K A AHD",
  "NPS MARINE": "NPS Marine",
  "PVM MARINE": "PVM Marine",
  "PVM Marines": "PVM Marine",
  "RAJENDRA PRASAD": "Rajendra Prasad Shashthri",
  "RISVAN": "Riswan P N",
  "RISVAN P": "Riswan P N",
  "RISVAN P N": "Riswan P N",
  "SH COMMN SHAHINSHA": "Shahinsha SH Commn",
  "SHAHINSHA": "Shahinsha SH Commn",
  "SHAHINSHA SH COMMN": "Shahinsha SH Commn",
  "SHAMNAD": "Shamnad YSF",
  "SHYAM SUNDER": "Shyam Sunder",
  "STEPHEN": "Stephen Andrews",
  "TRHAMPI": "Thambi",
  "VIJEESH M": "Vijeesh M",
  "YOOSUF": "Yousef Chandiroor",
  "YOUSEF": "Yousef Chandiroor",
  "YOUSEF CHANDIROOR": "Yousef Chandiroor"
};

// Merged on evidence rather than spelling: from mid-August names were typed short and
// in capitals, and each short form begins the week its long form stops, at the same
// price. The log asks for these to be confirmed.
const CONFIRM = ["ARF FISHERIES", "ASR FISHERIES", "CRH", "DSS", "EFADH", "KCM", "KKH",
                 "NOUSHAD K A", "RAJENDRA PRASAD", "SHAMNAD", "STEPHEN", "TRHAMPI",
                 "YOOSUF", "YOUSEF"];

const LOG_HEAD = ['When', 'Row', 'Lot No.', 'Field', 'Was', 'Now', 'Note'];

// ===================================================================== menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('PEI Purchase')
    .addItem('Put formatting back now', 'restoreAll')
    .addSeparator()
    .addItem('Set up statement… (owner, once)', 'setup')
    .addToUi();
}

// ==================================================================== set-up
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET);
  if (!sh) throw new Error('No sheet named "' + SHEET + '".');
  ensureRows_(sh);
  const log = cleanData_(ss, sh);
  writeParties_(ss, sh);
  writeLog_(ss, log);
  restoreAll();
  protect_(ss, sh);
  installTriggers_(ss);
  ss.toast(log.length + ' lines written to "' + LOG + '". Formatting, rules and protection are in place.',
           'Purchase statement set up', 10);
}

/** Puts every format, rule and check back. Runs on any formatting change, and nightly. */
function restoreAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET);
  if (!sh) return;
  withLock_(20000, function () {
    if (ensureRows_(sh)) openEntryRange_(sh);
    ensureMonth_(sh);
    formatSheet_(sh);
    applyRules_(sh);
    applyValidation_(ss, sh);
    const ps = ss.getSheetByName(PARTIES);
    if (ps) formatParties_(ps);
    const ls = ss.getSheetByName(LOG);
    if (ls) formatLog_(ls);
  });
}

// ================================================================= triggers
/** Simple trigger: runs as whoever typed. Numbers new rows and tidies what was pasted. */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET) return;
  const top = Math.max(e.range.getRow(), FIRST);
  const bottom = e.range.getRow() + e.range.getNumRows() - 1;
  if (bottom < top) return;
  try { fillLots_(sh, top, bottom); } catch (err) { console.warn('lot numbers: ' + err); }
  try { formatRows_(sh, top, bottom); } catch (err) { console.warn('formatting: ' + err); }
}

/** Installed trigger: runs as the owner, so it can reach protected cells too. */
function onChangeInstalled(e) {
  const t = e && e.changeType;
  const watch = ['FORMAT', 'OTHER', 'INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN',
                 'INSERT_GRID', 'REMOVE_GRID'];
  if (watch.indexOf(t) >= 0) restoreAll();
}

function nightlyRestore() {
  restoreAll();
}

function installTriggers_(ss) {
  const ours = { onChangeInstalled: true, nightlyRestore: true };
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (ours[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onChangeInstalled').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('nightlyRestore').timeBased().everyDays(1).atHour(2).create();
}

// ================================================================ cleaning
function partyKey_(s) {
  return String(s).trim().replace(/\s+/g, ' ').toUpperCase();
}

let PARTY_INDEX_ = null;
function canonicalParty_(s) {
  if (!PARTY_INDEX_) {
    PARTY_INDEX_ = {};
    Object.keys(PARTY_NAMES).forEach(function (k) {
      PARTY_INDEX_[partyKey_(k)] = PARTY_NAMES[k];
      PARTY_INDEX_[partyKey_(PARTY_NAMES[k])] = PARTY_NAMES[k];
    });
  }
  const t = String(s).trim().replace(/\s+/g, ' ');
  return PARTY_INDEX_[partyKey_(t)] || t;
}

function formatLot_(n) {
  let s = String(Math.round(n));
  while (s.length < LOT_DIGITS) s = '0' + s;
  return LOT_PREFIX + s;
}

function lotNumber_(v) {
  if (typeof v === 'number' && isFinite(v)) return Math.round(v);
  const m = String(v).trim().match(/^PUR-?\s*0*(\d+)$/i);
  return m ? Number(m[1]) : null;
}

function isBlank_(v) {
  return v === '' || v === null || v === undefined;
}

function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

function money_(n) {                    // 1234567.5 → "1,234,567.50", the sheet's own #,##0.00
  const neg = n < 0;
  const p = Math.abs(n).toFixed(2).split('.');
  return (neg ? '-' : '') + p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
}

/**
 * Cleans the purchase rows in place and returns the log lines.
 * Order per row: notes, weight, party, lot number; then the checks.
 */
function cleanData_(ss, sh) {
  const tz = ss.getSpreadsheetTimeZone();
  const day = function (d) { return Utilities.formatDate(d, tz, 'dd/MM/yyyy'); };
  const now = new Date();
  const last = lastDataRow_(sh);
  const log = [];
  if (last < FIRST) return log;
  const n = last - FIRST + 1;
  const rng = sh.getRange(FIRST, COL.lot, n, LAST_COL - COL.lot + 1);   // B:I
  const v = rng.getValues();
  const B = 0, C = 1, D = 2, E = 3, F = 4, G = 5, H = 6, I = 7;          // offsets within B:I
  const confirm = {};
  CONFIRM.forEach(function (k) { confirm[partyKey_(k)] = true; });
  let numbered = 0;

  v.forEach(function (row, i) {
    const r = FIRST + i;
    // notes: spaces trimmed
    if (typeof row[I] === 'string' && row[I] !== row[I].trim()) {
      const t = row[I].trim();
      log.push([now, r, '', 'Notes', row[I], t, t ? 'Spaces trimmed' : 'Only spaces — cleared']);
      row[I] = t;
    }
    // weight given in boxes: kept in Notes, the kg cell left for a real weight
    if (typeof row[E] === 'string' && row[E].trim() !== '') {
      const m = row[E].match(/^\s*(\d+(?:\.\d+)?)\s*box(?:es)?\s*$/i);
      if (m) {
        const note = m[1] + ' boxes';
        log.push([now, r, '', 'Weight (kg)', row[E], '',
                  'Boxes, not kg — moved to Notes as "' + note + '"']);
        row[I] = row[I] ? note + '; ' + row[I] : note;
        row[E] = '';
      } else if (/^\s*\d+(?:\.\d+)?\s*$/.test(row[E])) {
        log.push([now, r, '', 'Weight (kg)', row[E], Number(row[E]), 'Text made a number']);
        row[E] = Number(row[E]);
      }
    }
    // party: one name per supplier
    if (typeof row[D] === 'string' && row[D].trim() !== '') {
      const c = canonicalParty_(row[D]);
      if (c !== row[D]) {
        log.push([now, r, '', 'Party', row[D], c,
                  confirm[partyKey_(row[D])]
                    ? 'Same supplier by timing and price, not spelling — please confirm'
                    : 'Same supplier, another spelling']);
        row[D] = c;
      }
    }
    // lot number: PUR- text on rows with a purchase, nothing on empty rows
    const hasData = [C, D, E, F, G, H, I].some(function (k) { return !isBlank_(row[k]); });
    const num = isBlank_(row[B]) ? null : lotNumber_(row[B]);
    if (!hasData) {
      if (!isBlank_(row[B])) {
        log.push([now, r, '', 'Lot No.', num !== null ? formatLot_(num) : String(row[B]), '',
                  'Number on an empty row — removed; the next purchase is numbered automatically']);
        row[B] = '';
      }
    } else if (num !== null && row[B] !== formatLot_(num)) {
      row[B] = formatLot_(num);
      numbered++;
    }
  });

  // checks — nothing is changed, the row is listed for a second look
  const same1 = {}, same2 = {};
  const key1 = function (row) { return [row[C].getTime(), partyKey_(row[D]), row[F]].join('|'); };
  const key2 = function (row) { return [row[C].getTime(), row[E], row[F]].join('|'); };
  v.forEach(function (row, i) {
    if (!isDate_(row[C]) || typeof row[F] !== 'number') return;
    (same1[key1(row)] = same1[key1(row)] || []).push(i);
    if (typeof row[E] === 'number') (same2[key2(row)] = same2[key2(row)] || []).push(i);
  });
  const other = function (group, i) {
    for (let k = 0; k < group.length; k++) if (group[k] !== i) return group[k];
    return null;
  };
  v.forEach(function (row, i) {
    const r = FIRST + i, lot = row[B];
    const check = function (note) { log.push([now, r, lot, 'Check', '', '', note]); };
    if (isDate_(row[H]) && isDate_(row[C]) && row[H] < row[C]) {
      check('Payment date ' + day(row[H]) + ' is before the purchase date ' + day(row[C]));
    }
    if (typeof row[E] === 'number' && row[E] > 0 && typeof row[F] === 'number') {
      const rate = row[F] / row[E];
      if (rate < 20 || rate > 2000) check('Weight looks wrong for the amount: ₹' + money_(rate) + ' per kg');
    }
    if (typeof row[F] === 'number' && typeof row[G] === 'number' && Math.abs(row[F] - row[G]) > 0.005) {
      check('Amount and Swipe differ by ₹' + money_(Math.abs(row[F] - row[G])));
    }
    if (isDate_(row[C]) && typeof row[F] === 'number') {
      const j1 = other(same1[key1(row)], i);
      const j2 = typeof row[E] === 'number' ? other(same2[key2(row)], i) : null;
      if (j1 !== null) {
        check('Same date, party and amount as ' + v[j1][B]);
      } else if (j2 !== null) {
        check('Same date, weight and amount as ' + v[j2][B] + ' (' + v[j2][D] + ')');
      }
    }
    if (typeof row[I] === 'string' && /Rs\.\s*\/-/.test(row[I])) {
      check('Note has no amount: "' + row[I] + '"');
    }
  });

  rng.setValues(v);
  if (numbered) {
    log.push([now, '', '', 'Lot No.', 'numbers', 'PUR- text',
              numbered + ' lot numbers stored as text with the PUR- prefix']);
  }
  if (!sameFormula_(sh.getRange(FIRST, COL.month).getFormula(), monthFormula_())) {
    log.push([now, '', '', 'Month', 'typed labels', 'formula',
              'Month shown on every row, worked out from the Date']);
  }
  return log;
}

// ============================================================ lot numbers
function maxLot_(sh) {
  const last = lastDataRow_(sh);
  if (last < FIRST) return 0;
  return sh.getRange(FIRST, COL.lot, last - FIRST + 1, 1).getValues()
    .reduce(function (m, row) {
      const n = isBlank_(row[0]) ? null : lotNumber_(row[0]);
      return n !== null && n > m ? n : m;
    }, 0);
}

/** Gives the next PUR- number to every edited row that has a date or party but no lot. */
function fillLots_(sh, top, bottom) {
  const n = bottom - top + 1;
  const peek = sh.getRange(top, COL.lot, n, 3).getValues();        // lot, date, party
  const wants = peek.some(function (row) {
    return isBlank_(row[0]) && (!isBlank_(row[1]) || !isBlank_(row[2]));
  });
  if (!wants) return;
  withLock_(10000, function () {                       // two people typing at once get different numbers
    const rows = sh.getRange(top, COL.lot, n, 3).getValues();
    let next = maxLot_(sh) + 1;
    const out = rows.map(function (row) {
      if (isBlank_(row[0]) && (!isBlank_(row[1]) || !isBlank_(row[2]))) return [formatLot_(next++)];
      return [row[0]];
    });
    sh.getRange(top, COL.lot, n, 1).setValues(out);
  });
}

// ============================================================== formatting
function monthFormula_() {
  return '=ARRAYFORMULA(IF(C' + FIRST + ':C="","",UPPER(TEXT(C' + FIRST + ':C,"mmm"))))';
}

function sameFormula_(a, b) {
  const k = function (f) { return String(f).replace(/\s+/g, '').toUpperCase(); };
  return k(a) === k(b);
}

function ensureMonth_(sh) {
  const top = sh.getRange(FIRST, COL.month);
  if (sameFormula_(top.getFormula(), monthFormula_())) return;
  sh.getRange(FIRST, COL.month, sh.getMaxRows() - FIRST + 1, 1).clearContent();
  top.setFormula(monthFormula_());
}

/** Adds rows so the open entry area always has headroom. True if rows were added. */
function ensureRows_(sh) {
  const want = FIRST + ROWS - 1;
  const max = sh.getMaxRows();
  if (max < want) {
    sh.insertRowsAfter(max, want - max);
    return true;
  }
  if (max - lastDataRow_(sh) < HEADROOM) {
    sh.insertRowsAfter(max, 1000);
    return true;
  }
  return false;
}

function lastDataRow_(sh) {
  const max = sh.getMaxRows();
  if (max < FIRST) return FIRST - 1;
  const v = sh.getRange(FIRST, COL.lot, max - FIRST + 1, LAST_COL - COL.lot + 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (v[i].some(function (x) { return !isBlank_(x); })) return FIRST + i;
  }
  return FIRST - 1;
}

function plain_(range) {
  return range.setFontFamily(STYLE.font).setFontSize(STYLE.size).setFontColor(STYLE.ink)
    .setFontWeight('normal').setFontStyle('normal').setFontLine('none')
    .setBackground(null).setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

function formatSheet_(sh) {
  const max = sh.getMaxRows();
  const body = max - FIRST + 1;
  plain_(sh.getRange(1, 1, max, LAST_COL)).setBorder(false, false, false, false, false, false);

  sh.getRange(1, 1).setValue(TITLE);
  sh.getRange(1, 1, 1, LAST_COL).setFontWeight('bold').setFontSize(12).setFontColor(STYLE.title)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  sh.setRowHeight(1, 30);

  const head = sh.getRange(HEAD_ROW, 1, 1, LAST_COL);
  head.setValues([COLUMNS.map(function (c) { return c[0]; })])
    .setFontWeight('bold').setFontColor(STYLE.headInk).setBackground(STYLE.head)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sh.setRowHeight(HEAD_ROW, 34);

  COLUMNS.forEach(function (c, i) {
    const col = sh.getRange(FIRST, i + 1, body, 1);
    col.setHorizontalAlignment(c[2]);
    if (c[3]) col.setNumberFormat(c[3]);
    head.getCell(1, i + 1).setHorizontalAlignment(c[2]);
    sh.setColumnWidth(i + 1, c[1]);
  });
  sh.getRange(FIRST, COL.month, body, 1).setFontColor(STYLE.muted);
  sh.getRange(FIRST, 1, body, LAST_COL)
    .setBorder(null, null, true, null, null, true, STYLE.rule, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeights(FIRST, body, 21);
  sh.setFrozenRows(HEAD_ROW);
  sh.setHiddenGridlines(true);
}

/** The edited rows only, columns B to I — what an employee is allowed to touch. */
function formatRows_(sh, top, bottom) {
  const n = bottom - top + 1;
  plain_(sh.getRange(top, COL.lot, n, LAST_COL - COL.lot + 1))
    .setBorder(null, false, true, false, false, true, STYLE.rule, SpreadsheetApp.BorderStyle.SOLID);
  for (let c = COL.lot; c <= LAST_COL; c++) {
    const spec = COLUMNS[c - 1];
    const r = sh.getRange(top, c, n, 1).setHorizontalAlignment(spec[2]);
    if (spec[3]) r.setNumberFormat(spec[3]);
  }
}

function applyRules_(sh) {
  const body = sh.getMaxRows() - FIRST + 1;
  const range = function (c) { return sh.getRange(FIRST, c, body, 1); };
  const R = FIRST;
  sh.setConditionalFormatRules([
    // Amount greater → Swipe red (a Swipe not yet entered counts as nothing swiped)
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($F' + R + '),$F' + R + '>N($G' + R + '))')
      .setBackground(STYLE.red.fill).setFontColor(STYLE.red.ink).setBold(true)
      .setRanges([range(COL.swipe)]).build(),
    // Swipe greater → Amount green
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($G' + R + '),$G' + R + '>N($F' + R + '))')
      .setBackground(STYLE.green.fill).setFontColor(STYLE.green.ink).setBold(true)
      .setRanges([range(COL.amount)]).build(),
    // the same date, party and amount twice → amber Lot No., a likely double entry
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C' + R + '<>"",$D' + R + '<>"",ISNUMBER($F' + R + '),COUNTIFS($C$' + R +
        ':$C,$C' + R + ',$D$' + R + ':$D,$D' + R + ',$F$' + R + ':$F,$F' + R + ')>1)')
      .setBackground(STYLE.amber).setRanges([range(COL.lot)]).build(),
  ]);
}

function applyValidation_(ss, sh) {
  const body = sh.getMaxRows() - FIRST + 1;
  const range = function (c) { return sh.getRange(FIRST, c, body, 1); };
  const V = function () { return SpreadsheetApp.newDataValidation().setAllowInvalid(false); };
  const R = FIRST;
  range(COL.month).setDataValidation(null);
  range(COL.lot).setDataValidation(V()
    .requireFormulaSatisfied('=AND(REGEXMATCH(TO_TEXT(B' + R + '),"^PUR-\\d{3,}$"),COUNTIF($B$' + R +
      ':$B,B' + R + ')=1)')
    .setHelpText('Leave blank: the next PUR- number fills in by itself. If typed, it must be PUR- and a number not used before, e.g. PUR-746.')
    .build());
  range(COL.date).setDataValidation(V().requireDate().setHelpText('The purchase date, e.g. 05/10/2026.').build());
  const ps = ss.getSheetByName(PARTIES);
  if (ps) {
    range(COL.party).setDataValidation(V()
      .requireValueInRange(ps.getRange(2, 1, ps.getMaxRows() - 1, 1), true)
      .setHelpText('Pick the party from the list. A new party is added at the foot of the Parties sheet first.')
      .build());
  }
  range(COL.weight).setDataValidation(V().requireNumberGreaterThanOrEqualTo(0)
    .setHelpText('Kilograms, as a number. A count of boxes goes in Notes.').build());
  [COL.amount, COL.swipe].forEach(function (c) {
    range(c).setDataValidation(V().requireNumberGreaterThanOrEqualTo(0)
      .setHelpText('Rupees, as a number.').build());
  });
  range(COL.payment).setDataValidation(V()
    .requireFormulaSatisfied('=AND(ISDATE(H' + R + '),OR($C' + R + '="",H' + R + '>=$C' + R + '))')
    .setHelpText('The payment date, on or after the purchase date.').build());
  range(COL.notes).setDataValidation(null);
}

// ============================================================ other sheets
function writeParties_(ss, sh) {
  let ps = ss.getSheetByName(PARTIES);
  if (!ps) ps = ss.insertSheet(PARTIES);
  const names = {};
  const add = function (s) {
    if (typeof s !== 'string' || !s.trim()) return;
    const c = canonicalParty_(s);
    names[partyKey_(c)] = c;
  };
  const last = lastDataRow_(sh);
  if (last >= FIRST) sh.getRange(FIRST, COL.party, last - FIRST + 1, 1).getValues().forEach(function (r) { add(r[0]); });
  if (ps.getMaxRows() > 1) ps.getRange(2, 1, ps.getMaxRows() - 1, 1).getValues().forEach(function (r) { add(r[0]); });
  const list = Object.keys(names).map(function (k) { return names[k]; })
    .sort(function (a, b) { return a.toUpperCase() < b.toUpperCase() ? -1 : a.toUpperCase() > b.toUpperCase() ? 1 : 0; });
  const rows = Math.max(list.length + 200, 300);
  if (ps.getMaxRows() < rows + 1) ps.insertRowsAfter(ps.getMaxRows(), rows + 1 - ps.getMaxRows());
  ps.getRange(1, 1).setValue('Party');
  ps.getRange(2, 1, ps.getMaxRows() - 1, 1).clearContent();
  if (list.length) ps.getRange(2, 1, list.length, 1).setValues(list.map(function (s) { return [s]; }));
  return list;
}

function formatParties_(ps) {
  const max = ps.getMaxRows();
  plain_(ps.getRange(1, 1, max, 2)).setHorizontalAlignment('left');
  ps.getRange(1, 1, 1, 1).setFontWeight('bold').setFontColor(STYLE.headInk).setBackground(STYLE.head);
  ps.getRange(1, 2).setValue('One name per supplier. Add a new one in the first empty cell; it then appears in the Party list.');
  ps.getRange(1, 2).setFontColor(STYLE.muted).setFontStyle('italic');
  ps.getRange(2, 1, max - 1, 1).setNumberFormat('@')
    .setBorder(null, null, true, null, null, true, STYLE.rule, SpreadsheetApp.BorderStyle.SOLID);
  ps.setColumnWidth(1, 260);
  ps.setColumnWidth(2, 560);
  ps.setFrozenRows(1);
  ps.setHiddenGridlines(true);
  ps.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",COUNTIF($A$2:$A,$A2)>1)')
      .setBackground(STYLE.red.fill).setFontColor(STYLE.red.ink)
      .setRanges([ps.getRange(2, 1, max - 1, 1)]).build(),
  ]);
}

function writeLog_(ss, lines) {
  let ls = ss.getSheetByName(LOG);
  if (!ls) {
    ls = ss.insertSheet(LOG);
    ls.getRange(1, 1, 1, LOG_HEAD.length).setValues([LOG_HEAD]);
  }
  // a check already on the log is not listed again
  const at = lastLogRow_(ls) + 1;
  const known = {};
  if (at > 2) {
    ls.getRange(2, 1, at - 2, LOG_HEAD.length).getValues().forEach(function (l) {
      if (l[3] === 'Check') known[[l[1], l[2], l[6]].join('|')] = true;
    });
  }
  lines = lines.filter(function (l) { return l[3] !== 'Check' || !known[[l[1], l[2], l[6]].join('|')]; });
  if (!lines.length) lines = [[new Date(), '', '', 'Set-up', '', '', 'Run again — nothing needed changing']];
  if (ls.getMaxRows() < at + lines.length) ls.insertRowsAfter(ls.getMaxRows(), at + lines.length - ls.getMaxRows());
  ls.getRange(at, 1, lines.length, LOG_HEAD.length).setValues(lines);
}

function lastLogRow_(ls) {
  const v = ls.getRange(1, 1, ls.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (!isBlank_(v[i][0])) return i + 1;
  return 1;
}

function formatLog_(ls) {
  const max = ls.getMaxRows();
  plain_(ls.getRange(1, 1, max, LOG_HEAD.length)).setHorizontalAlignment('left');
  ls.getRange(1, 1, 1, LOG_HEAD.length).setFontWeight('bold').setFontColor(STYLE.headInk).setBackground(STYLE.head);
  ls.getRange(2, 1, max - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  ls.getRange(2, 2, max - 1, 1).setHorizontalAlignment('right');
  [130, 50, 90, 90, 210, 190, 520].forEach(function (w, i) { ls.setColumnWidth(i + 1, w); });
  ls.setFrozenRows(1);
  ls.setHiddenGridlines(true);
}

// ============================================================== protection
function protect_(ss, sh) {
  [sh, ss.getSheetByName(PARTIES), ss.getSheetByName(LOG)].forEach(function (s) {
    if (!s) return;
    [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE].forEach(function (type) {
      s.getProtections(type).forEach(function (p) {
        if (p.getDescription().indexOf(TAG) === 0) p.remove();
      });
    });
  });
  lockSheet_(sh, TAG + ' — purchases', [entryRange_(sh)]);
  const ps = ss.getSheetByName(PARTIES);
  if (ps) lockSheet_(ps, TAG + ' — parties', [ps.getRange(2, 1, ps.getMaxRows() - 1, 1)]);
  const ls = ss.getSheetByName(LOG);
  if (ls) lockSheet_(ls, TAG + ' — log', []);
}

function lockSheet_(s, description, open) {
  const p = s.protect().setDescription(description);
  p.setUnprotectedRanges(open);
  // Owner only. The current user is added first, as Google's own example does, so a
  // permission held through a group cannot lock the owner out.
  const me = Session.getEffectiveUser();
  p.addEditor(me);
  p.removeEditors(p.getEditors());
  if (p.canDomainEdit()) p.setDomainEdit(false);
  return p;
}

function entryRange_(sh) {
  return sh.getRange(FIRST, COL.lot, sh.getMaxRows() - FIRST + 1, LAST_COL - COL.lot + 1);
}

/** After rows are added, the new ones are opened for entry too. */
function openEntryRange_(sh) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
    if (p.getDescription().indexOf(TAG) === 0) p.setUnprotectedRanges([entryRange_(sh)]);
  });
}

/**
 * Runs fn holding the document lock. Returns false if another run holds it. Where
 * locking is not available (a simple trigger may be refused it), fn runs anyway:
 * the Lot No. rule still rejects a number used twice.
 */
function withLock_(ms, fn) {
  let lock = null, held = false;
  try { lock = LockService.getDocumentLock(); } catch (err) { lock = null; }
  if (lock) {
    try {
      held = lock.tryLock(ms);
      if (!held) return false;
    } catch (err) {
      held = false;
    }
  }
  try {
    fn();
    return true;
  } finally {
    if (held) lock.releaseLock();
  }
}
