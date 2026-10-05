// Runs purchase-statement/Code.gs against the original purchase statement, in a
// strict stand-in for Google Sheets, and checks the result cell for cell against
// what tools/build-purchase-xlsx.py produced for the Excel file.
//
//   python tools/build-purchase-xlsx.py --dump /tmp/ptest
//   node tools/test-purchase-gs.js /tmp/ptest
//
// The stand-in implements only the Apps Script methods Code.gs is expected to
// use, with Google's names and argument rules; a misspelt method, a range off
// the sheet or a bad enum value fails the run, as it would in Google Sheets.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const DIR = process.argv[2] || '/tmp/ptest';
const SRC = JSON.parse(fs.readFileSync(path.join(DIR, 'purchase-source.json'), 'utf8'));
const EXP = JSON.parse(fs.readFileSync(path.join(DIR, 'purchase-expected.json'), 'utf8'));
const CODE = fs.readFileSync(path.join(__dirname, '..', 'purchase-statement', 'Code.gs'), 'utf8');

const fails = [];
const fail = (m) => fails.push(m);
const decode = (v) => (v && typeof v === 'object' && v.$date)
  ? (([y, m, d]) => new Date(y, m - 1, d))(v.$date.split('-').map(Number)) : v;

// ------------------------------------------------------------------ the stand-in
const ENUM = {
  WrapStrategy: { WRAP: 'WRAP', OVERFLOW: 'OVERFLOW', CLIP: 'CLIP' },
  BorderStyle: { SOLID: 'SOLID', DOTTED: 'DOTTED', DASHED: 'DASHED' },
  ProtectionType: { SHEET: 'SHEET', RANGE: 'RANGE' },
};
const calls = [];
const need = (ok, msg) => { if (!ok) throw new Error(msg); };
const colour = (c) => c === null || /^#[0-9a-f]{6}$/i.test(c);

class Range {
  constructor(sheet, row, col, nr, nc) {
    need([row, col, nr, nc].every(Number.isInteger) && row >= 1 && col >= 1 && nr >= 1 && nc >= 1,
         `bad range ${row},${col},${nr},${nc}`);
    need(row + nr - 1 <= sheet.maxRows && col + nc - 1 <= sheet.maxCols,
         `range ${row},${col},${nr},${nc} outside ${sheet.name} (${sheet.maxRows}×${sheet.maxCols})`);
    Object.assign(this, { sheet, row, col, nr, nc });
  }
  rec(m, a) { calls.push({ sheet: this.sheet.name, m, r: [this.row, this.col, this.nr, this.nc], a }); return this; }
  getRow() { return this.row; } getColumn() { return this.col; }
  getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getSheet() { return this.sheet; }
  getCell(r, c) { need(r <= this.nr && c <= this.nc, 'getCell outside range'); return new Range(this.sheet, this.row + r - 1, this.col + c - 1, 1, 1); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nr; r++) {
      const row = [];
      for (let c = 0; c < this.nc; c++) row.push(this.sheet.get(this.row + r, this.col + c));
      out.push(row);
    }
    return out;
  }
  setValues(v) {
    need(Array.isArray(v) && v.length === this.nr && v.every((r) => r.length === this.nc),
         `setValues: ${v.length}×${v[0] && v[0].length} into ${this.nr}×${this.nc}`);
    v.forEach((row, r) => row.forEach((x, c) => this.sheet.set(this.row + r, this.col + c, x)));
    return this;
  }
  getValue() { return this.sheet.get(this.row, this.col); }
  setValue(x) { this.sheet.set(this.row, this.col, x); return this; }
  getFormula() { return this.sheet.formulas.get(`${this.row},${this.col}`) || ''; }
  setFormula(f) { need(f.startsWith('='), 'formula must start with ='); this.sheet.formulas.set(`${this.row},${this.col}`, f); this.sheet.set(this.row, this.col, ''); return this; }
  clearContent() {
    for (let r = 0; r < this.nr; r++) for (let c = 0; c < this.nc; c++) {
      this.sheet.set(this.row + r, this.col + c, ''); this.sheet.formulas.delete(`${this.row + r},${this.col + c}`);
    }
    return this;
  }
  setFontFamily(f) { need(typeof f === 'string', 'font'); return this.rec('font', f); }
  setFontSize(n) { need(Number.isInteger(n), 'size'); return this.rec('size', n); }
  setFontColor(c) { need(colour(c), `colour ${c}`); return this.rec('color', c); }
  setFontWeight(w) { need(['bold', 'normal', null].includes(w), `weight ${w}`); return this.rec('weight', w); }
  setFontStyle(s) { need(['italic', 'normal', null].includes(s), `style ${s}`); return this.rec('style', s); }
  setFontLine(l) { need(['underline', 'line-through', 'none', null].includes(l), `line ${l}`); return this.rec('line', l); }
  setBackground(c) { need(colour(c), `background ${c}`); return this.rec('background', c); }
  setNumberFormat(f) { need(typeof f === 'string', 'numberFormat'); return this.rec('numberFormat', f); }
  setHorizontalAlignment(a) { need(['left', 'center', 'right', 'normal', null].includes(a), `align ${a}`); return this.rec('align', a); }
  setVerticalAlignment(a) { need(['top', 'middle', 'bottom', null].includes(a), `valign ${a}`); return this.rec('valign', a); }
  setWrapStrategy(w) { need(Object.values(ENUM.WrapStrategy).includes(w), `wrap ${w}`); return this.rec('wrap', w); }
  setBorder(...a) {
    need(a.length === 6 || a.length === 8, `setBorder takes 6 or 8 arguments, got ${a.length}`);
    need(a.slice(0, 6).every((x) => x === true || x === false || x === null), 'setBorder sides are true/false/null');
    if (a.length === 8) need(colour(a[6]) && Object.values(ENUM.BorderStyle).includes(a[7]), 'setBorder colour/style');
    return this.rec('border', a);
  }
  setDataValidation(rule) { need(rule === null || rule instanceof Rule, 'validation rule'); this.sheet.validations.push({ r: [this.row, this.col, this.nr, this.nc], rule }); return this; }
}

class Sheet {
  constructor(name, rows, cols) { Object.assign(this, { name, maxRows: rows, maxCols: cols, cells: new Map(), formulas: new Map(), validations: [], cf: [], protections: [], frozen: 0, widths: {} }); }
  get(r, c) { const v = this.cells.get(`${r},${c}`); return v === undefined ? '' : v; }
  set(r, c, v) { need(v !== undefined && v !== null, `writing ${v} to ${r},${c}`); if (v === '') this.cells.delete(`${r},${c}`); else this.cells.set(`${r},${c}`, v); }
  getName() { return this.name; }
  getMaxRows() { return this.maxRows; } getMaxColumns() { return this.maxCols; }
  getRange(...a) { need(a.length >= 2 && a.length <= 4 && a.every(Number.isInteger), `getRange(${a}) — A1 strings not used here`); return new Range(this, a[0], a[1], a[2] || 1, a[3] || 1); }
  insertRowsAfter(after, n) { need(after === this.maxRows && n > 0, `insertRowsAfter(${after}, ${n}) — only appending is expected`); this.maxRows += n; return this; }
  setColumnWidth(c, w) { need(Number.isInteger(c) && Number.isInteger(w), 'width'); this.widths[c] = w; return this; }
  setRowHeight(r, h) { need(r <= this.maxRows && Number.isInteger(h), 'row height'); return this; }
  setRowHeights(r, n, h) { need(r + n - 1 <= this.maxRows && Number.isInteger(h), 'row heights'); return this; }
  setFrozenRows(n) { this.frozen = n; return this; }
  setHiddenGridlines(b) { need(typeof b === 'boolean', 'gridlines'); return this; }
  setConditionalFormatRules(rules) { need(rules.every((r) => r instanceof CfRule), 'cf rules'); this.cf = rules; return this; }
  protect() { const p = new Protection(this, 'SHEET'); this.protections.push(p); return p; }
  getProtections(t) { need(Object.values(ENUM.ProtectionType).includes(t), 'protection type'); return this.protections.filter((p) => p.type === t && !p.removed); }
}

const OWNER = { getEmail: () => 'owner@example.com' };
class Protection {
  constructor(sheet, type) { Object.assign(this, { sheet, type, description: '', open: [], editors: [OWNER, { getEmail: () => 'staff@example.com' }], domain: true, removed: false }); }
  setDescription(d) { this.description = d; return this; } getDescription() { return this.description; }
  setUnprotectedRanges(r) { need(this.type === 'SHEET' && r.every((x) => x instanceof Range && x.sheet === this.sheet), 'unprotected ranges'); this.open = r; return this; }
  addEditor(u) { if (!this.editors.includes(u)) this.editors.push(u); return this; }
  removeEditors(list) { this.editors = this.editors.filter((u) => u === OWNER || !list.includes(u)); return this; }
  getEditors() { return this.editors.slice(); }
  canDomainEdit() { return this.domain; } setDomainEdit(b) { this.domain = b; return this; }
  remove() { this.removed = true; }
}

class CfRule { constructor(o) { Object.assign(this, o); } }
class CfBuilder {
  constructor() { this.o = {}; }
  whenFormulaSatisfied(f) { need(f.startsWith('='), 'cf formula'); this.o.formula = f; return this; }
  setBackground(c) { need(colour(c), 'cf colour'); this.o.background = c; return this; }
  setFontColor(c) { need(colour(c), 'cf font colour'); this.o.color = c; return this; }
  setBold(b) { this.o.bold = b; return this; }
  setRanges(r) { need(r.every((x) => x instanceof Range), 'cf ranges'); this.o.ranges = r; return this; }
  build() { need(this.o.formula && this.o.ranges, 'cf rule incomplete'); return new CfRule(this.o); }
}
class Rule { constructor(o) { Object.assign(this, o); } }
class DvBuilder {
  constructor() { this.o = {}; }
  set(k, v) { need(!this.o.kind, 'two criteria on one rule'); this.o.kind = k; this.o.arg = v; return this; }
  requireFormulaSatisfied(f) { need(f.startsWith('='), 'dv formula'); return this.set('formula', f); }
  requireDate() { return this.set('date', null); }
  requireNumberGreaterThanOrEqualTo(n) { return this.set('number>=', n); }
  requireValueInRange(r, show) { need(r instanceof Range && typeof show === 'boolean', 'value in range'); return this.set('inRange', r); }
  setAllowInvalid(b) { this.o.allowInvalid = b; return this; }
  setHelpText(t) { this.o.help = t; return this; }
  build() { need(this.o.kind, 'dv rule has no criteria'); return new Rule(this.o); }
}

class Spreadsheet {
  constructor() { this.sheets = []; this.toasts = []; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  getSheets() { return this.sheets.slice(); }
  insertSheet(n) { need(!this.getSheetByName(n), `sheet ${n} exists`); const s = new Sheet(n, 1000, 26); this.sheets.push(s); return s; }
  getSpreadsheetTimeZone() { return 'Asia/Kolkata'; }
  toast(m, t, s) { this.toasts.push(m); }
}

const triggers = [];
const ss = new Spreadsheet();
const ctx = vm.createContext({
  console: { warn: (m) => fail('console.warn: ' + m), log: () => {} },
  SpreadsheetApp: Object.assign({
    getActiveSpreadsheet: () => active,
    newConditionalFormatRule: () => new CfBuilder(),
    newDataValidation: () => new DvBuilder(),
    getUi: () => { throw new Error('getUi is not available here'); },
  }, ENUM),
  Session: { getEffectiveUser: () => OWNER },
  // NO_LOCK=1 plays a simple trigger that is refused the lock
  LockService: { getDocumentLock: () => { if (process.env.NO_LOCK) throw new Error('no permission'); return { tryLock: () => true, releaseLock: () => {} }; } },
  Utilities: {
    formatDate: (d, tz, f) => {
      need(f === 'dd/MM/yyyy' && typeof tz === 'string', `formatDate ${f}`);
      const p = (n) => String(n).padStart(2, '0');
      return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()}`;
    },
  },
  ScriptApp: {
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: (t) => triggers.splice(triggers.indexOf(t), 1),
    newTrigger: (fn) => ({
      forSpreadsheet: (s) => ({ onChange: () => ({ create: () => { const t = { fn, kind: 'change', getHandlerFunction: () => fn }; triggers.push(t); return t; } }) }),
      timeBased: () => ({ everyDays: (n) => ({ atHour: (h) => ({ create: () => { const t = { fn, kind: `daily@${h}`, getHandlerFunction: () => fn }; triggers.push(t); return t; } }) }) }),
    }),
  },
});

// ------------------------------------------------------------------ the original sheet
// TAB=Sheet1 plays a file whose purchase tab was never named "Purchase"
let active = ss;
const TAB = process.env.TAB || 'Purchase';
ss.sheets.push(new Sheet('Notes', 100, 10));      // another tab, ahead of the purchases
const sh = new Sheet(TAB, SRC.maxRows, Math.max(SRC.maxCols, 26));
ss.sheets.push(sh);
for (const [k, v] of Object.entries(SRC.cells)) {
  const [r, c] = k.split(',').map(Number);
  sh.set(r, c, decode(v));
}

vm.runInContext(CODE, ctx, { filename: 'Code.gs' });
const run = (src) => vm.runInContext(src, ctx);

// ------------------------------------------------------------------ 0. what Run runs, and the plain-English errors
const firstFn = (CODE.match(/^function (\w+)/m) || [])[1];
if (firstFn !== 'setup') fail(`first function is ${firstFn}: the editor's Run button would not run set-up`);
const throws = (src) => { try { run(src); return ''; } catch (err) { return String(err.message || err); } };
active = null;
if (!/not attached to a spreadsheet/.test(throws('setup()'))) fail('unbound script: no clear error');
active = new Spreadsheet(); active.sheets.push(new Sheet('Sheet1', 50, 10));
const noTab = throws('setup()');
if (!/Could not find the purchase sheet.*"Sheet1"/.test(noTab)) fail(`no purchase tab: ${noTab}`);
active = ss;

// ------------------------------------------------------------------ 1. set-up
run('setup()');
const FIRST = EXP.first, LASTROW = FIRST + 5000 - 1;
const same = (a, b) => {
  const n = (x) => (x instanceof Date ? x.toISOString().slice(0, 10)
    : x && typeof x === 'object' && x.$date ? decode(x).toISOString().slice(0, 10)
      : typeof x === 'number' ? Math.round(x * 1e6) / 1e6 : x);
  return n(a) === n(b);
};
EXP.rows.forEach((want, i) => {
  const r = FIRST + i;
  want.forEach((w, c) => {
    const got = sh.get(r, c + 2);
    if (!same(got, w === null ? '' : w)) fail(`row ${r} col ${c + 2}: got ${JSON.stringify(got)} want ${JSON.stringify(w)}`);
  });
});
for (let r = EXP.last + 1; r <= sh.maxRows; r++) for (let c = 2; c <= 9; c++) {
  if (sh.get(r, c) !== '') fail(`row ${r} col ${c} should be empty`);
}
if (sh.maxRows < LASTROW) fail(`sheet has ${sh.maxRows} rows, want at least ${LASTROW}`);
const monthF = sh.formulas.get(`${FIRST},1`);
if (!/^=ARRAYFORMULA\(IF\(C3:C="","",UPPER\(TEXT\(C3:C,"mmm"\)\)\)\)$/.test(monthF || '')) fail(`month formula: ${monthF}`);
for (let r = FIRST + 1; r <= sh.maxRows; r++) if (sh.get(r, 1) !== '') { fail(`A${r} not cleared for the array formula`); break; }
if (sh.get(1, 1) !== 'PREMIER EXPORTS INTERNATIONAL  ·  PURCHASE STATEMENT') fail('title');
if (sh.get(2, 1) !== 'Month' || sh.get(2, 2) !== 'Lot No.' || sh.get(2, 9) !== 'Notes') fail('headings');

const ps = ss.getSheetByName('Parties');
const plist = [];
for (let r = 2; r <= ps.maxRows && ps.get(r, 1) !== ''; r++) plist.push(ps.get(r, 1));
if (JSON.stringify(plist) !== JSON.stringify(EXP.parties)) fail(`parties: ${JSON.stringify(plist)}`);

const ls = ss.getSheetByName('Cleanup log');
const lines = [];
for (let r = 2; r <= ls.maxRows && ls.get(r, 1) !== ''; r++) lines.push([2, 3, 4, 5, 6, 7].map((c) => ls.get(r, c)));
if (lines.length !== EXP.log.length) fail(`log has ${lines.length} lines, want ${EXP.log.length}`);
EXP.log.forEach((want, i) => {
  const got = lines[i] || [];
  want.forEach((w, c) => { if (!same(got[c], w === null ? '' : w)) fail(`log line ${i + 1} col ${c}: got ${JSON.stringify(got[c])} want ${JSON.stringify(w)}`); });
});

// conditional formatting: exactly the three rules, on the right columns
const cf = sh.cf.map((r) => ({ f: r.formula, bg: r.background, col: r.ranges[0].col, rows: [r.ranges[0].row, r.ranges[0].row + r.ranges[0].nr - 1] }));
const wantCf = [
  { f: '=AND(ISNUMBER($F3),$F3>N($G3))', bg: '#f4c7c3', col: 7 },
  { f: '=AND(ISNUMBER($G3),$G3>N($F3))', bg: '#b7e1cd', col: 6 },
  { f: '=AND($C3<>"",$D3<>"",ISNUMBER($F3),COUNTIFS($C$3:$C,$C3,$D$3:$D,$D3,$F$3:$F,$F3)>1)', bg: '#fce8b2', col: 2 },
];
if (cf.length !== 3) fail(`${cf.length} conditional format rules`);
wantCf.forEach((w, i) => {
  const g = cf[i] || {};
  if (g.f !== w.f || g.bg !== w.bg || g.col !== w.col || g.rows[0] !== FIRST || g.rows[1] !== sh.maxRows) fail(`cf rule ${i + 1}: ${JSON.stringify(g)}`);
});

// validation: the last rule set on each column wins
const dv = {};
sh.validations.forEach((v) => { if (v.r[0] === FIRST && v.r[2] === sh.maxRows - FIRST + 1 && v.r[3] === 1) dv[v.r[1]] = v.rule; });
const kind = (c) => (dv[c] === undefined ? 'unset' : dv[c] === null ? 'none' : dv[c].kind + (dv[c].allowInvalid === false ? '/reject' : '/warn'));
const wantDv = { 1: 'none', 2: 'formula/reject', 3: 'date/reject', 4: 'inRange/reject', 5: 'number>=/reject', 6: 'number>=/reject', 7: 'number>=/reject', 8: 'formula/reject', 9: 'none' };
for (const [c, w] of Object.entries(wantDv)) if (kind(+c) !== w) fail(`validation col ${c}: ${kind(+c)} want ${w}`);
if (dv[2] && dv[2].arg !== '=AND(REGEXMATCH(TO_TEXT(B3),"^PUR-\\d{3,}$"),COUNTIF($B$3:$B,B3)=1)') fail(`lot rule: ${dv[2].arg}`);
if (dv[4] && dv[4].arg.sheet !== ps) fail('party list is not the Parties sheet');

// protection: owner only, entry cells open
const prot = sh.getProtections('SHEET');
if (prot.length !== 1) fail(`${prot.length} sheet protections on Purchase`);
const p0 = prot[0];
if (p0) {
  const o = p0.open.map((r) => [r.row, r.col, r.nr, r.nc].join(','));
  if (JSON.stringify(o) !== JSON.stringify([[FIRST, 2, sh.maxRows - FIRST + 1, 8].join(',')])) fail(`open range ${o}`);
  if (p0.editors.length !== 1 || p0.editors[0] !== OWNER || p0.domain) fail('Purchase protection is not owner-only');
}
const pp = ps.getProtections('SHEET')[0];
if (!pp || pp.open.length !== 1 || pp.open[0].col !== 1 || pp.open[0].row !== 2) fail('Parties protection');
const lp = ls.getProtections('SHEET')[0];
if (!lp || lp.open.length !== 0) fail('Cleanup log should be fully locked');
const trig = triggers.map((t) => `${t.fn}:${t.kind}`).sort();
if (JSON.stringify(trig) !== JSON.stringify(['nightlyRestore:daily@2', 'onChangeInstalled:change'])) fail(`triggers ${trig}`);

// formatting covers the whole table in one font, and every column has its format
const fontAll = calls.some((c) => c.sheet === TAB && c.m === 'font' && c.a === 'Arial' && c.r[0] === 1 && c.r[2] === sh.maxRows && c.r[3] === 9);
if (!fontAll) fail('no single font over the whole table');
const fmts = {};
calls.filter((c) => c.sheet === TAB && c.m === 'numberFormat' && c.r[0] === FIRST && c.r[3] === 1).forEach((c) => { fmts[c.r[1]] = c.a; });
const wantF = { 2: '@', 3: 'dd/mm/yyyy', 4: '@', 5: '#,##0.00', 6: '[$₹]#,##0.00', 7: '[$₹]#,##0.00', 8: 'dd/mm/yyyy', 9: '@' };
for (const [c, f] of Object.entries(wantF)) if (fmts[c] !== f) fail(`number format col ${c}: ${fmts[c]}`);
if (sh.frozen !== 2) fail('frozen rows');
if (sh.widths[4] !== 230) fail('party column width');

// ------------------------------------------------------------------ 2. set-up again changes nothing
const before = JSON.stringify([...sh.cells]);
run('setup()');
if (JSON.stringify([...sh.cells]) !== before) fail('second set-up changed the data');
const tail = [2, 3, 4, 5, 6, 7].map((c) => ls.get(lines.length + 2, c));
if (tail[2] !== 'Set-up' || tail[5] !== 'Run again — nothing needed changing') fail(`second run log: ${JSON.stringify(tail)}`);
if (sh.getProtections('SHEET').length !== 1) fail('second set-up stacked protections');
if (triggers.length !== 2) fail('second set-up stacked triggers');

// ------------------------------------------------------------------ 3. an employee's day
const nextRow = EXP.last + 1 - 9;                // the first row after the last purchase
const edit = (r, c, nr, nc) => { ctx.__e = { range: sh.getRange(r, c, nr, nc) }; run('onEdit(__e)'); };
sh.set(nextRow, 3, new Date(2026, 9, 5)); sh.set(nextRow, 4, 'FBS');
edit(nextRow, 3, 1, 2);
if (sh.get(nextRow, 2) !== 'PUR-737') fail(`new row got ${sh.get(nextRow, 2)}, want PUR-737`);
for (let k = 1; k <= 3; k++) { sh.set(nextRow + k, 3, new Date(2026, 9, 5)); sh.set(nextRow + k, 6, 1000 * k); }
edit(nextRow + 1, 3, 3, 4);                        // a pasted block of three
['PUR-738', 'PUR-739', 'PUR-740'].forEach((w, k) => { if (sh.get(nextRow + 1 + k, 2) !== w) fail(`pasted row ${k + 1}: ${sh.get(nextRow + 1 + k, 2)}`); });
sh.set(10, 4, 'EKS'); edit(10, 4, 1, 1);           // editing an old row keeps its number
if (sh.get(10, 2) !== 'PUR-008') fail(`old row renumbered: ${sh.get(10, 2)}`);
sh.set(nextRow + 4, 9, 'note only'); edit(nextRow + 4, 9, 1, 1);
if (sh.get(nextRow + 4, 2) !== '') fail('a note alone should not take a lot number');
const fmtEdit = calls.filter((c) => c.m === 'numberFormat' && c.r[0] === nextRow + 1 && c.r[2] === 3);
if (fmtEdit.length !== 8) fail(`pasted rows reformatted in ${fmtEdit.length} columns, want 8 (B–I)`);
if (calls.some((c) => c.r && c.r[0] === nextRow + 1 && c.r[2] === 3 && c.r[1] === 1)) fail('onEdit touched the protected Month column');
edit(1, 1, 2, 9);                                   // headings: nothing to number

// ------------------------------------------------------------------ 4. formatting changed → put back; rows run low → more added
const n0 = calls.length;
ctx.__e = { changeType: 'EDIT' }; run('onChangeInstalled(__e)');
if (calls.length !== n0) fail('a plain edit should not trigger a full restore');
ctx.__e = { changeType: 'FORMAT' }; run('onChangeInstalled(__e)');
if (calls.length === n0) fail('a format change did not restore');
const max0 = sh.maxRows;
sh.set(max0 - 50, 3, new Date(2026, 9, 5));
run('nightlyRestore()');
if (sh.maxRows !== max0 + 1000) fail(`rows not extended: ${sh.maxRows}`);
const open = sh.getProtections('SHEET')[0].open[0];
if (open.row + open.nr - 1 !== sh.maxRows) fail('new rows not opened for entry');

console.log(fails.length ? fails.slice(0, 40).join('\n') + (fails.length > 40 ? `\n… ${fails.length - 40} more` : '')
  : `all checks passed — ${EXP.rows.length} rows match the Excel build cell for cell, ${EXP.log.length} log lines, `
    + `${plist.length} parties, 3 colour rules, 9 column rules, owner-only protection, 2 triggers; `
    + 'set-up re-runs cleanly; new rows numbered PUR-737 onwards; formatting restored on change');
process.exit(fails.length ? 1 : 0);
