// Builds PEI-Bills-on-Hand.docx — the Word counterpart of bills/index.html:
// the same letterhead, tiles, buyer chart, advances panel and bills table,
// laid out on one A4 page.
//
//   npm install docx
//   node tools/build-bills-docx.js bills/PEI-Bills-on-Hand.docx [--fx 90] [--date 30.09.2026]
//
// The bills are read from the DATA array in bills/index.html, so the page and
// the Word file cannot disagree about a figure. Word does not recalculate: to
// change the exchange rate, rebuild with --fx.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, HeightRule,
  Tab, TabStopType,
} = require('docx');

// ---------------------------------------------------------------- arguments
const argv = process.argv.slice(2);
const opt = (k, d) => { const i = argv.indexOf(k); return i >= 0 ? argv[i + 1] : d; };
const OUT = argv.find((a, i) => !a.startsWith('--') && !(i > 0 && argv[i - 1].startsWith('--')))
  || 'bills/PEI-Bills-on-Hand.docx';
const FX = parseFloat(opt('--fx', '90'));
const AS_AT = opt('--date', '30.09.2026');
const LIST_DATE = '30/09/26';

// ---------------------------------------------------------------- the bills
// [invoice, item, buyer, amount US$, advance received US$, note]
const html = fs.readFileSync(path.join(__dirname, '..', 'bills', 'index.html'), 'utf8');
const m = html.match(/var DATA = \[([\s\S]*?)\n\s*\];/);
if (!m) throw new Error('DATA array not found in bills/index.html');
const BILLS = Function('"use strict"; return [' + m[1] + '];')();

// ---------------------------------------------------------------- tokens
// Same design system as the page and the purchase order.
const BRAND = '00A2D3', DEEP = '065E7A', INK = '0B2C3A', MUTED = '5E7C8B';
const WASH = 'F2F8FB', BAND = 'E7F3F9', LINE = 'CBDFE8', RULE = '9EC4D5', WHITE = 'FFFFFF';
const ON_DEEP = 'CFE8F2';

const FS = { micro: 14, fine: 15, small: 17, body: 18, lead: 26, hero: 34, title: 39, mark: 37 };
const SANS = 'Arial', MONO = 'Consolas';
const W = 10772;                               // A4 less 1 cm margins, twips
const GAP = 120;                               // between tiles and panels
const PAD = { x: 110, y: 30 };

const hair = { style: BorderStyle.SINGLE, size: 6, color: LINE };
const rule = (sz = 10) => ({ style: BorderStyle.SINGLE, size: sz, color: RULE });
const none = { style: BorderStyle.NONE, size: 0, color: WHITE };
const boxAll = { top: hair, bottom: hair, left: hair, right: hair };
const boxNone = { top: none, bottom: none, left: none, right: none };

// ---------------------------------------------------------------- numbers
// US$ in international grouping, rupees in lakh/crore grouping.
const usd = (n) => n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const usd0 = (n) => '$ ' + Math.round(n).toLocaleString('en-US');
const inr = (n) => n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const compactINR = (n) => {
  const a = Math.abs(n);
  if (a >= 1e7) return '₹ ' + (n / 1e7).toFixed(2) + ' Cr';
  if (a >= 1e5) return '₹ ' + (n / 1e5).toFixed(2) + ' L';
  return '₹ ' + Math.round(n).toLocaleString('en-IN');
};
if (inr(12341658.2) !== '1,23,41,658.20') throw new Error('Node lacks en-IN grouping (needs full ICU)');

// ---------------------------------------------------------------- helpers
const run = (text, o = {}) => new TextRun({
  text, font: o.font || SANS, size: o.size || FS.body,
  color: o.color || INK, bold: o.bold, italics: o.italics, characterSpacing: o.ls,
});
const mono = (text, o = {}) => run(text, { ...o, font: MONO });

const para = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: o.align,
  spacing: { before: o.before || 0, after: o.after || 0, line: o.line, lineRule: o.line ? 'auto' : undefined },
  keepNext: o.keepNext,
});

// an uppercase micro label — the one tracked style used everywhere
const label = (text, o = {}) => para(
  run(text.toUpperCase(), { size: o.size || FS.micro, bold: true, color: o.color || MUTED, ls: 24 }),
  { align: o.align, after: o.after === undefined ? 20 : o.after });

const tiny = () => para(run('', { size: 2 }), { line: 24 });   // Word needs a paragraph after a nested table

const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children],
  width: { size: o.w, type: WidthType.DXA },
  columnSpan: o.span,
  borders: o.borders || boxAll,
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  margins: {
    top: o.py === undefined ? PAD.y : o.py, bottom: o.py === undefined ? PAD.y : o.py,
    left: o.px === undefined ? PAD.x : o.px, right: o.px === undefined ? PAD.x : o.px,
  },
  verticalAlign: o.valign || VerticalAlign.CENTER,
});

const table = (widths, rows, o = {}) => {
  const sum = widths.reduce((a, b) => a + b, 0);
  if (o.check !== false && sum !== (o.total || W)) throw new Error(`table widths ${sum} != ${o.total || W}`);
  return new Table({
    columnWidths: widths,
    width: { size: sum, type: WidthType.DXA },
    rows, borders: boxNone, alignment: o.align,
  });
};
const spacer = (h) => para(run('', { size: 2 }), { after: h, line: 24 });

// ---------------------------------------------------------------- figures
const T = { usd: 0, adv: 0, n: BILLS.length };
const mix = { mackerel: 0, tuna: 0, other: 0 };
const byBuyer = new Map();
const UNNOTED = 'Buyer not noted';
for (const [, item, buyer, amt, adv] of BILLS) {
  T.usd += amt; T.adv += adv;
  const k = /^tuna$/i.test(item) ? 'tuna' : /^mackerel$/i.test(item) ? 'mackerel' : 'other';
  mix[k]++;
  const b = (buyer || '').trim() || UNNOTED;
  const e = byBuyer.get(b) || { name: b, billed: 0, adv: 0, n: 0 };
  e.billed += amt; e.adv += adv; e.n++;
  byBuyer.set(b, e);
}
const BAL = T.usd - T.adv;
const buyers = [...byBuyer.values()].map((b) => ({ ...b, bal: b.billed - b.adv }))
  .filter((b) => b.bal > 0).sort((a, b) => b.bal - a.bal);
const advRows = BILLS.filter((r) => r[4] > 0);

// ---------------------------------------------------------------- letterhead
const logo = fs.readFileSync(path.join(__dirname, '..', 'stock-statement', 'assets', 'pei-mark.png'));
const LH = [1000, 4100, 5672];
const letterhead = table(LH, [new TableRow({ children: [
  cell(para(new ImageRun({ data: logo, type: 'png', transformation: { width: 62, height: 50 } })),
    { w: LH[0], borders: boxNone, py: 0, px: 0 }),
  cell([
    new Paragraph({ children: [run('Premier Exports', { size: FS.mark, color: BRAND })],
      spacing: { after: 0, line: 420, lineRule: 'exact' } }),
    new Paragraph({ children: [run('International', { size: FS.mark, color: BRAND })],
      spacing: { after: 0, line: 420, lineRule: 'exact' } }),
  ], { w: LH[1], borders: boxNone, py: 0 }),
  cell([
    para(run('AP X/453, NH-66 Highway, Chandiroor P.O.,', { size: FS.small, color: MUTED }), { align: AlignmentType.RIGHT }),
    para(run('Aroor, Alappuzha, Kerala - 688 537, India', { size: FS.small, color: MUTED }), { align: AlignmentType.RIGHT }),
    para(mono('GSTIN 32AADFP3158P1ZZ', { size: FS.small }), { align: AlignmentType.RIGHT }),
  ], { w: LH[2], borders: boxNone, py: 0, px: 0 }),
] })]);

const brandRule = new Paragraph({
  children: [],
  border: { bottom: { style: BorderStyle.SINGLE, size: 20, color: BRAND, space: 1 } },
  spacing: { before: 80, after: 140 },
});

// ---------------------------------------------------------------- title band
const DB = [780, 1420, 1100, 900];
const dateBox = table(DB, [new TableRow({ children: [
  cell(label('As at', { after: 0 }), { w: DB[0], fill: WASH }),
  cell(para(mono(AS_AT, { bold: true }), { align: AlignmentType.RIGHT }), { w: DB[1] }),
  cell(label('US$ → ₹', { after: 0 }), { w: DB[2], fill: WASH }),
  cell(para(mono(FX.toFixed(2), { bold: true }), { align: AlignmentType.RIGHT }), { w: DB[3] }),
] })], { check: false, align: AlignmentType.RIGHT });
const TBW = [W - 4300, 4300];
const titleBand = table(TBW, [new TableRow({ children: [
  cell([
    para(run('BILLS ON HAND', { size: FS.title, bold: true, ls: 30 }), { after: 40 }),
    label('Export bills · billed, advances received, balance due', { after: 0 }),
  ], { w: TBW[0], borders: boxNone, px: 0, valign: VerticalAlign.BOTTOM }),
  cell([dateBox, tiny()], { w: TBW[1], borders: boxNone, px: 0, py: 0, valign: VerticalAlign.BOTTOM }),
] })]);

// ---------------------------------------------------------------- tiles
const TW = (() => { const u = (W - 3 * GAP) / 4.25; const a = Math.round(u);
  return [a, a, a, W - 3 * GAP - 3 * a]; })();
const tile = (k, v, u, o = {}) => cell([
  label(k, { color: o.accent ? ON_DEEP : MUTED, after: 30 }),
  para(mono(v, { size: o.accent ? FS.hero : FS.lead, bold: true, color: o.accent ? WHITE : INK }), { after: 20 }),
  o.plain
    ? para(run(u, { size: FS.micro, color: MUTED }))
    : para(mono(u, { size: FS.small, bold: true, color: o.accent ? WHITE : INK })),
], { w: o.w, fill: o.accent ? DEEP : WASH, valign: VerticalAlign.TOP, py: 70,
     borders: o.accent ? { top: { ...hair, color: DEEP }, bottom: { ...hair, color: DEEP },
                           left: { ...hair, color: DEEP }, right: { ...hair, color: DEEP } } : boxAll });
const gapCell = (w) => cell(para(run('', { size: 2 })), { w, borders: boxNone, px: 0, py: 0 });
const mixText = [mix.mackerel && `${mix.mackerel} mackerel`, mix.tuna && `${mix.tuna} tuna`,
                 mix.other && `${mix.other} other`].filter(Boolean).join(' · ');
const tiles = table([TW[0], GAP, TW[1], GAP, TW[2], GAP, TW[3]], [new TableRow({ children: [
  tile('Bills on hand', String(T.n), mixText, { w: TW[0], plain: true }), gapCell(GAP),
  tile('Total billed', usd0(T.usd), compactINR(T.usd * FX), { w: TW[1] }), gapCell(GAP),
  tile('Advances received', usd0(T.adv), compactINR(T.adv * FX), { w: TW[2] }), gapCell(GAP),
  tile('Balance receivable', usd0(BAL), compactINR(BAL * FX), { w: TW[3], accent: true }),
] })]);

// ---------------------------------------------------------------- panels
const PW = [5900, GAP, W - 5900 - GAP];
const panelHead = (text, note, inner) => new Paragraph({
  children: [
    run(text.toUpperCase(), { size: FS.micro, bold: true, color: DEEP, ls: 24 }),
    ...(note ? [new TextRun({ children: [new Tab()] }), run(note, { size: FS.micro, color: MUTED })] : []),
  ],
  tabStops: note ? [{ type: TabStopType.RIGHT, position: inner }] : undefined,
  spacing: { after: 70 },
});

// chart — one measure, one hue; bars are shaded cells on a 2% grid so the
// chart stays native Word, crisp at any print resolution
const GRID = 50, GW = 62;
const CW = [1510, ...Array(GRID).fill(GW), PW[0] - 2 * PAD.x - 1510 - GRID * GW];
const max = buyers.length ? buyers[0].bal : 1;
const gapB = { style: BorderStyle.SINGLE, size: 20, color: WHITE };
const barRows = buyers.map((b) => {
  const k = Math.max(1, Math.round(b.bal / max * GRID));
  const unnoted = b.name === UNNOTED;
  const barCell = (span, fill) => cell(para(run('', { size: 2 })),
    { w: span * GW, span, fill, px: 0, py: 0,
      borders: { top: gapB, bottom: gapB, left: none, right: none } });
  return new TableRow({
    height: { value: 228, rule: HeightRule.EXACT },
    children: [
      cell(para(run(b.name, { size: FS.fine, color: unnoted ? MUTED : INK, italics: unnoted }),
        { align: AlignmentType.RIGHT }), { w: CW[0], borders: boxNone, py: 0, px: 0 + 80 }),
      barCell(k, DEEP),
      ...(k < GRID ? [barCell(GRID - k, WASH)] : []),
      cell(para(mono(usd0(b.bal), { size: FS.fine }), { align: AlignmentType.RIGHT }),
        { w: CW[CW.length - 1], borders: boxNone, py: 0, px: 0 }),
    ],
  });
});
const chart = table(CW, barRows, { total: PW[0] - 2 * PAD.x });

// advances
const AW = [600, 1290, 1180, PW[2] - 2 * PAD.x - 600 - 1290 - 1180];
const aTh = (t, i) => cell(label(t, { after: 0, align: i >= 2 ? AlignmentType.RIGHT : AlignmentType.LEFT }),
  { w: AW[i], px: 60, borders: { top: none, left: none, right: none, bottom: rule(8) } });
const aTd = (child, i, o = {}) => cell(child, { w: AW[i], px: 60,
  borders: { top: none, left: none, right: none, bottom: o.last ? none : hair, ...(o.top ? { top: rule(10) } : {}) } });
const advBilled = advRows.reduce((s, r) => s + r[3], 0);
const advTable = table(AW, [
  new TableRow({ children: ['Inv.', 'Buyer', 'US$', '₹'].map(aTh) }),
  ...(advRows.length ? advRows.map((r) => new TableRow({ children: [
    aTd(para(run(String(r[0]).replace(/^PEI\//, '').replace(/\/\d{4}-\d{2}$/, ''), { size: FS.fine })), 0),
    aTd(para(run(r[2] || UNNOTED, { size: FS.fine })), 1),
    aTd(para(mono(usd(r[4]), { size: FS.fine }), { align: AlignmentType.RIGHT }), 2),
    aTd(para(mono(inr(r[4] * FX), { size: FS.fine }), { align: AlignmentType.RIGHT }), 3),
  ] })) : [new TableRow({ children: [cell(para(run('No advances recorded.', { size: FS.fine, color: MUTED, italics: true })),
      { w: PW[2] - 2 * PAD.x, span: 4, borders: boxNone })] })]),
  ...(advRows.length ? [new TableRow({ children: [
    cell(para(run(`${advRows.length} ${advRows.length === 1 ? 'bill' : 'bills'}`, { size: FS.fine, bold: true })),
      { w: AW[0] + AW[1], span: 2, px: 60, borders: { top: rule(10), bottom: none, left: none, right: none } }),
    aTd(para(mono(usd(T.adv), { size: FS.fine, bold: true }), { align: AlignmentType.RIGHT }), 2, { last: true, top: true }),
    aTd(para(mono(inr(T.adv * FX), { size: FS.fine, bold: true }), { align: AlignmentType.RIGHT }), 3, { last: true, top: true }),
  ] })] : []),
], { total: PW[2] - 2 * PAD.x });
const advNote = advRows.length
  ? `These ${advRows.length === 1 ? 'bill was' : advRows.length + ' bills were'} raised for US$ ${usd(advBilled)}; ` +
    `US$ ${usd(advBilled - T.adv)} is still to come in on ${advRows.length === 1 ? 'it' : 'them'}.`
  : '';

const panels = table(PW, [new TableRow({ children: [
  cell([panelHead('Balance receivable by buyer — US$', `${buyers.length} buyers`, PW[0] - 2 * PAD.x), chart, tiny()],
    { w: PW[0], valign: VerticalAlign.TOP, py: 90 }),
  gapCell(GAP),
  cell([panelHead('Advances received'), advTable,
        para(run(advNote, { size: FS.micro, color: MUTED }), { before: 90 })],
    { w: PW[2], valign: VerticalAlign.TOP, py: 90 }),
] })]);

// ---------------------------------------------------------------- bills table
const BW = [540, 1615, 1077, 1723, 1508, 1831, 1185, 1293];
const R = AlignmentType.RIGHT;
const th = (t, i) => cell(
  t.split('\n').map((line, j, a) => para(run(line.toUpperCase(), { size: FS.micro, bold: true, color: WHITE, ls: 24 }),
    { align: [0, 4, 5, 6].includes(i) ? R : undefined, after: j < a.length - 1 ? 0 : 0 })),
  { w: BW[i], fill: DEEP, valign: VerticalAlign.BOTTOM, py: 60,
    borders: { top: { ...hair, color: DEEP }, bottom: { ...hair, color: DEEP },
               left: { ...hair, color: DEEP }, right: { style: BorderStyle.SINGLE, size: 4, color: '3B7F95' } } });
const header = new TableRow({ tableHeader: true, cantSplit: true, children: [
  'No', 'Invoice no', 'Item', 'Buyer', 'Amount\n(US$)', 'Amount\n(₹)', 'Advance\n(US$)', 'Note',
].map(th) });

const td = (child, i, fill) => cell(child, { w: BW[i], fill, py: 22 });
const body = BILLS.map((r, i) => {
  const fill = i % 2 ? WASH : undefined;
  return new TableRow({ cantSplit: true, children: [
    td(para(run(String(i + 1), { size: FS.small, color: MUTED }), { align: R }), 0, fill),
    td(para(run(r[0], { size: FS.small })), 1, fill),
    td(para(run(r[1], { size: FS.small })), 2, fill),
    td(para(run(r[2], { size: FS.small })), 3, fill),
    td(para(mono(usd(r[3]), { size: FS.small }), { align: R }), 4, fill),
    td(para(mono(inr(r[3] * FX), { size: FS.small }), { align: R }), 5, fill),
    td(para(mono(r[4] ? usd(r[4]) : '', { size: FS.small, bold: true }), { align: R }), 6, fill),
    td(para(run(r[5], { size: FS.small })), 7, fill),
  ] });
});

const footRow = (lbl, a, b, c, o = {}) => {
  const col = o.net ? WHITE : o.muted ? MUTED : INK;
  const bd = o.net ? boxNone
    : o.gross ? { top: rule(14), bottom: rule(8), left: none, right: none }
    : { top: none, bottom: rule(8), left: none, right: none };
  const f = o.net ? DEEP : undefined;
  const sz = FS.small;
  const c2 = (child, i, span) => cell(child, { w: span ? BW.slice(i, i + span).reduce((x, y) => x + y, 0) : BW[i],
    span, fill: f, borders: bd, py: o.net ? 75 : 50 });
  return new TableRow({ cantSplit: true, children: [
    c2(para(run('')), 0),
    c2(para(run(lbl.toUpperCase(), { size: FS.micro, bold: true, color: o.net ? WHITE : MUTED, ls: 24 })), 1, 3),
    c2(para(mono(a, { size: sz, bold: !o.muted, color: col }), { align: R }), 4),
    c2(para(mono(b, { size: sz, bold: !o.muted, color: col }), { align: R }), 5),
    c2(para(mono(c, { size: sz, bold: !o.muted, color: col }), { align: R }), 6),
    c2(para(run('')), 7),
  ] });
};
const billsTable = table(BW, [
  header, ...body,
  footRow(`Total billed — ${T.n} ${T.n === 1 ? 'bill' : 'bills'}`, usd(T.usd), inr(T.usd * FX),
          T.adv ? usd(T.adv) : '—', { gross: true }),
  footRow('Less: advances received', T.adv ? '− ' + usd(T.adv) : '—', T.adv ? '− ' + inr(T.adv * FX) : '—', '',
          { muted: true }),
  footRow('Balance receivable', usd(BAL), inr(BAL * FX), '', { net: true }),
]);

// ---------------------------------------------------------------- foot
const foot = new Paragraph({
  children: [
    run(`Prepared from the bills-on-hand list dated ${LIST_DATE}. Rupee figures at ₹ ${FX.toFixed(2)} per US$.`,
      { size: FS.fine, color: MUTED }),
    new TextRun({ children: [new Tab()] }),
    mono(`As at ${AS_AT}`, { size: FS.fine, color: MUTED }),
  ],
  tabStops: [{ type: TabStopType.RIGHT, position: W }],
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 4 } },
  spacing: { before: 160 },
});

// ---------------------------------------------------------------- document
const doc = new Document({
  creator: 'Premier Exports International',
  title: `Bills on hand — ${AS_AT}`,
  description: 'Export bills on hand: amounts billed, advances received and balance receivable.',
  styles: { default: { document: { run: { font: SANS, size: FS.body, color: INK } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 },
      margin: { top: 567, bottom: 567, left: 567, right: 567, header: 300, footer: 300 } } },
    children: [
      letterhead, brandRule, titleBand, spacer(150), tiles, spacer(150),
      panels, spacer(150), billsTable, foot,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, buf);
  console.log(`wrote ${OUT}: ${T.n} bills, US$ ${usd(T.usd)} billed, US$ ${usd(T.adv)} advances, ` +
              `US$ ${usd(BAL)} balance, ₹ at ${FX.toFixed(2)}`);
});
