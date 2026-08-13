// Builds PEI-Purchase-Order.docx — the Word counterpart of
// purchase-order/index.html, sharing its palette, type scale and terms.
//
//   npm install docx
//   node tools/build-docx.js purchase-order/PEI-Purchase-Order.docx
//
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign,
} = require('docx');

// ---------------------------------------------------------------- tokens
// Same design system as purchase-order/index.html: one type scale, one
// spacing scale, the brand cyan sampled from the supplied logo artwork.
const BRAND = '00A2D3', DEEP = '065E7A', INK = '0B2C3A', MUTED = '5E7C8B';
const WASH = 'F2F8FB', BAND = 'E7F3F9', LINE = 'CBDFE8', WHITE = 'FFFFFF';

const FS = { micro: 14, fine: 15, small: 17, body: 18, lead: 23, title: 39, mark: 37 };
const SANS = 'Arial', SERIF = 'Georgia', MONO = 'Consolas';

const SP = { gap: 180, tight: 90 };          // twips between blocks
const PAD = { x: 120, y1: 40, y2: 80 };      // cell padding, twips

const hair = { style: BorderStyle.SINGLE, size: 6, color: LINE };
const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const boxAll = { top: hair, bottom: hair, left: hair, right: hair };
const boxNone = { top: none, bottom: none, left: none, right: none };

// ---------------------------------------------------------------- helpers
const run = (text, o = {}) => new TextRun({
  text, font: o.font || SANS, size: o.size || FS.body,
  color: o.color || INK, bold: o.bold, italics: o.italics,
  characterSpacing: o.ls,
});

const para = (text, o = {}) => new Paragraph({
  children: Array.isArray(text) ? text : [run(text, o)],
  alignment: o.align, spacing: o.spacing,
});

// an uppercase micro label — the one tracked style used everywhere
const label = (text, o = {}) => new Paragraph({
  children: [run(text.toUpperCase(), {
    size: FS.micro, bold: true, color: o.color || MUTED, ls: 24,
  })],
  alignment: o.align,
  spacing: { after: o.after === undefined ? 20 : o.after },
});

const cell = (children, o = {}) => new TableCell({
  children: Array.isArray(children) ? children : [children],
  width: { size: o.w, type: WidthType.DXA },
  columnSpan: o.span,
  borders: o.borders || boxAll,
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  margins: { top: o.py === undefined ? PAD.y1 : o.py, bottom: o.py === undefined ? PAD.y1 : o.py,
             left: PAD.x, right: PAD.x },
  verticalAlign: o.valign || VerticalAlign.TOP,
});

const table = (widths, rows) => new Table({
  columnWidths: widths,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  rows,
  borders: boxNone,
});

const spacer = (h) => new Paragraph({ children: [], spacing: { before: 0, after: h } });

// blank line the user types onto, inside a bordered cell
const blank = (o = {}) => para('', { size: o.size || FS.body });

// ---------------------------------------------------------------- letterhead
const logo = fs.readFileSync(path.join(__dirname, '..', 'purchase-order', 'assets', 'pei-mark.png'));

const LH = [1000, 4100, 5672];
const letterhead = table(LH, [
  new TableRow({
    children: [
      cell(new Paragraph({
        children: [new ImageRun({ data: logo, type: 'png', transformation: { width: 62, height: 50 } })],
      }), { w: LH[0], borders: boxNone, py: 0 }),
      cell([
        new Paragraph({
          children: [run('Premier Exports', { size: FS.mark, color: BRAND })],
          spacing: { after: 0, line: 420, lineRule: 'exact' },
        }),
        new Paragraph({
          children: [run('International', { size: FS.mark, color: BRAND })],
          spacing: { before: 0, after: 0, line: 420, lineRule: 'exact' },
        }),
      ], { w: LH[1], borders: boxNone, py: 0, valign: VerticalAlign.CENTER }),
      cell([
        para('AP X/453, NH-66 Highway, Chandiroor P.O.,', { size: FS.small, color: MUTED, align: AlignmentType.RIGHT }),
        para('Aroor, Alappuzha, Kerala - 688 537, India', { size: FS.small, color: MUTED, align: AlignmentType.RIGHT }),
        para('Mob: +91 81290 99866  ·  premier.pei@gmail.com', { size: FS.small, color: MUTED, align: AlignmentType.RIGHT }),
        new Paragraph({
          children: [run('GSTIN 32AADFP3158P1ZZ', { size: FS.small, font: MONO })],
          alignment: AlignmentType.RIGHT,
        }),
      ], { w: LH[2], borders: boxNone, py: 0 }),
    ],
  }),
]);

const brandRule = new Paragraph({
  children: [],
  border: { bottom: { style: BorderStyle.SINGLE, size: 20, color: BRAND } },
  spacing: { before: 100, after: SP.gap },
});

// ---------------------------------------------------------------- title band
const TB = [6032, 1418, 3322];
// rowSpan lives on TableCell, so build the first row explicitly
const titleCell = new TableCell({
  children: [
    new Paragraph({ children: [run('PURCHASE ORDER', { size: FS.title, bold: true, ls: 30 })], spacing: { after: 40 } }),
    label('Original for supplier', { after: 0 }),
  ],
  width: { size: TB[0], type: WidthType.DXA },
  rowSpan: 3, borders: boxNone, verticalAlign: VerticalAlign.BOTTOM,
  margins: { top: 0, bottom: PAD.y2, left: 0, right: PAD.x },
});
const kv = (k, mono) => [
  cell(label(k, { after: 0 }), { w: TB[1], fill: WASH, valign: VerticalAlign.CENTER }),
  cell(blank({ size: mono ? FS.body : FS.small }), { w: TB[2], valign: VerticalAlign.CENTER }),
];
const titleBand = table(TB, [
  new TableRow({ children: [titleCell, ...kv('P.O. No.')] }),
  new TableRow({ children: kv('P.O. Date') }),
  new TableRow({ children: kv('Your quotation') }),
]);

// ---------------------------------------------------------------- parties
const PT = [5136, 500, 5136];
const partyHead = (t) => cell(label(t, { color: DEEP, after: 0 }), { w: PT[0], fill: BAND, py: PAD.y1 });
const parties = table(PT, [
  new TableRow({
    children: [
      partyHead('Supplier'),
      cell(para(''), { w: PT[1], borders: boxNone }),
      partyHead('Deliver to'),
    ],
  }),
  new TableRow({
    children: [
      cell([blank(), blank(), blank(), blank()], { w: PT[0], py: PAD.y2 }),
      cell(para(''), { w: PT[1], borders: boxNone }),
      cell([
        para('Premier Exports International', { bold: true }),
        para('AP X/453, NH-66 Highway, Chandiroor P.O.,'),
        para('Aroor, Alappuzha, Kerala - 688 537, India'),
        para([run('GSTIN 32AADFP3158P1ZZ', { font: MONO })]),
        blank(),
      ], { w: PT[2], py: PAD.y2 }),
    ],
  }),
]);

// ---------------------------------------------------------------- order strip
const ST = [2693, 2693, 2693, 2693];
const stripCell = (k, v, i) => cell([label(k), para(v || '')], { w: ST[i], py: PAD.y2 });
const strip = table(ST, [
  new TableRow({
    children: [
      stripCell('Place of supply', 'Kerala (32)', 0),
      stripCell('Payment terms', '', 1),
      stripCell('Delivery on or before', '', 2),
      stripCell('Dispatch mode', '', 3),
    ],
  }),
]);

// ---------------------------------------------------------------- items
const IT = [646, 3662, 1293, 754, 1185, 1508, 1724];
const th = (t, i, align) => cell(
  new Paragraph({
    children: [run(t.toUpperCase(), { size: FS.micro, bold: true, color: WHITE, ls: 24 })],
    alignment: align,
  }),
  { w: IT[i], fill: DEEP, py: PAD.y2,
    borders: { top: hair, bottom: hair, left: hair, right: { style: BorderStyle.SINGLE, size: 6, color: DEEP } } }
);
const itemRows = [
  new TableRow({
    tableHeader: true,
    children: [
      th('Sl', 0, AlignmentType.CENTER), th('Description of goods / services', 1),
      th('HSN / SAC', 2), th('Qty', 3, AlignmentType.RIGHT), th('Unit', 4),
      th('Rate (₹)', 5, AlignmentType.RIGHT), th('Amount (₹)', 6, AlignmentType.RIGHT),
    ],
  }),
];
for (let i = 1; i <= 5; i++) {
  const shade = i % 2 === 0 ? WASH : undefined;
  itemRows.push(new TableRow({
    children: [
      cell(para(String(i), { font: MONO, size: FS.small, color: MUTED, align: AlignmentType.CENTER }), { w: IT[0], fill: shade, py: PAD.y2 }),
      cell(blank(), { w: IT[1], fill: shade, py: PAD.y2 }),
      cell(blank(), { w: IT[2], fill: shade, py: PAD.y2 }),
      cell(blank(), { w: IT[3], fill: shade, py: PAD.y2 }),
      cell(blank(), { w: IT[4], fill: shade, py: PAD.y2 }),
      cell(blank(), { w: IT[5], fill: shade, py: PAD.y2 }),
      cell(blank(), { w: IT[6], fill: shade, py: PAD.y2 }),
    ],
  }));
}
const items = table(IT, itemRows);

// ---------------------------------------------------------------- foot
const TOT = [2844, 1896];
const totRow = (k, opts = {}) => new TableRow({
  children: [
    cell(para(k, { size: FS.small, color: opts.grand ? WHITE : MUTED, bold: opts.bold || opts.grand,
                   ls: opts.grand ? 24 : 0 }),
         { w: TOT[0], fill: opts.grand ? DEEP : (opts.tax ? WASH : undefined), py: opts.grand ? 70 : PAD.y1 }),
    cell(para(opts.value || '', { font: MONO, size: opts.grand ? FS.lead : FS.small,
                                  color: opts.grand ? WHITE : INK, bold: opts.grand,
                                  align: AlignmentType.RIGHT }),
         { w: TOT[1], fill: opts.grand ? DEEP : (opts.tax ? WASH : undefined), py: opts.grand ? 70 : PAD.y1 }),
  ],
});
const totals = table(TOT, [
  totRow('Sub total'), totRow('Less: discount'), totRow('Add: freight & packing'),
  totRow('Taxable value', { bold: true }),
  totRow('CGST @ 9 %', { tax: true }), totRow('SGST @ 9 %', { tax: true }),
  totRow('Round off'), totRow('Grand total', { grand: true }),
]);

const FT = [5732, 300, 4740];
const foot = table(FT, [
  new TableRow({
    children: [
      new TableCell({
        width: { size: FT[0], type: WidthType.DXA }, borders: boxNone,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [
          table([FT[0]], [new TableRow({
            children: [cell([
              label('Amount chargeable in words'),
              para('', { font: SERIF, italics: true }),
            ], { w: FT[0], fill: WASH, py: PAD.y2 })],
          })]),
          spacer(SP.gap),
          table([FT[0]], [
            new TableRow({ children: [cell(label('Notes & special instructions', { color: DEEP, after: 0 }), { w: FT[0], fill: BAND, py: PAD.y1 })] }),
            new TableRow({ children: [cell([blank(), blank()], { w: FT[0], py: PAD.y2 })] }),
          ]),
        ],
      }),
      cell(para(''), { w: FT[1], borders: boxNone }),
      new TableCell({
        width: { size: FT[2], type: WidthType.DXA }, borders: boxNone,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [totals],
      }),
    ],
  }),
]);

// ---------------------------------------------------------------- terms
const TERMS = [
  'Quote this order number and date on all invoices, challans, packing lists and correspondence.',
  'Supply strictly to the make, model and specification stated. Substitutions need our prior written approval.',
  'Rates are firm and include packing. GST extra as shown; freight at actuals where indicated.',
  'A GST-compliant tax invoice quoting our GSTIN is required. Invoices without HSN / SAC codes will be returned.',
  'Deliver on or before the date stated. Intimate any delay in writing; we may cancel the balance without liability.',
  'Material is accepted subject to inspection at our works. Rejections are returnable at your cost.',
  'Warranty twelve months from commissioning, unless your quotation offers longer.',
  'Test weights, labour and lifting equipment for on-site calibration will be arranged by us where agreed in advance.',
  'Payment as per the terms above, against material received in good condition and a correct invoice.',
  'Disputes are subject to the jurisdiction of the courts at Alappuzha, Kerala.',
];
const termPara = (n) => new Paragraph({
  children: [run(`${n + 1}.  ${TERMS[n]}`, { font: SERIF, size: FS.fine })],
  spacing: { after: 60, line: 230, lineRule: 'auto' },
  indent: { left: 220, hanging: 220 },
});
const termsHead = new Paragraph({
  children: [run('TERMS & CONDITIONS', { size: FS.micro, bold: true, color: DEEP, ls: 24 })],
  border: { bottom: hair },
  spacing: { before: SP.gap, after: 100 },
});
const TM = [5136, 500, 5136];
const terms = table(TM, [
  new TableRow({
    children: [
      cell([0, 1, 2, 3, 4].map(termPara), { w: TM[0], borders: boxNone, py: 0 }),
      cell(para(''), { w: TM[1], borders: boxNone }),
      cell([5, 6, 7, 8, 9].map(termPara), { w: TM[2], borders: boxNone, py: 0 }),
    ],
  }),
]);

// ---------------------------------------------------------------- signature
const SG = [4886, 1000, 4886];
const sigBox = (title, line, w) => cell([
  para(title, { size: FS.small, bold: true }),
  spacer(420),
  new Paragraph({
    children: [run(line.toUpperCase(), { size: FS.micro, bold: true, color: MUTED, ls: 24 })],
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: INK } },
    spacing: { before: 0, after: 0 },
  }),
], { w, borders: boxNone, py: 0 });

const signature = table(SG, [
  new TableRow({
    children: [
      sigBox("Supplier's acknowledgement", 'Signature, name & date', SG[0]),
      cell(para(''), { w: SG[1], borders: boxNone }),
      sigBox('For Premier Exports International', 'Authorised signatory', SG[2]),
    ],
  }),
]);

const footNote = new Paragraph({
  children: [run('Please return one signed copy as confirmation of acceptance.', { size: FS.fine, color: MUTED })],
  border: { top: hair },
  spacing: { before: SP.gap, after: 0 },
});

// ---------------------------------------------------------------- document
const doc = new Document({
  creator: 'Premier Exports International',
  title: 'Purchase Order',
  description: 'Purchase order form',
  styles: { default: { document: { run: { font: SANS, size: FS.body, color: INK } } } },
  sections: [{
    properties: { page: { margin: { top: 567, right: 567, bottom: 567, left: 567 } } },
    children: [
      letterhead, brandRule,
      titleBand, spacer(SP.gap),
      parties, spacer(SP.gap),
      strip, spacer(SP.gap),
      items, spacer(SP.gap),
      foot,
      termsHead, terms,
      spacer(SP.gap), signature,
      footNote,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2], buf);
  console.log('wrote', process.argv[2], (buf.length / 1024).toFixed(1) + ' KB');
});
