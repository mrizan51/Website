#!/usr/bin/env python3
"""Builds PEI-Stock-Register.xlsx — a daily stock register driven by formulas,
with a print-ready 'Daily Report' sheet for the directors.

    python tools/build-stock-xlsx.py out.xlsx                 # from the baseline below
    python tools/build-stock-xlsx.py out.xlsx --from live.xlsx  # keep a live file's data

Design note. The paper statement re-keys every opening balance each morning
from yesterday's total, which is both tedious and the easiest place to make a
mistake. Here the opening balance is a FROZEN baseline and every later movement
is a dated row on 'Daily Entry'. Closing = baseline + production to date −
shipment to date, so the register rolls itself forward and nothing is re-typed.

--from reads products, headings, the movement log and any rates out of an
existing workbook, so the report sheet can be added to a file already in use
without losing a day's work.
"""
import sys, os, datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.drawing.image import Image as XLImage
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO = os.path.join(ROOT, 'stock-statement', 'assets', 'pei-mark.png')

args = [a for a in sys.argv[1:] if not a.startswith('--')]
OUT = args[0] if args else 'PEI-Stock-Register.xlsx'
SRC = None
if '--from' in sys.argv:
    SRC = sys.argv[sys.argv.index('--from') + 1]
# Excel sheet protection is a guardrail against accidents, not security —
# the password is trivially removable by anyone determined. It exists to stop
# staff overwriting a formula or restyling the sheet by mistake.
PASSWORD = (sys.argv[sys.argv.index('--password') + 1]
            if '--password' in sys.argv else 'PEI2026')
UNLOCK = Protection(locked=False)

ENTRY_FIRST, ENTRY_MAX, ENTRY_FMT = 7, 5000, 250
MOVE_ROWS = 20                      # movement lines the daily report prints
# Spare capacity so a new product or brand is three cells of typing rather than
# surgery on fixed ranges. Blank spares contribute nothing and print blank.
SPARE_PRODUCTS = 20
SPARE_BRANDS = 3

# LibreOffice drops <workbookProtection> attributes when it recalculates, and
# re-saving through openpyxl would strip every cached formula value. So the
# structure lock is re-applied by patching the XML in place, after recalc:
#     python tools/build-stock-xlsx.py --relock <file.xlsx>
if '--relock' in sys.argv:
    import zipfile, shutil, tempfile, re as _re
    from openpyxl.utils.protection import hash_password
    target = sys.argv[sys.argv.index('--relock') + 1]
    zin = zipfile.ZipFile(target)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    blob = dict((i.filename, d) for i, d in items)

    # 1. workbook structure lock
    el = f'<workbookProtection workbookPassword="{hash_password(PASSWORD)}" lockStructure="1"/>'
    wbx = blob['xl/workbook.xml'].decode('utf-8')
    if _re.search(r'<workbookProtection[^>]*/>', wbx):
        wbx = _re.sub(r'<workbookProtection[^>]*/>', el, wbx, count=1)
    elif '<workbookProtection' in wbx:
        wbx = _re.sub(r'<workbookProtection.*?</workbookProtection>', el, wbx, count=1, flags=_re.S)
    else:
        for a in ('<bookViews', '<sheets'):
            if a in wbx:
                wbx = wbx.replace(a, el + a, 1)
                break

    # 2. Daily Entry's unlocked columns. LibreOffice drops <protection> from a
    #    column's style, which would re-lock every row past the ruled block.
    #    Append fresh xfs rather than editing shared ones, so no other cell
    #    silently becomes writable, then repoint the <col> elements at them.
    rels = blob['xl/_rels/workbook.xml.rels'].decode('utf-8')
    rid = dict(_re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    sheet_file = None
    for m in _re.finditer(r'<sheet\b[^>]*/>', wbx):
        a = dict(_re.findall(r'(?:r:)?(\w+)="([^"]*)"', m.group(0)))
        if a.get('name') == 'Daily Entry':
            t = rid.get(a.get('id', ''), '')
            sheet_file = 'xl/' + t.lstrip('/').replace('worksheets/', 'worksheets/')
            break
    if sheet_file and sheet_file in blob:
        sx = blob[sheet_file].decode('utf-8')
        st_xml = blob['xl/styles.xml'].decode('utf-8')
        head, rest = st_xml.split('<cellXfs', 1)
        attrs, body_and_tail = rest.split('>', 1)
        body, tail = body_and_tail.split('</cellXfs>', 1)
        xfs = _re.findall(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', body, _re.S)
        added = []
        def unlocked_copy(xf):
            if '<protection' in xf:
                # LibreOffice writes an explicit locked="1"; flip it rather
                # than assuming the presence of the element means unlocked
                xf = _re.sub(r'<protection\b[^>]*/>',
                             '<protection locked="0" hidden="0"/>', xf, count=1)
                open_tag = xf.split('>', 1)[0]
                if 'applyProtection' not in open_tag:
                    xf = xf.replace(open_tag, open_tag + ' applyProtection="1"', 1)
                return xf
            if xf.endswith('/>'):
                base = xf[:-2]
                if 'applyProtection' not in base:
                    base += ' applyProtection="1"'
                return base + '><protection locked="0" hidden="0"/></xf>'
            base = xf[:-len('</xf>')]
            if 'applyProtection' not in base.split('>', 1)[0]:
                open_tag, inner = base.split('>', 1)
                base = open_tag + ' applyProtection="1">' + inner
            return base + '<protection locked="0" hidden="0"/></xf>'
        def repoint(m):
            c = m.group(0)
            a = dict(_re.findall(r'(\w+)="([^"]*)"', c))
            lo, hi = int(a.get('min', 0)), int(a.get('max', 0))
            if lo > 5 or hi < 1:
                return c
            sid = int(a.get('style', 0))
            new_xf = unlocked_copy(xfs[sid])
            new_id = len(xfs) + len(added)
            added.append(new_xf)
            if 'style=' in c:
                return _re.sub(r'style="\d+"', f'style="{new_id}"', c)
            return c[:-2] + f' style="{new_id}"/>'
        sx = _re.sub(r'<col\b[^>]*/>', repoint, sx)
        # LibreOffice also clips data-validation ranges back to the used rows,
        # which would leave the dropdown and the date check absent on every
        # later row. Stretch them back over the whole entry range.
        def stretch(m):
            return f'sqref="{m.group(1)}{ENTRY_FIRST}:{m.group(2)}{ENTRY_MAX}"'
        sx = _re.sub(rf'sqref="([A-Z]+){ENTRY_FIRST}:([A-Z]+)(\d+)"', stretch, sx)
        if added:
            body += ''.join(added)
            attrs = _re.sub(r'count="\d+"', f'count="{len(xfs) + len(added)}"', attrs)
            blob['xl/styles.xml'] = (head + '<cellXfs' + attrs + '>' + body +
                                     '</cellXfs>' + tail).encode('utf-8')
            blob[sheet_file] = sx.encode('utf-8')

    blob['xl/workbook.xml'] = wbx.encode('utf-8')
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            zout.writestr(info, blob[info.filename])
    tmp.close()
    shutil.move(tmp.name, target)
    os.chmod(target, 0o644)
    print(f'relocked {target}: structure locked, {len(added) if sheet_file else 0} '
          f'entry columns re-unlocked')
    sys.exit(0)


# ---------------------------------------------------------------- baseline
BASE_DATE = datetime.date(2026, 8, 20)
STOCK = [
    ('Blue Dolphin', 'PD PVN 100/200', 11459), ('Blue Dolphin', 'PD KKD 100/200', 517),
    ('Blue Dolphin', 'PUD PVN 200/300', 2262), ('Blue Dolphin', 'PUD KKD 200/300', 2835),
    ('Blue Dolphin', 'PUD PVN 300/500', 1673), ('Blue Dolphin', 'PUD KKD 300/500', 7542),
    ('Blue Dolphin', 'PD PVN 200/300', 2357),
    ('Primus Japan', 'PUD KZN 40/60', 64), ('Primus Japan', 'PUD KZN 60/80', 191),
    ('Primus Japan', 'PUD PVN 80/120', 16849), ('Primus Japan', 'PUD KZN 80/120', 466),
    ('Primus Japan', 'PUD NRN 80/120', 69), ('Primus Japan', 'PUD PVN 100/200', 2617),
    ('Primus Japan', 'PUD NRN 100/200', 13), ('Primus Japan', 'PUD KZN 100/200', 4209),
    ('Primus Japan', 'PUD PVN 200/300', 0), ('Primus Japan', 'PUD KKD 200/300', 2757),
    ('Primus Japan', 'PUD KZN 200/300', 1080), ('Primus Japan', 'PUD PVN 300/500', 421),
    ('Primus Japan', 'PUD KKD 300/500', 580),
    ('Primus EU', 'PUD 20/40', 185), ('Primus EU', 'PUD 40/60', 368),
    ('Primus EU', 'PUD 60/80', 380), ('Primus EU', 'PUD 80/120', 1151),
    ('Primus EU', 'PUD 100/200', 3797), ('Primus EU', 'PUD 200/300', 4118),
    ('Primus EU', 'PUD 300/500', 1490), ('Primus EU', 'BKN', 15130),
    ('AMF Brand', 'PD PVN 100/200', 0), ('AMF Brand', 'PD KKD 100/200', 0),
    ('AMF Brand', 'PUD PVN 200/300', 1516), ('AMF Brand', 'PUD KKD 200/300', 0),
    ('AMF Brand', 'PUD PVN 300/500', 18), ('AMF Brand', 'PUD KKD 300/500', 0),
    ('THT Brand', 'PUD PVN 80/120', 1389), ('THT Brand', 'PUD PVN 100/200', 7055),
    ('THT Brand', 'PUD PVN 200/300', 4682),
    ('Deepsea', '300/500', 670),
    ('China', 'PUD PVN 300/500', 451), ('China', 'PUD KKD 300/500', 188),
    ('China', 'PUD PVN 500/800', 401), ('China', 'PUD KKD 500/800', 1920),
]
LOG = [(BASE_DATE, 'Blue Dolphin - PD PVN 100/200', 1068, 0, "Day's production per statement 20/08/26"),
       (BASE_DATE, 'Blue Dolphin - PD KKD 100/200', 65, 0, ''),
       (BASE_DATE, 'Blue Dolphin - PUD PVN 200/300', 267, 0, ''),
       (BASE_DATE, 'Primus Japan - PUD KZN 40/60', 29, 0, ''),
       (BASE_DATE, 'Primus Japan - PUD KZN 60/80', 39, 0, ''),
       (BASE_DATE, 'Primus Japan - PUD KZN 80/120', 50, 0, ''),
       (BASE_DATE, 'Primus Japan - PUD NRN 80/120', 3, 0, ''),
       (BASE_DATE, 'Primus Japan - PUD KZN 100/200', 50, 0, ''),
       (BASE_DATE, 'Primus Japan - PUD KZN 200/300', 33, 0, ''),
       (BASE_DATE, 'Primus Japan - PUD PVN 300/500', 156, 0, ''),
       (BASE_DATE, 'Primus Japan - PUD KKD 300/500', 596, 0, ''),
       (BASE_DATE, 'Primus EU - BKN', 52, 0, ''),
       (BASE_DATE, 'THT Brand - PUD PVN 80/120', 1872, 0, ''),
       (BASE_DATE, 'THT Brand - PUD PVN 100/200', 600, 0, '')]
RATES = {}
EXTRA_BRANDS = ['Shrimp Whole 5x3kgs']       # tracked, no stock held
STMT_DATE = BASE_DATE
FX = None                                    # US$ -> INR, set daily by the user

# ---------------------------------------------------------------- --from
def as_date(v):
    """Log dates must be real dates: a text date never matches a date criterion
    in SUMIFS, which silently zeroes the two 'today' columns."""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y'):
        try:
            return datetime.datetime.strptime(str(v).strip(), fmt).date()
        except (ValueError, TypeError):
            pass
    return None

if SRC:
    src = load_workbook(SRC)
    s_st, s_de, s_sm = src['Stock'], src['Daily Entry'], src['Summary']
    STOCK, RATES = [], {}
    r = 8
    while s_st[f'A{r}'].value and s_st[f'A{r}'].value != 'GRAND TOTAL':
        a, b, c = s_st[f'A{r}'].value, s_st[f'B{r}'].value, s_st[f'C{r}'].value
        STOCK.append((a, b, c if isinstance(c, (int, float)) else 0))
        rate = s_st[f'I{r}'].value
        if isinstance(rate, (int, float)) and rate:
            RATES[f'{a} - {b}'] = rate
        r += 1
    LOG, skipped = [], 0
    for r in range(ENTRY_FIRST, s_de.max_row + 1):
        item = s_de[f'B{r}'].value
        if not item:
            continue
        d = as_date(s_de[f'A{r}'].value)
        if d is None:
            skipped += 1
            continue
        LOG.append((d, item, s_de[f'C{r}'].value or 0, s_de[f'D{r}'].value or 0,
                    s_de[f'E{r}'].value or ''))
    # Brands listed on Summary but holding no stock. Stop at GRAND TOTAL:
    # the explanatory footnote sits below it and is not a brand.
    EXTRA_BRANDS = []
    seen = {b for b, _, _ in STOCK}
    stop = s_sm.max_row + 1
    for rr in range(9, s_sm.max_row + 1):
        if s_sm[f'A{rr}'].value == 'GRAND TOTAL':
            stop = rr
            break
    for rr in range(9, stop):
        v = s_sm[f'A{rr}'].value
        if v and v not in seen:
            EXTRA_BRANDS.append(v)
            seen.add(v)
    STMT_DATE = as_date(s_st['C3'].value) or BASE_DATE
    _fx = s_st['C4'].value
    FX = _fx if isinstance(_fx, (int, float)) and _fx else None
    print(f'read {SRC}: {len(STOCK)} products, {len(LOG)} log rows, '
          f'{len(RATES)} rates' + (f', {skipped} rows skipped (unreadable date)' if skipped else ''))

BRANDS = []
for b, _, _ in STOCK:
    if b not in BRANDS:
        BRANDS.append(b)
BRANDS += [b for b in EXTRA_BRANDS if b not in BRANDS]

# ---------------------------------------------------------------- style
FONT = 'Arial'
DEEP, BAND, WASH, LINE = '065E7A', 'E7F3F9', 'F2F8FB', 'CBDFE8'
INK, MUTED, BRAND_C = '0B2C3A', '5E7C8B', '00A2D3'
BLUE_IN, YELLOW = '0000FF', 'FFFF00'

thin = Side(style='thin', color=LINE)
box = Border(left=thin, right=thin, top=thin, bottom=thin)
under = Border(bottom=thin)

N_IND = r'[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;#,##0'
N_DASH = r'#,##0;-#,##0;"–"'
CUR_IND = r'[>=10000000]"₹" ##\,##\,##\,##0.00;[>=100000]"₹" ##\,##\,##0.00;"₹" #,##0.00'
RATE_F = r'"$" #,##0.00'          # rate per slab is an export price, in USD
FX_F = r'"₹" #,##0.00'
DATE_F = 'DD/MM/YYYY'

def head(ws, row, cols, widths=None, align_from=3):
    for i, t in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=DEEP)
        c.alignment = Alignment(horizontal='right' if i >= align_from else 'left',
                                vertical='bottom', wrap_text=True)
        c.border = box
    ws.row_dimensions[row].height = 30
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

def title_block(ws, subtitle):
    ws['A1'] = 'PREMIER EXPORTS INTERNATIONAL'
    ws['A1'].font = Font(name=FONT, size=14, bold=True, color=DEEP)
    ws['A2'] = subtitle
    ws['A2'].font = Font(name=FONT, size=10, bold=True, color=MUTED)

wb = Workbook()

# ================================================================ Read me
rm = wb.active
rm.title = 'Read me'
title_block(rm, 'STOCK REGISTER — HOW IT WORKS')
rm.column_dimensions['A'].width = 2
rm.column_dimensions['B'].width = 108
lines = [
    ('h', 'Every day, in one minute'),
    ('p', "1.  Open 'Daily Entry' and add one row per product that moved."),
    ('p', '2.  Type the date, pick the product from the dropdown, enter slabs produced and/or shipped.'),
    ('p', "3.  On 'Stock', set the Statement date and the US$ → ₹ exchange rate for today."),
    ('p', "4.  Open 'Daily Report' and save it as PDF — that is the directors' copy."),
    ('n', ''),
    ('h', 'Sending the report'),
    ('p', "'Daily Report' is already sized to one A4 page. In Excel: File ▸ Export ▸ Create PDF/XPS"),
    ('p', '(or File ▸ Save a Copy and choose PDF), pick "Selected sheet", and send that file.'),
    ('p', 'Nothing else needs installing — Excel writes the PDF itself.'),
    ('n', ''),
    ('h', 'Why you never re-type an opening balance'),
    ('p', 'Opening balance is a fixed baseline, not a daily figure.'),
    ('p', 'Closing balance = opening balance + production to date − shipment to date,'),
    ('p', "counted from every row on 'Daily Entry'. The register rolls forward on its own,"),
    ('p', "so yesterday's closing is today's starting point without anyone copying a number."),
    ('n', ''),
    ('h', 'What can be typed into — everything else is locked'),
    ('b', "'Daily Entry' — the whole grid: date, product, production, shipment, note."),
    ('y', "'Stock' — the Statement date, the US$ → ₹ exchange rate, and the Rate / Slab column."),
    ('k', 'Every other cell is protected: headings, opening balances, formulas, totals,'),
    ('k', 'the Summary and the Daily Report. Formatting is locked too, so column widths,'),
    ('k', 'fonts, colours and number formats cannot be changed by accident.'),
    ('p', 'The owner holds the password. To change a locked cell: Review ▸ Unprotect Sheet,'),
    ('p', 'make the change, then Review ▸ Protect Sheet again to put the guard back.'),
    ('n', ''),
    ('h', 'How the stock is valued'),
    ('p', 'Rate / Slab is the export price, in US dollars — that is what the column holds.'),
    ('p', 'Stock Value (₹) = Closing Balance × Rate / Slab (US$) × the US$ → ₹ rate in C4.'),
    ('p', 'The exchange rate lives in that one cell, so changing it revalues everything at once.'),
    ('p', 'Leave it blank and the value column stays blank rather than showing a wrong number;'),
    ('p', 'the Daily Report says "Exchange rate not set" instead of a figure.'),
    ('n', ''),
    ('h', 'Dates must be real dates'),
    ('p', 'Type 21/08/2026, not text. A date stored as text never matches the Statement date,'),
    ('p', "and the two 'today' columns would quietly read zero."),
    ('n', ''),
    ('h', 'Adding a product'),
    ('p', "'Stock' keeps 20 blank rows under the last product. Type the Brand, the Product and"),
    ('p', 'the Opening Balance into the first blank row — nothing else, and no unprotecting.'),
    ('p', 'Every formula on that row is already in place, and the row stays blank until you'),
    ('p', "fill it. The 'Daily Entry' dropdown picks the new product up straight away."),
    ('n', ''),
    ('h', 'Adding a brand'),
    ('p', "Do the same on 'Stock', then type the brand name into a blank row on 'Summary'."),
    ('p', 'Until you do, a red line on Summary and on the Daily Report will tell you the'),
    ('p', 'brand list is incomplete and the totals are understated. Both clear themselves.'),
    ('n', ''),
    ('h', 'When the blank rows run out'),
    ('p', 'Ask for a rebuild — the file is generated, and a new one carries your data across.'),
    ('n', ''),
    ('h', 'Source'),
    ('p', 'Opening balances and the first day of production come from the purchase & processing'),
    ('p', 'statement dated 20/08/26, supplied by Premier Exports International.'),
]
r = 4
for kind, text in lines:
    c = rm.cell(row=r, column=2, value=text)
    if kind == 'h':
        c.font = Font(name=FONT, size=10, bold=True, color=DEEP)
    elif kind == 'b':
        c.font = Font(name=FONT, size=10, color=BLUE_IN, bold=True)
    elif kind == 'y':
        c.font = Font(name=FONT, size=10, bold=True)
        c.fill = PatternFill('solid', fgColor=YELLOW)
    elif kind == 'g':
        c.font = Font(name=FONT, size=10, color=MUTED)
    else:
        c.font = Font(name=FONT, size=10, color=INK)
    r += 1
rm.print_area = f'A1:B{r}'
rm.page_setup.fitToWidth = 1
rm.page_setup.fitToHeight = 0
rm.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Stock
st = wb.create_sheet('Stock')
title_block(st, 'STOCK REGISTER — LIVE POSITION')
st['A3'] = 'Statement date'
st['A3'].font = Font(name=FONT, size=10, bold=True, color=INK)
st['C3'] = STMT_DATE
st['C3'].font = Font(name=FONT, size=10, bold=True, color=BLUE_IN)
st['C3'].fill = PatternFill('solid', fgColor=YELLOW)
st['C3'].border = box
st['C3'].number_format = DATE_F
st['C3'].alignment = Alignment(horizontal='center')
st['C3'].protection = UNLOCK
st['D3'] = "← set to today; drives the two \"today\" columns and the Daily Report"
st['D3'].font = Font(name=FONT, size=9, italic=True, color=MUTED)
st['A4'] = 'US$ → ₹ exchange rate'
st['A4'].font = Font(name=FONT, size=10, bold=True, color=INK)
st['C4'] = FX
st['C4'].font = Font(name=FONT, size=10, bold=True, color=BLUE_IN)
st['C4'].fill = PatternFill('solid', fgColor=YELLOW)
st['C4'].border = box
st['C4'].number_format = FX_F
st['C4'].alignment = Alignment(horizontal='center')
st['C4'].protection = UNLOCK
st['D4'] = "← rupees per US dollar; converts the export rates into the value column"
st['D4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)

st['A5'] = ("Opening balance is a frozen baseline — never re-key it. "
            "Closing follows from the dated rows on 'Daily Entry'.")
st['A5'].font = Font(name=FONT, size=9, italic=True, color=MUTED)

HDR = 7
head(st, HDR,
     ['Brand', 'Product', 'Opening Balance', "Today's Production", "Today's Shipment",
      'Production to Date', 'Shipment to Date', 'Closing Balance', 'Rate / Slab (US$)',
      'Stock Value (₹)', '', 'Key (do not edit)', 'Moved today', 'Rank'],
     [22, 22, 14, 13, 13, 13, 13, 13, 12, 17, 2, 34, 11, 7])
st.cell(row=HDR, column=11).fill = PatternFill('solid', fgColor='FFFFFF')
for col in (12, 13, 14):
    c = st.cell(row=HDR, column=col)
    c.fill = PatternFill('solid', fgColor='D9D9D9')
    c.font = Font(name=FONT, size=9, bold=True, color=MUTED)

st.cell(row=HDR, column=3).comment = Comment(
    'The position this register starts from. A fixed baseline — do not update\n'
    'it daily. Closing Balance is derived from it plus the Daily Entry log.',
    'Premier Exports International', height=80, width=340)

first = HDR + 1
ent = lambda col: f"'Daily Entry'!${col}${ENTRY_FIRST}:${col}${ENTRY_MAX}"
ROWS = list(STOCK) + [(None, None, None)] * SPARE_PRODUCTS
for i, (brand, product, ob) in enumerate(ROWS):
    r = first + i
    spare = brand is None
    st.cell(row=r, column=1, value=brand).font = Font(name=FONT, size=10, color=INK)
    st.cell(row=r, column=2, value=product).font = Font(name=FONT, size=10, color=INK)
    c = st.cell(row=r, column=3, value=ob)
    c.font = Font(name=FONT, size=10, color=BLUE_IN)
    c.number_format = N_IND
    if spare:
        # capacity, not history: these three accept typing without unprotecting
        for col in (1, 2, 3):
            st.cell(row=r, column=col).protection = UNLOCK

    g = lambda f: f'=IF($A{r}="","",{f})'          # blank row stays blank
    st.cell(row=r, column=4, value=g(f"SUMIFS({ent('C')},{ent('B')},$L{r},{ent('A')},$C$3)"))
    st.cell(row=r, column=5, value=g(f"SUMIFS({ent('D')},{ent('B')},$L{r},{ent('A')},$C$3)"))
    st.cell(row=r, column=6, value=g(f"SUMIFS({ent('C')},{ent('B')},$L{r})"))
    st.cell(row=r, column=7, value=g(f"SUMIFS({ent('D')},{ent('B')},$L{r})"))
    st.cell(row=r, column=8, value=g(f'$C{r}+$F{r}-$G{r}'))
    rate = st.cell(row=r, column=9)
    rate.fill = PatternFill('solid', fgColor=YELLOW)
    rate.number_format = RATE_F
    rate.font = Font(name=FONT, size=10, color=BLUE_IN)
    rate.protection = UNLOCK          # periodic input, kept editable
    key = f'{brand} - {product}'
    if key in RATES:
        rate.value = RATES[key]
    # export price (US$) x closing slabs x the one exchange-rate cell
    st.cell(row=r, column=10,
            value=f'=IF(OR($A{r}="",$I{r}="",$C$4=""),"",$H{r}*$I{r}*$C$4)')\
      .number_format = CUR_IND
    for col, val in ((12, f'=IF($A{r}="","",$A{r}&" - "&$B{r})'),
                     (13, f'=IF($A{r}="",0,IF(OR($D{r}<>0,$E{r}<>0),1,0))'),
                     (14, f'=IF($M{r}=1,SUM($M${first}:$M{r}),"")')):
        h = st.cell(row=r, column=col, value=val)
        h.font = Font(name=FONT, size=9, color=MUTED)

    for col in (4, 5):
        st.cell(row=r, column=col).number_format = N_DASH
    for col in (6, 7, 8):
        st.cell(row=r, column=col).number_format = N_IND
    for col in range(1, 11):
        cell = st.cell(row=r, column=col)
        cell.border = box
        if col in (4, 5, 6, 7, 8, 10):
            cell.font = Font(name=FONT, size=10, color=INK, bold=(col == 8))
        if i % 2 and col != 9:
            cell.fill = PatternFill('solid', fgColor=WASH)

last = first + len(ROWS) - 1
tot = last + 1
st.cell(row=tot, column=1, value='GRAND TOTAL')
for col in range(1, 11):
    c = st.cell(row=tot, column=col)
    c.fill = PatternFill('solid', fgColor=DEEP)
    c.border = box
    c.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    if col >= 3 and col != 9:
        L = get_column_letter(col)
        rng = f'{L}{first}:{L}{last}'
        c.value = f'=IF(SUM({rng})=0,"",SUM({rng}))' if col == 10 else f'=SUM({rng})'
        c.number_format = CUR_IND if col == 10 else (N_DASH if col in (4, 5) else N_IND)

st.cell(row=tot + 2, column=1,
        value='Closing Balance = Opening Balance + Production to Date − Shipment to Date        '
              'Stock Value = Closing Balance × Rate per Slab').font = \
    Font(name=FONT, size=9, italic=True, color=MUTED)

st.freeze_panes = f'C{first}'
st.auto_filter.ref = f'A{HDR}:J{last}'
st.conditional_formatting.add(
    f'H{first}:H{last}',
    CellIsRule(operator='lessThan', formula=['0'],
               fill=PatternFill('solid', fgColor='F8D7DA'),
               font=Font(name=FONT, size=10, bold=True, color='B3261E')))
st.print_area = f'A1:J{tot}'
st.print_title_rows = f'{HDR}:{HDR}'
st.page_setup.orientation = 'portrait'
st.page_setup.fitToWidth = 1
st.page_setup.fitToHeight = 0
st.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Daily Entry
de = wb.create_sheet('Daily Entry')
title_block(de, 'DAILY ENTRY — ADD ONE ROW PER PRODUCT THAT MOVED')
de['A4'] = ("After production, add a row below: date, product from the dropdown, slabs produced "
            "and/or shipped. Never delete past rows — they are what the closing position is built from.")
de['A4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)

DH = 6
head(de, DH, ['Date', 'Item (pick from list)', 'Production (slabs)', 'Shipment (slabs)', 'Note'],
     [14, 40, 16, 16, 40], align_from=3)
for col in (1, 2, 5):
    de.cell(row=DH, column=col).alignment = Alignment(horizontal='left', wrap_text=True)

er = DH + 1
for d, item, prod, ship, note in LOG:
    dc = de.cell(row=er, column=1, value=d)
    dc.number_format = DATE_F
    dc.font = Font(name=FONT, size=10, color=BLUE_IN)
    de.cell(row=er, column=2, value=item).font = Font(name=FONT, size=10, color=BLUE_IN)
    for col, v in ((3, prod), (4, ship)):
        c = de.cell(row=er, column=col, value=v)
        c.font = Font(name=FONT, size=10, color=BLUE_IN)
        c.number_format = N_DASH
    de.cell(row=er, column=5, value=note).font = Font(name=FONT, size=9, italic=True, color=MUTED)
    for col in range(1, 6):
        de.cell(row=er, column=col).border = box
        de.cell(row=er, column=col).protection = UNLOCK
    er += 1

for r in range(er, ENTRY_FMT + 1):
    for col in range(1, 6):
        cell = de.cell(row=r, column=col)
        cell.border = box
        cell.font = Font(name=FONT, size=10, color=BLUE_IN)
        cell.protection = UNLOCK
    de.cell(row=r, column=1).number_format = DATE_F
    de.cell(row=r, column=3).number_format = N_DASH
    de.cell(row=r, column=4).number_format = N_DASH

# Rows past the ruled block must stay unlocked too, or entry dies partway
# through the year. Unlocking them cell by cell does not survive: LibreOffice
# discards empty cells that carry only a style, which re-locked everything
# past ~row 1020. A column-level default costs nothing and cannot be dropped,
# because an unmaterialised cell inherits its column's style.
_entry_fmt = {'A': DATE_F, 'B': 'General', 'C': N_DASH, 'D': N_DASH, 'E': 'General'}
for _col, _fmt in _entry_fmt.items():
    _cd = de.column_dimensions[_col]
    _cd.protection = Protection(locked=False)
    _cd.font = Font(name=FONT, size=10, color=BLUE_IN)
    _cd.number_format = _fmt

wb.defined_names.add(DefinedName(
    'Products',
    attr_text=f'OFFSET(Stock!$L${first},0,0,MAX(1,COUNTA(Stock!$A${first}:$A${last})),1)'))
dv_item = DataValidation(type='list', formula1='=Products', allow_blank=True)
dv_item.error = 'Pick a product from the dropdown so the entry reaches the Stock sheet.'
dv_item.errorTitle = 'Unknown product'
de.add_data_validation(dv_item)
dv_item.add(f'B{DH+1}:B{ENTRY_MAX}')

dv_qty = DataValidation(type='decimal', operator='greaterThanOrEqual', formula1='0', allow_blank=True)
dv_qty.error = 'Slabs cannot be negative. Record a despatch in the Shipment column.'
dv_qty.errorTitle = 'Negative quantity'
de.add_data_validation(dv_qty)
dv_qty.add(f'C{DH+1}:D{ENTRY_MAX}')

dv_date = DataValidation(type='date', operator='greaterThan', formula1='DATE(2000,1,1)', allow_blank=True)
dv_date.error = 'Enter a real date such as 21/08/2026, not text.'
dv_date.errorTitle = 'Date required'
de.add_data_validation(dv_date)
dv_date.add(f'A{DH+1}:A{ENTRY_MAX}')

de['G4'] = 'Rows that will not reach Stock:'
de['G4'].font = Font(name=FONT, size=9, bold=True, color=MUTED)
de['J4'] = (f'=SUMPRODUCT((B{DH+1}:B{ENTRY_MAX}<>"")*'
            f'(COUNTIF(Stock!$L${first}:$L${last},B{DH+1}:B{ENTRY_MAX})=0))')
de['J4'].font = Font(name=FONT, size=10, bold=True, color='B3261E')
de['K4'] = '← must stay at 0; anything else is a mistyped product'
de['K4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)
de.column_dimensions['G'].width = 24
de.column_dimensions['J'].width = 6

de.freeze_panes = f'A{DH+1}'
de.auto_filter.ref = f'A{DH}:E{ENTRY_FMT}'
de.print_title_rows = f'{DH}:{DH}'
de.page_setup.fitToWidth = 1
de.page_setup.fitToHeight = 0
de.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Summary
sm = wb.create_sheet('Summary')
title_block(sm, 'SUMMARY BY BRAND')
sm['A4'] = 'Closing Stock'
sm['A4'].font = Font(name=FONT, size=9, bold=True, color=MUTED)
sm['A5'] = f'=Stock!$H${tot}'
sm['A5'].font = Font(name=FONT, size=16, bold=True, color=DEEP)
sm['A5'].number_format = N_IND
sm['A6'] = 'slabs'
sm['A6'].font = Font(name=FONT, size=9, color=MUTED)
sm['C4'] = 'Stock Value'
sm['C4'].font = Font(name=FONT, size=9, bold=True, color=MUTED)
sm['C5'] = f'=IF(Stock!$J${tot}="","—",Stock!$J${tot})'
sm['C5'].font = Font(name=FONT, size=16, bold=True, color=DEEP)
sm['C5'].number_format = CUR_IND
sm['C6'] = 'at the export rates and exchange rate on Stock'
sm['C6'].font = Font(name=FONT, size=9, color=MUTED)
sm['E4'] = 'Stock Value (US$)'
sm['E4'].font = Font(name=FONT, size=9, bold=True, color=MUTED)
sm['E5'] = (f'=IF(OR(Stock!$C$4="",Stock!$J${tot}=""),"—",Stock!$J${tot}/Stock!$C$4)')
sm['E5'].font = Font(name=FONT, size=16, bold=True, color=DEEP)
sm['E5'].number_format = r'"$" #,##0.00'
sm['E6'] = 'the same stock at the export prices'
sm['E6'].font = Font(name=FONT, size=9, color=MUTED)

SH = 8
head(sm, SH, ['Brand', 'Opening Balance', "Today's Production", "Today's Shipment",
              'Closing Balance', 'Stock Value (₹)'], [24, 15, 15, 15, 14, 18], align_from=2)
sr = SH + 1
BRAND_ROWS = list(BRANDS) + [None] * SPARE_BRANDS
for i, b in enumerate(BRAND_ROWS):
    r = sr + i
    sm.cell(row=r, column=1, value=b).font = Font(name=FONT, size=10, color=INK)
    if b is None:
        sm.cell(row=r, column=1).protection = UNLOCK   # room for a new brand
    for col, src_col in ((2, 'C'), (3, 'D'), (4, 'E'), (5, 'H'), (6, 'J')):
        rng = (f'Stock!$A${first}:$A${last},$A{r},'
               f'Stock!${src_col}${first}:${src_col}${last}')
        inner = f'SUMIF({rng})'
        c = sm.cell(row=r, column=col,
                    value=(f'=IF(OR($A{r}="",{inner}=0),"",{inner})' if col == 6
                           else f'=IF($A{r}="","",{inner})'))
        c.number_format = CUR_IND if col == 6 else (N_DASH if col in (3, 4) else N_IND)
        c.font = Font(name=FONT, size=10, color=INK, bold=(col == 5))
    for col in range(1, 7):
        cell = sm.cell(row=r, column=col)
        cell.border = box
        if i % 2:
            cell.fill = PatternFill('solid', fgColor=WASH)

slast = sr + len(BRAND_ROWS) - 1
stot = slast + 1
sm.cell(row=stot, column=1, value='GRAND TOTAL')
for col in range(1, 7):
    c = sm.cell(row=stot, column=col)
    c.fill = PatternFill('solid', fgColor=DEEP)
    c.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    c.border = box
    if col > 1:
        L = get_column_letter(col)
        rng = f'{L}{sr}:{L}{slast}'
        c.value = f'=IF(SUM({rng})=0,"",SUM({rng}))' if col == 6 else f'=SUM({rng})'
        c.number_format = CUR_IND if col == 6 else (N_DASH if col in (3, 4) else N_IND)

MISMATCH = (f'ROUND(SUM($E${sr}:$E${slast}),2)<>ROUND(Stock!$H${tot},2)')
warn = sm.cell(row=stot + 2, column=1,
               value=f'=IF({MISMATCH},"A brand used on Stock is missing from this list — '
                     f'type it into a blank row above, or the totals understate the stock.","")')
warn.font = Font(name=FONT, size=10, bold=True, color='B3261E')
sm.cell(row=stot + 3, column=1,
        value="Brand totals read straight off 'Stock', so they can never disagree with it.")\
    .font = Font(name=FONT, size=9, italic=True, color=MUTED)
sm.print_area = f'A1:F{stot}'
sm.page_setup.fitToWidth = 1
sm.page_setup.fitToHeight = 0
sm.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Daily Report
dr = wb.create_sheet('Daily Report')
for col, w in zip('ABCDEF', (23, 24, 15, 15, 16, 19)):
    dr.column_dimensions[col].width = w
dr.column_dimensions['G'].width = 2
dr.column_dimensions['H'].width = 6      # match helper, outside the print area

dr.row_dimensions[1].height = 34
dr.row_dimensions[2].height = 14
dr.row_dimensions[3].height = 13
if os.path.exists(LOGO):
    img = XLImage(LOGO)
    img.height, img.width = 42, 52
    dr.add_image(img, 'A1')

dr['B1'] = 'Premier Exports International'
dr['B1'].font = Font(name=FONT, size=17, color=BRAND_C)
dr['B1'].alignment = Alignment(vertical='center')
for i, t in enumerate(['AP X/453, NH-66 Highway, Chandiroor P.O.,',
                       'Aroor, Alappuzha, Kerala - 688 537, India',
                       'GSTIN 32AADFP3158P1ZZ']):
    c = dr.cell(row=1 + i, column=6, value=t)
    c.font = Font(name=FONT, size=8, color=MUTED)
    c.alignment = Alignment(horizontal='right', vertical='center')
dr.merge_cells('E1:F1'); dr.merge_cells('E2:F2'); dr.merge_cells('E3:F3')
for i in range(3):
    dr.cell(row=1 + i, column=5).alignment = Alignment(horizontal='right', vertical='center')
    dr.cell(row=1 + i, column=5).font = Font(name=FONT, size=8, color=MUTED)
dr['E1'] = 'AP X/453, NH-66 Highway, Chandiroor P.O.,'
dr['E2'] = 'Aroor, Alappuzha, Kerala - 688 537, India'
dr['E3'] = 'GSTIN 32AADFP3158P1ZZ'

dr.row_dimensions[4].height = 5
for col in range(1, 7):
    dr.cell(row=4, column=col).border = Border(
        bottom=Side(style='medium', color=BRAND_C))

dr.row_dimensions[5].height = 26
dr['A5'] = 'DAILY STOCK REPORT'
dr['A5'].font = Font(name=FONT, size=17, bold=True, color=INK)
dr['A5'].alignment = Alignment(vertical='center')
dr['A6'] = 'PURCHASE & PROCESSING · PRODUCTION'
dr['A6'].font = Font(name=FONT, size=8, bold=True, color=MUTED)
dr.merge_cells('E5:F5')
dr['E5'] = '=Stock!$C$3'
dr['E5'].number_format = DATE_F
dr['E5'].font = Font(name=FONT, size=13, bold=True, color=DEEP)
dr['E5'].alignment = Alignment(horizontal='right', vertical='center')
dr.merge_cells('E6:F6')
dr['E6'] = 'STATEMENT DATE'
dr['E6'].font = Font(name=FONT, size=8, bold=True, color=MUTED)
dr['E6'].alignment = Alignment(horizontal='right')

KPI = 8
kpis = [('Opening Balance', f'=Stock!$C${tot}', N_IND, 'slabs'),
        ("Today's Production", f'=Stock!$D${tot}', N_DASH, 'slabs'),
        ("Today's Shipment", f'=Stock!$E${tot}', N_DASH, 'slabs'),
        ('Closing Balance', f'=Stock!$H${tot}', N_IND, 'slabs')]
for i, (lab, f, fmt, unit) in enumerate(kpis):
    col = 1 + i
    k = dr.cell(row=KPI, column=col, value=lab.upper())
    k.font = Font(name=FONT, size=8, bold=True, color=MUTED)
    k.fill = PatternFill('solid', fgColor=WASH)
    k.border = Border(left=thin, right=thin, top=thin)
    k.alignment = Alignment(horizontal='left', indent=1)
    v = dr.cell(row=KPI + 1, column=col, value=f)
    v.font = Font(name=FONT, size=14, bold=True, color=INK)
    v.number_format = fmt
    v.fill = PatternFill('solid', fgColor=WASH)
    v.border = Border(left=thin, right=thin, bottom=thin)
    v.alignment = Alignment(horizontal='left', indent=1)
dr.merge_cells(f'E{KPI}:F{KPI}')
dr.merge_cells(f'E{KPI+1}:F{KPI+1}')
kv = dr.cell(row=KPI, column=5,
             value=f'=IF(Stock!$C$4="","STOCK VALUE (₹)",'
                   f'"STOCK VALUE (₹) AT "&TEXT(Stock!$C$4,"0.00")&" / US$")')
kv.font = Font(name=FONT, size=8, bold=True, color='CFE8F2')
kv.fill = PatternFill('solid', fgColor=DEEP)
kv.alignment = Alignment(horizontal='left', indent=1)
vv = dr.cell(row=KPI + 1, column=5,
             value=f'=IF(Stock!$C$4="","Exchange rate not set",'
                   f'IF(Stock!$J${tot}="","Export rates not entered",Stock!$J${tot}))')
vv.font = Font(name=FONT, size=14, bold=True, color='FFFFFF')
vv.number_format = CUR_IND
vv.fill = PatternFill('solid', fgColor=DEEP)
vv.alignment = Alignment(horizontal='left', indent=1)
for c in (5, 6):
    dr.cell(row=KPI, column=c).fill = PatternFill('solid', fgColor=DEEP)
    dr.cell(row=KPI + 1, column=c).fill = PatternFill('solid', fgColor=DEEP)
dr.row_dimensions[KPI].height = 13
dr.row_dimensions[KPI + 1].height = 22

# never let an incomplete brand list go out as a finished report
dr.merge_cells(start_row=KPI + 2, start_column=1, end_row=KPI + 2, end_column=6)
alert = dr.cell(row=KPI + 2, column=1,
                value=f'=IF(ROUND(Summary!$E${stot},2)<>ROUND(Stock!$H${tot},2),'
                      f'"A brand on Stock is missing from the Summary list — this report '
                      f'understates the total. Add it before sending.","")')
alert.font = Font(name=FONT, size=9, bold=True, color='B3261E')
dr.row_dimensions[KPI + 2].height = 12

def section(row, text):
    dr.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    c = dr.cell(row=row, column=1, value=text.upper())
    c.font = Font(name=FONT, size=9, bold=True, color=DEEP)
    c.fill = PatternFill('solid', fgColor=BAND)
    c.alignment = Alignment(vertical='center')
    dr.row_dimensions[row].height = 17
    for col in range(1, 7):
        dr.cell(row=row, column=col).fill = PatternFill('solid', fgColor=BAND)
        dr.cell(row=row, column=col).border = Border(top=thin, bottom=thin)

def sub_head(row, cols, align_from=3):
    for i, t in enumerate(cols, start=1):
        c = dr.cell(row=row, column=i, value=t)
        c.font = Font(name=FONT, size=8, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=DEEP)
        c.alignment = Alignment(horizontal='right' if i >= align_from else 'left', wrap_text=True)
        c.border = box
    dr.row_dimensions[row].height = 22

# --- position by brand (fixed height, so it never floats) ---
BSEC = 11
section(BSEC, 'Position by brand')
sub_head(BSEC + 1, ['Brand', 'Opening Balance', "Today's Production", "Today's Shipment",
                    'Closing Balance', 'Stock Value (₹)'], align_from=2)
brow = BSEC + 2
for i in range(len(BRAND_ROWS)):
    r = brow + i
    src_r = sr + i
    for col, L in enumerate('ABCDEF', start=1):
        # A plain =Summary!A17 on an empty cell returns 0, not blank — which
        # printed a stray zero on every spare brand row and defeated the
        # rule that hides their borders. Test the source brand first.
        c = dr.cell(row=r, column=col,
                    value=f'=IF(Summary!$A${src_r}="","",Summary!${L}${src_r})')
        c.border = box
        c.font = Font(name=FONT, size=9, color=INK, bold=(col == 5))
        c.number_format = (CUR_IND if col == 6 else
                           N_DASH if col in (3, 4) else
                           N_IND if col > 1 else 'General')
        if i % 2:
            c.fill = PatternFill('solid', fgColor=WASH)
    dr.row_dimensions[r].height = 13
blast = brow + len(BRAND_ROWS) - 1
btot = blast + 1
for col, L in enumerate('ABCDEF', start=1):
    c = dr.cell(row=btot, column=col, value=f'=Summary!${L}${stot}')
    c.fill = PatternFill('solid', fgColor=DEEP)
    c.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
    c.border = box
    c.number_format = (CUR_IND if col == 6 else N_DASH if col in (3, 4)
                       else N_IND if col > 1 else 'General')
dr.row_dimensions[btot].height = 15

# --- today's movements ---
MSEC = btot + 2
section(MSEC, "Today's movements")
sub_head(MSEC + 1, ['Brand', 'Product', 'Production', 'Shipment', 'Closing Balance', ''])
mrow = MSEC + 2
for i in range(MOVE_ROWS):
    r = mrow + i
    rank = i + 1
    dr.cell(row=r, column=8, value=f'=IFERROR(MATCH({rank},Stock!$N${first}:$N${last},0),"")')\
      .font = Font(name=FONT, size=8, color=MUTED)
    for col, src_col in ((1, 'A'), (2, 'B'), (3, 'D'), (4, 'E'), (5, 'H')):
        c = dr.cell(row=r, column=col,
                    value=f'=IF($H{r}="","",INDEX(Stock!${src_col}${first}:${src_col}${last},$H{r}))')
        c.font = Font(name=FONT, size=9, color=INK, bold=(col == 5))
        c.number_format = N_IND if col == 5 else (N_DASH if col in (3, 4) else 'General')
    dr.cell(row=r, column=6, value='')
    dr.row_dimensions[r].height = 13
mlast = mrow + MOVE_ROWS - 1
# rule the row only when it carries a movement, so unused lines print blank
dr.conditional_formatting.add(
    f'A{mrow}:F{mlast}',
    FormulaRule(formula=[f'$B{mrow}<>""'],
                border=Border(left=thin, right=thin, top=thin, bottom=thin), stopIfTrue=False))
dr.conditional_formatting.add(
    f'A{brow}:F{blast}',
    FormulaRule(formula=[f'$A{brow}=""'],
                border=Border(), stopIfTrue=True))

mtot = mlast + 1
dr.cell(row=mtot, column=1, value='TOTAL MOVED TODAY')
for col in range(1, 7):
    c = dr.cell(row=mtot, column=col)
    c.fill = PatternFill('solid', fgColor=WASH)
    c.border = Border(top=Side(style='medium', color=DEEP), bottom=thin, left=thin, right=thin)
    c.font = Font(name=FONT, size=9, bold=True, color=INK)
dr.cell(row=mtot, column=3, value=f'=SUM(C{mrow}:C{mlast})').number_format = N_DASH
dr.cell(row=mtot, column=4, value=f'=SUM(D{mrow}:D{mlast})').number_format = N_DASH
dr.cell(row=mtot, column=2,
        value=f'=COUNTIF(Stock!$M${first}:$M${last},1)&" product(s)"')
dr.cell(row=mtot, column=2).font = Font(name=FONT, size=9, color=MUTED)
dr.cell(row=mtot, column=2).alignment = Alignment(horizontal='left')
dr.row_dimensions[mtot].height = 15

note = mtot + 1
dr.merge_cells(start_row=note, start_column=1, end_row=note, end_column=6)
dr.cell(row=note, column=1, value=(
    f'=IF(COUNTIF(Stock!$M${first}:$M${last},1)>{MOVE_ROWS},'
    f'"Note: "&(COUNTIF(Stock!$M${first}:$M${last},1)-{MOVE_ROWS})&'
    f'" further product(s) moved today — see the Stock sheet for the full list.","")'))
dr.cell(row=note, column=1).font = Font(name=FONT, size=8, bold=True, color='B3261E')

foot = note + 2
dr.merge_cells(start_row=foot, start_column=1, end_row=foot, end_column=6)
fc = dr.cell(row=foot, column=1,
             value='Prepared from the stock register. Closing Balance = Opening Balance '
                   '+ Production to Date − Shipment to Date.')
fc.font = Font(name=FONT, size=8, italic=True, color=MUTED)
for col in range(1, 7):
    dr.cell(row=foot, column=col).border = Border(top=thin)

dr.print_area = f'A1:F{foot}'
dr.page_setup.orientation = 'portrait'
dr.page_setup.fitToWidth = 1
dr.page_setup.fitToHeight = 1
dr.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
dr.page_margins.left = dr.page_margins.right = 0.4
dr.page_margins.top = dr.page_margins.bottom = 0.4

wb.active = wb.index(dr)
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False

# ---------------------------------------------------------------- protection
# Cells are locked by default; only those given UNLOCK above accept typing.
# Locking bites only once sheet protection is switched on, below.
for ws in wb.worksheets:
    ws.protection.password = PASSWORD
    ws.protection.sheet = True
    ws.protection.enable()
    # True == "locked", i.e. the operation is NOT allowed
    ws.protection.formatCells = True      # no restyling
    ws.protection.formatColumns = True    # no widths
    ws.protection.formatRows = True       # no heights
    ws.protection.insertRows = True
    ws.protection.insertColumns = True
    ws.protection.deleteRows = True
    ws.protection.deleteColumns = True
    ws.protection.insertHyperlinks = True
    ws.protection.sort = True             # sorting would scramble the log
    ws.protection.pivotTables = True
    ws.protection.objects = True
    ws.protection.scenarios = True
    ws.protection.autoFilter = False      # filtering stays available
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False

# stop sheets being added, renamed, moved or deleted
wb.security = WorkbookProtection(workbookPassword=PASSWORD, lockStructure=True)
wb.save(OUT)
print(f'wrote {OUT}: {len(STOCK)} products, {len(LOG)} log rows, {len(BRANDS)} brands, '
      f'report prints {MOVE_ROWS} movement lines')
