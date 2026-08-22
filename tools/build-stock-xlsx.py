#!/usr/bin/env python3
"""Builds PEI-Stock-Register.xlsx — a daily stock register driven by formulas,
with per-brand valuation tables and a one-page Daily Report for the directors.

    python tools/build-stock-xlsx.py out.xlsx                    # from stock-data.json
    python tools/build-stock-xlsx.py out.xlsx --from live.xlsx   # keep a live file's data
    python tools/build-stock-xlsx.py --relock out.xlsx           # after recalc

Two models sit underneath this sheet.

Roll-forward. Opening Balance is a FROZEN baseline; every later movement is a
dated row on 'Daily Entry'. Closing = baseline + production to date − shipment
to date, so the register rolls itself forward and nothing is ever re-keyed.

Units. Balances are counted in each product's own unit — 2 kg slabs for the
shrimp brands, kilos for tuna and mackerel, other slab weights by buyer — so
slabs and kilos must never be added together. Every product carries a Kg / Unit
pack size, and all totalling happens in kilos. The rate is per kilo, in US
dollars, and the stock reports in rupees at one exchange rate.
"""
import sys, os, json, datetime
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
DATA = os.path.join(ROOT, 'tools', 'stock-data.json')

args = [a for a in sys.argv[1:] if not a.startswith('--')]
OUT = args[0] if args else 'PEI-Stock-Register.xlsx'
SRC = sys.argv[sys.argv.index('--from') + 1] if '--from' in sys.argv else None
# Excel sheet protection is a guardrail against accidents, not security — the
# password is trivially removable. It stops staff overwriting a formula.
PASSWORD = (sys.argv[sys.argv.index('--password') + 1]
            if '--password' in sys.argv else 'PEI2026')
UNLOCK = Protection(locked=False)

ENTRY_FIRST, ENTRY_MAX, ENTRY_FMT = 7, 5000, 250
MOVE_ROWS = 20            # movement lines the daily report prints
SPARE_PRODUCTS = 25       # blank rows under the last product on 'Stock'
SLOT_SPARES = 3           # blank product slots in each Summary brand table
SPARE_BLOCKS = 2          # blank brand tables on Summary

# ---------------------------------------------------------------- relock
# LibreOffice discards several kinds of protection when it recalculates, and
# re-saving through openpyxl would strip every cached formula value. So they
# are re-applied by patching the XML in place, after recalc.
if '--relock' in sys.argv:
    import zipfile, shutil, tempfile, re as _re
    from openpyxl.utils.protection import hash_password
    target = sys.argv[sys.argv.index('--relock') + 1]
    zin = zipfile.ZipFile(target)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    blob = {i.filename: d for i, d in items}

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

    rid = dict(_re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"',
                           blob['xl/_rels/workbook.xml.rels'].decode('utf-8')))
    sheet_file = None
    for m in _re.finditer(r'<sheet\b[^>]*/>', wbx):
        a = dict(_re.findall(r'(?:r:)?(\w+)="([^"]*)"', m.group(0)))
        if a.get('name') == 'Daily Entry':
            sheet_file = 'xl/' + rid.get(a.get('id', ''), '').lstrip('/')
            break
    added = []
    if sheet_file and sheet_file in blob:
        sx = blob[sheet_file].decode('utf-8')
        st_xml = blob['xl/styles.xml'].decode('utf-8')
        head, rest = st_xml.split('<cellXfs', 1)
        attrs, body_and_tail = rest.split('>', 1)
        body, tail = body_and_tail.split('</cellXfs>', 1)
        xfs = _re.findall(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', body, _re.S)

        def unlocked_copy(xf):
            if '<protection' in xf:
                # LibreOffice writes an explicit locked="1"; flip it
                xf = _re.sub(r'<protection\b[^>]*/>', '<protection locked="0" hidden="0"/>',
                             xf, count=1)
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
                o, i2 = base.split('>', 1)
                base = o + ' applyProtection="1">' + i2
            return base + '<protection locked="0" hidden="0"/></xf>'

        def repoint(m):
            c = m.group(0)
            a = dict(_re.findall(r'(\w+)="([^"]*)"', c))
            lo, hi = int(a.get('min', 0)), int(a.get('max', 0))
            if lo > 5 or hi < 1:
                return c
            new_id = len(xfs) + len(added)
            added.append(unlocked_copy(xfs[int(a.get('style', 0))]))
            if 'style=' in c:
                return _re.sub(r'style="\d+"', f'style="{new_id}"', c)
            return c[:-2] + f' style="{new_id}"/>'

        sx = _re.sub(r'<col\b[^>]*/>', repoint, sx)
        sx = _re.sub(rf'sqref="([A-Z]+){ENTRY_FIRST}:([A-Z]+)(\d+)"',
                     lambda m: f'sqref="{m.group(1)}{ENTRY_FIRST}:{m.group(2)}{ENTRY_MAX}"', sx)
        if added:
            body += ''.join(added)
            attrs = _re.sub(r'count="\d+"', f'count="{len(xfs) + len(added)}"', attrs)
            blob['xl/styles.xml'] = (head + '<cellXfs' + attrs + '>' + body +
                                     '</cellXfs>' + tail).encode('utf-8')
            blob[sheet_file] = sx.encode('utf-8')

    blob['xl/workbook.xml'] = wbx.encode('utf-8')
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, _ in items:
            zout.writestr(info, blob[info.filename])
    tmp.close()
    shutil.move(tmp.name, target)
    os.chmod(target, 0o644)
    print(f'relocked {target}: structure locked, {len(added)} entry columns re-unlocked')
    sys.exit(0)

# ---------------------------------------------------------------- data
def as_date(v):
    """Log dates must be real dates: text never matches a date criterion in
    SUMIFS, which would silently zero the two 'today' columns."""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    for f in ('%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%m-%Y', '%d/%m/%y'):
        try:
            return datetime.datetime.strptime(str(v).strip(), f).date()
        except (ValueError, TypeError):
            pass
    return None

seed = json.load(open(DATA))
ROWS = [tuple(r) for r in seed['rows']]          # (brand, product, unit, kg_per_unit, opening)
LOG = [(as_date(d), it, p, s) for d, it, p, s in seed['log']]
LOG = [x for x in LOG if x[0]]
FX = seed.get('fx')
STMT_DATE = as_date(seed.get('date')) or datetime.date(2026, 8, 20)
RATES = dict(seed.get('rates') or {})

if SRC:
    src = load_workbook(SRC)
    s_st, s_de = src['Stock'], src['Daily Entry']
    hdr = {s_st.cell(row=7, column=c).value: get_column_letter(c) for c in range(1, 30)}
    cU, cK = hdr.get('Unit'), hdr.get('Kg / Unit')
    cO = hdr.get('Opening Balance', 'C')
    cR = hdr.get('Rate / Kg (US$)') or hdr.get('Rate / Slab (US$)') or 'I'
    ROWS, RATES, seen = [], {}, set()
    r = 8
    while r < 5000:
        a, b = s_st[f'A{r}'].value, s_st[f'B{r}'].value
        if a == 'GRAND TOTAL':
            break
        if a and b:
            key = f'{a} - {b}'
            if key in seen:
                print(f'  skipped duplicate product: {key} (row {r})')
            else:
                seen.add(key)
                ob = s_st[f'{cO}{r}'].value
                unit = s_st[f'{cU}{r}'].value if cU else None
                kg = s_st[f'{cK}{r}'].value if cK else None
                ROWS.append((a, b, unit or 'Slab',
                             kg if isinstance(kg, (int, float)) else None,
                             ob if isinstance(ob, (int, float)) else 0))
                rate = s_st[f'{cR}{r}'].value
                if isinstance(rate, (int, float)) and rate:
                    RATES[key] = rate
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
        LOG.append((d, item, s_de[f'C{r}'].value or 0, s_de[f'D{r}'].value or 0))
    fxcell = hdr.get('Kg / Unit') and s_st['E4'].value or s_st['C4'].value
    FX = fxcell if isinstance(fxcell, (int, float)) else FX
    sd = as_date(s_st['E3'].value) or as_date(s_st['C3'].value)
    STMT_DATE = sd or STMT_DATE
    print(f'read {SRC}: {len(ROWS)} products, {len(LOG)} log rows, {len(RATES)} rates'
          + (f', {skipped} skipped (bad date)' if skipped else ''))

BRANDS = []
for r in ROWS:
    if r[0] not in BRANDS:
        BRANDS.append(r[0])

# ---------------------------------------------------------------- style
FONT = 'Arial'
DEEP, BAND, WASH, LINE = '065E7A', 'E7F3F9', 'F2F8FB', 'CBDFE8'
INK, MUTED, BRAND_C, RED = '0B2C3A', '5E7C8B', '00A2D3', 'B3261E'
BLUE_IN, YELLOW, GREY = '0000FF', 'FFFF00', 'D9D9D9'

thin = Side(style='thin', color=LINE)
box = Border(left=thin, right=thin, top=thin, bottom=thin)

N_IND = r'[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;#,##0'
KG_F = r'[>=100000]##\,##\,##0.0;#,##0.0'
N_DASH = r'#,##0;-#,##0;"–"'
CUR_IND = r'[>=10000000]"₹" ##\,##\,##\,##0.00;[>=100000]"₹" ##\,##\,##0.00;"₹" #,##0.00'
USD_F = r'"$" #,##0.00'
PACK_F = r'#,##0.###'
DATE_F = 'DD/MM/YYYY'

wb = Workbook()

def head_row(ws, row, cols, widths=None, align_from=3, size=9, height=30):
    for i, t in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = Font(name=FONT, size=size, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=DEEP)
        c.alignment = Alignment(horizontal='right' if i >= align_from else 'left',
                                vertical='bottom', wrap_text=True)
        c.border = box
    ws.row_dimensions[row].height = height
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

def title_block(ws, subtitle):
    ws['A1'] = 'PREMIER EXPORTS INTERNATIONAL'
    ws['A1'].font = Font(name=FONT, size=14, bold=True, color=DEEP)
    ws['A2'] = subtitle
    ws['A2'].font = Font(name=FONT, size=10, bold=True, color=MUTED)

# ================================================================ Stock
st = wb.create_sheet('Stock')
title_block(st, 'STOCK REGISTER — LIVE POSITION')
st['A3'] = 'Statement date'
st['A4'] = 'US$ → ₹ exchange rate'
for a in ('A3', 'A4'):
    st[a].font = Font(name=FONT, size=10, bold=True, color=INK)
st['E3'] = STMT_DATE
st['E3'].number_format = DATE_F
st['E4'] = FX
st['E4'].number_format = r'"₹" #,##0.00'
for a in ('E3', 'E4'):
    st[a].font = Font(name=FONT, size=10, bold=True, color=BLUE_IN)
    st[a].fill = PatternFill('solid', fgColor=YELLOW)
    st[a].border = box
    st[a].alignment = Alignment(horizontal='center')
    st[a].protection = UNLOCK
st['F3'] = "← today; drives the two \"today\" columns and the Daily Report"
st['F4'] = "← rupees per US dollar; converts the per-kilo export rates"
for a in ('F3', 'F4'):
    st[a].font = Font(name=FONT, size=9, italic=True, color=MUTED)
st['A5'] = ("Balances are counted in each product's own Unit. Only kilos are comparable "
            "across products, so every total is taken in kilos.")
st['A5'].font = Font(name=FONT, size=9, italic=True, color=MUTED)

HDR = 7
COLS = ['Brand', 'Product', 'Unit', 'Kg / Unit', 'Opening Balance', "Today's Production",
        "Today's Shipment", 'Production to Date', 'Shipment to Date', 'Closing Balance',
        'Closing Kg', 'Rate / Kg (US$)', 'Stock Value (₹)', '',
        'Key (do not edit)', 'Moved today', 'Rank', 'Brand seq',
        'Opening kg', 'Prod kg today', 'Ship kg today', 'Prod kg to date',
        'Ship kg to date', 'Needs Kg/Unit', 'Duplicate key']
WIDTHS = [18, 24, 8, 9, 13, 12, 12, 12, 12, 13, 12, 12, 16, 2,
          32, 11, 6, 24, 11, 12, 12, 12, 12, 12, 12]
head_row(st, HDR, COLS, WIDTHS)
st.cell(row=HDR, column=14).fill = PatternFill('solid', fgColor='FFFFFF')
for c in range(15, 26):
    cc = st.cell(row=HDR, column=c)
    cc.fill = PatternFill('solid', fgColor=GREY)
    cc.font = Font(name=FONT, size=9, bold=True, color=MUTED)

st.cell(row=HDR, column=4).comment = Comment(
    'Kilos in one unit of this product.\n'
    '  2   a 2 kg slab (the shrimp brands)\n'
    '  1   the balance is already in kilos (tuna, mackerel)\n'
    '  other  slab weights that vary by buyer — set per product.\n'
    'Leave it blank and the row is not valued; the count of unset rows\n'
    'is shown under the table and on the Daily Report.',
    'Premier Exports International', height=120, width=360)

first = HDR + 1
ent = lambda col: f"'Daily Entry'!${col}${ENTRY_FIRST}:${col}${ENTRY_MAX}"
ALL = list(ROWS) + [(None, None, None, None, None)] * SPARE_PRODUCTS
for i, (brand, product, unit, kgpu, ob) in enumerate(ALL):
    r = first + i
    spare = brand is None
    g = lambda f: f'=IF($A{r}="","",{f})'
    st.cell(row=r, column=1, value=brand)
    st.cell(row=r, column=2, value=product)
    st.cell(row=r, column=3, value=unit)
    st.cell(row=r, column=4, value=kgpu).number_format = PACK_F
    st.cell(row=r, column=5, value=ob).number_format = N_IND
    st.cell(row=r, column=6, value=g(f"SUMIFS({ent('C')},{ent('B')},$O{r},{ent('A')},$E$3)"))
    st.cell(row=r, column=7, value=g(f"SUMIFS({ent('D')},{ent('B')},$O{r},{ent('A')},$E$3)"))
    st.cell(row=r, column=8, value=g(f"SUMIFS({ent('C')},{ent('B')},$O{r})"))
    st.cell(row=r, column=9, value=g(f"SUMIFS({ent('D')},{ent('B')},$O{r})"))
    st.cell(row=r, column=10, value=g(f'$E{r}+$H{r}-$I{r}'))
    st.cell(row=r, column=11, value=f'=IF(OR($A{r}="",$D{r}=""),"",$J{r}*$D{r})')
    rate_cell = st.cell(row=r, column=12,
                        value=RATES.get(f'{brand} - {product}') if brand else None)
    rate_cell.number_format = USD_F
    # value: kilos x per-kilo export price x the one exchange rate
    st.cell(row=r, column=13,
            value=f'=IF(OR($A{r}="",$D{r}="",$L{r}="",$E$4=""),"",$K{r}*$L{r}*$E$4)')
    helpers = [
        (15, f'=IF($A{r}="","",$A{r}&" - "&$B{r})'),
        (16, f'=IF($A{r}="",0,IF(OR($F{r}<>0,$G{r}<>0),1,0))'),
        (17, f'=IF($P{r}=1,SUM($P${first}:$P{r}),"")'),
        (18, f'=IF($A{r}="","",$A{r}&"#"&COUNTIF($A${first}:$A{r},$A{r}))'),
        (19, f'=IF(OR($A{r}="",$D{r}=""),0,$E{r}*$D{r})'),
        (20, f'=IF(OR($A{r}="",$D{r}=""),0,$F{r}*$D{r})'),
        (21, f'=IF(OR($A{r}="",$D{r}=""),0,$G{r}*$D{r})'),
        (22, f'=IF(OR($A{r}="",$D{r}=""),0,$H{r}*$D{r})'),
        (23, f'=IF(OR($A{r}="",$D{r}=""),0,$I{r}*$D{r})'),
        (24, f'=IF(AND($A{r}<>"",$D{r}=""),1,0)'),
        (25, f'=IF($A{r}="",0,IF(COUNTIF($O${first}:$O${first + len(ALL) - 1},$O{r})>1,1,0))'),
    ]
    for col, val in helpers:
        st.cell(row=r, column=col, value=val).font = Font(name=FONT, size=9, color=MUTED)

    for col in (6, 7):
        st.cell(row=r, column=col).number_format = N_DASH
    for col in (8, 9, 10):
        st.cell(row=r, column=col).number_format = N_IND
    st.cell(row=r, column=11).number_format = KG_F
    st.cell(row=r, column=13).number_format = CUR_IND

    for col in range(1, 14):
        c = st.cell(row=r, column=col)
        c.border = box
        c.font = Font(name=FONT, size=10,
                      color=BLUE_IN if col in (3, 4, 5, 12) else INK,
                      bold=(col in (10, 11)))
        if col in (3, 4, 5, 12):
            c.fill = PatternFill('solid', fgColor=YELLOW if col in (4, 12) else 'FFFFFF')
        elif i % 2:
            c.fill = PatternFill('solid', fgColor=WASH)
        if spare and col in (1, 2, 3, 4, 5):
            c.protection = UNLOCK
    st.cell(row=r, column=4).protection = UNLOCK      # pack size varies by buyer
    st.cell(row=r, column=12).protection = UNLOCK     # export price
    st.cell(row=r, column=3).protection = UNLOCK      # unit

last = first + len(ALL) - 1
tot = last + 1
st.cell(row=tot, column=1, value='GRAND TOTAL (kg)')
for col in range(1, 14):
    c = st.cell(row=tot, column=col)
    c.fill = PatternFill('solid', fgColor=DEEP)
    c.border = box
    c.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
st.cell(row=tot, column=11, value=f'=SUM($K${first}:$K${last})').number_format = KG_F
st.cell(row=tot, column=13,
        value=f'=IF(SUM($M${first}:$M${last})=0,"",SUM($M${first}:$M${last}))')\
  .number_format = CUR_IND
for col in (11, 13):
    st.cell(row=tot, column=col).font = Font(name=FONT, size=10, bold=True, color='FFFFFF')

note = tot + 1
st.cell(row=note, column=1,
        value='Balances are in each product\'s own unit, so only the kilo column totals. '
              'Stock Value (₹) = Closing Kg × Rate / Kg (US$) × the exchange rate.')\
  .font = Font(name=FONT, size=9, italic=True, color=MUTED)
warn = note + 1
st.cell(row=warn, column=1,
        value=f'=IF(SUM($X${first}:$X${last})=0,"",SUM($X${first}:$X${last})&'
              f'" product(s) have no Kg / Unit — they are not valued. Fill column D.")')\
  .font = Font(name=FONT, size=10, bold=True, color=RED)
dup = warn + 1
st.cell(row=dup, column=1,
        value=f'=IF(SUM($Y${first}:$Y${last})=0,"",'
              f'"Two rows share a Brand + Product — production would be counted twice. '
              f'Rename one.")')\
  .font = Font(name=FONT, size=10, bold=True, color=RED)

st.freeze_panes = f'C{first}'
st.auto_filter.ref = f'A{HDR}:M{last}'
st.conditional_formatting.add(
    f'J{first}:J{last}',
    CellIsRule(operator='lessThan', formula=['0'],
               fill=PatternFill('solid', fgColor='F8D7DA'),
               font=Font(name=FONT, size=10, bold=True, color=RED)))
st.conditional_formatting.add(
    f'D{first}:D{last}',
    FormulaRule(formula=[f'AND($A{first}<>"",$D{first}="")'],
                fill=PatternFill('solid', fgColor='FDE7E9'), stopIfTrue=False))
st.print_area = f'A1:M{dup}'
st.print_title_rows = f'{HDR}:{HDR}'
st.page_setup.orientation = 'landscape'
st.page_setup.fitToWidth = 1
st.page_setup.fitToHeight = 0
st.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

dv_unit = DataValidation(type='list', formula1='"Slab,Kg"', allow_blank=True)
dv_unit.error = 'Unit must be Slab or Kg.'
dv_unit.errorTitle = 'Unit'
st.add_data_validation(dv_unit)
dv_unit.add(f'C{first}:C{last}')

# ================================================================ Daily Entry
de = wb.create_sheet('Daily Entry')
title_block(de, 'DAILY ENTRY — ADD ONE ROW PER PRODUCT THAT MOVED')
de['A4'] = ("After production, add a row below: date, product from the dropdown, and the "
            "quantity in that product's own unit — slabs for the shrimp brands, kilos for "
            "tuna and mackerel. Never delete past rows.")
de['A4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)

DH = 6
head_row(de, DH, ['Date', 'Item (pick from list)', 'Production (in product unit)',
                  'Shipment (in product unit)', 'Note'], [14, 40, 18, 18, 40])
for col in (1, 2, 5):
    de.cell(row=DH, column=col).alignment = Alignment(horizontal='left', wrap_text=True)

er = DH + 1
for d, item, prod, ship in LOG:
    c = de.cell(row=er, column=1, value=d)
    c.number_format = DATE_F
    de.cell(row=er, column=2, value=item)
    de.cell(row=er, column=3, value=prod).number_format = N_DASH
    de.cell(row=er, column=4, value=ship).number_format = N_DASH
    for col in range(1, 6):
        cc = de.cell(row=er, column=col)
        cc.border = box
        cc.protection = UNLOCK
        cc.font = Font(name=FONT, size=10, color=BLUE_IN)
    er += 1
for r in range(er, ENTRY_FMT + 1):
    for col in range(1, 6):
        cc = de.cell(row=r, column=col)
        cc.border = box
        cc.font = Font(name=FONT, size=10, color=BLUE_IN)
        cc.protection = UNLOCK
    de.cell(row=r, column=1).number_format = DATE_F
    de.cell(row=r, column=3).number_format = N_DASH
    de.cell(row=r, column=4).number_format = N_DASH
# a column-level default survives LibreOffice; per-cell styles on empty cells do not
for col, fmt in (('A', DATE_F), ('B', 'General'), ('C', N_DASH), ('D', N_DASH), ('E', 'General')):
    cd = de.column_dimensions[col]
    cd.protection = Protection(locked=False)
    cd.font = Font(name=FONT, size=10, color=BLUE_IN)
    cd.number_format = fmt

wb.defined_names.add(DefinedName(
    'Products',
    attr_text=f'OFFSET(Stock!$O${first},0,0,MAX(1,COUNTA(Stock!$A${first}:$A${last})),1)'))
dv_item = DataValidation(type='list', formula1='=Products', allow_blank=True)
dv_item.error = 'Pick a product from the dropdown so the entry reaches the Stock sheet.'
dv_item.errorTitle = 'Unknown product'
de.add_data_validation(dv_item)
dv_item.add(f'B{DH+1}:B{ENTRY_MAX}')
dv_qty = DataValidation(type='decimal', operator='greaterThanOrEqual', formula1='0', allow_blank=True)
dv_qty.error = 'Quantity cannot be negative. Record a despatch in the Shipment column.'
dv_qty.errorTitle = 'Negative quantity'
de.add_data_validation(dv_qty)
dv_qty.add(f'C{DH+1}:D{ENTRY_MAX}')
dv_date = DataValidation(type='date', operator='greaterThan', formula1='DATE(2000,1,1)', allow_blank=True)
dv_date.error = 'Enter a real date such as 22/08/2026, not text.'
dv_date.errorTitle = 'Date required'
de.add_data_validation(dv_date)
dv_date.add(f'A{DH+1}:A{ENTRY_MAX}')

de['G4'] = 'Rows that will not reach Stock:'
de['G4'].font = Font(name=FONT, size=9, bold=True, color=MUTED)
de['J4'] = (f'=SUMPRODUCT((B{DH+1}:B{ENTRY_MAX}<>"")*'
            f'(COUNTIF(Stock!$O${first}:$O${last},B{DH+1}:B{ENTRY_MAX})=0))')
de['J4'].font = Font(name=FONT, size=10, bold=True, color=RED)
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
title_block(sm, 'VALUATION BY BRAND')
for col, w in zip('ABCDEFG', (26, 8, 9, 15, 14, 12, 18)):
    sm.column_dimensions[col].width = w
sm.column_dimensions['I'].width = 7

def hero(col, label, formula, fmt, note, span=1):
    # a crore-scale rupee figure needs more width than one column gives,
    # and Excel prints ### rather than wrapping it
    from openpyxl.utils import column_index_from_string as cidx
    c0 = cidx(col)
    for row, val, font in ((4, label, Font(name=FONT, size=9, bold=True, color=MUTED)),
                           (5, formula, Font(name=FONT, size=15, bold=True, color=DEEP)),
                           (6, note, Font(name=FONT, size=9, color=MUTED))):
        if span > 1:
            sm.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c0 + span - 1)
        cell = sm.cell(row=row, column=c0, value=val)
        cell.font = font
        if row == 5:
            cell.number_format = fmt

hero('A', 'Closing Stock', f'=Stock!$K${tot}', KG_F, 'kilos, across every product', span=2)
hero('C', 'Stock Value', f'=IF(Stock!$M${tot}="","—",Stock!$M${tot})', CUR_IND,
     'at the per-kilo export prices', span=3)
hero('F', 'Stock Value (US$)',
     f'=IF(OR(Stock!$E$4="",Stock!$M${tot}=""),"—",Stock!$M${tot}/Stock!$E$4)',
     USD_F, 'the same stock in dollars', span=2)

sm['A7'] = (f'=IF(SUM(Stock!$X${first}:$X${last})=0,"",'
            f'SUM(Stock!$X${first}:$X${last})&" product(s) have no Kg / Unit on Stock — '
            f'they are counted in kilos as zero and are not valued.")')
sm['A7'].font = Font(name=FONT, size=10, bold=True, color=RED)

SUB_ROWS, BLOCK_START = [], 9
row = BLOCK_START
BLOCK_BRANDS = list(BRANDS) + [None] * SPARE_BLOCKS
for bi, brand in enumerate(BLOCK_BRANDS):
    n = sum(1 for x in ROWS if x[0] == brand) if brand else 0
    slots = n + SLOT_SPARES if brand else 5
    band = row
    sm.merge_cells(start_row=band, start_column=1, end_row=band, end_column=7)
    bc = sm.cell(row=band, column=1, value=brand)
    bc.font = Font(name=FONT, size=10, bold=True, color=DEEP)
    for col in range(1, 8):
        sm.cell(row=band, column=col).fill = PatternFill('solid', fgColor=BAND)
        sm.cell(row=band, column=col).border = Border(top=thin, bottom=thin)
    if brand is None:
        bc.protection = UNLOCK
    sm.row_dimensions[band].height = 16

    hrow = band + 1
    head_row(sm, hrow, ['Product', 'Unit', 'Kg / Unit', 'Closing Balance', 'Closing Kg',
                        'Rate / Kg (US$)', 'Stock Value (₹)'], None, align_from=2,
             size=8, height=20)

    p0 = hrow + 1
    for k in range(slots):
        rr = p0 + k
        sm.cell(row=rr, column=9,
                value=f'=IFERROR(MATCH($A${band}&"#"&{k+1},Stock!$R${first}:$R${last},0),"")')\
          .font = Font(name=FONT, size=8, color=MUTED)
        srcs = [(1, 'B'), (2, 'C'), (3, 'D'), (4, 'J'), (5, 'K'), (6, 'L'), (7, 'M')]
        for col, sc in srcs:
            # INDEX on an empty source cell returns 0, not blank — so a product
            # with no pack size or no rate would show a misleading zero.
            idx = f'INDEX(Stock!${sc}${first}:${sc}${last},$I{rr})'
            expr = (f'=IF($I{rr}="","",IF({idx}=0,"",{idx}))' if col in (2, 3, 6)
                    else f'=IF($I{rr}="","",{idx})')
            c = sm.cell(row=rr, column=col, value=expr)
            c.font = Font(name=FONT, size=9, color=INK, bold=(col == 5))
            c.number_format = {1: 'General', 2: 'General', 3: PACK_F, 4: N_IND,
                               5: KG_F, 6: USD_F, 7: CUR_IND}[col]
            c.border = box
            if k % 2:
                c.fill = PatternFill('solid', fgColor=WASH)
        sm.row_dimensions[rr].height = 13
    p1 = p0 + slots - 1
    sub = p1 + 1
    sm.cell(row=sub, column=1, value=f'=IF($A${band}="","",$A${band}&" total")')
    for col in range(1, 8):
        c = sm.cell(row=sub, column=col)
        c.fill = PatternFill('solid', fgColor=WASH)
        c.border = Border(top=Side(style='medium', color=DEEP), bottom=thin,
                          left=thin, right=thin)
        c.font = Font(name=FONT, size=9, bold=True, color=INK)
    for col, L in ((5, 'E'), (7, 'G')):
        sm.cell(row=sub, column=col, value=f'=SUM({L}{p0}:{L}{p1})')\
          .number_format = KG_F if col == 5 else CUR_IND
        sm.cell(row=sub, column=col).font = Font(name=FONT, size=9, bold=True, color=INK)
    SUB_ROWS.append(sub)
    # hide the ruling on unused slots
    sm.conditional_formatting.add(
        f'A{p0}:G{p1}',
        FormulaRule(formula=[f'$A{p0}=""'], border=Border(), stopIfTrue=True))
    row = sub + 2

gt = row
sm.cell(row=gt, column=1, value='GRAND TOTAL')
for col in range(1, 8):
    c = sm.cell(row=gt, column=col)
    c.fill = PatternFill('solid', fgColor=DEEP)
    c.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    c.border = box
sm.cell(row=gt, column=5, value='+'.join(f'E{r}' for r in SUB_ROWS).join(['=', '']))\
  .number_format = KG_F
sm.cell(row=gt, column=7, value='+'.join(f'G{r}' for r in SUB_ROWS).join(['=', '']))\
  .number_format = CUR_IND
for col in (5, 7):
    sm.cell(row=gt, column=col).font = Font(name=FONT, size=10, bold=True, color='FFFFFF')

sm.cell(row=gt + 2, column=1,
        value=f'=IF(ROUND($E${gt},1)<>ROUND(Stock!$K${tot},1),'
              f'"A brand on Stock has no table here — type its name into a blank brand band '
              f'above, or these totals understate the stock.","")')\
  .font = Font(name=FONT, size=10, bold=True, color=RED)
sm.cell(row=gt + 3, column=1,
        value="Every figure is read from 'Stock', so this can never disagree with it.")\
  .font = Font(name=FONT, size=9, italic=True, color=MUTED)

sm.print_area = f'A1:G{gt + 3}'
sm.print_title_rows = '1:2'
sm.page_setup.fitToWidth = 1
sm.page_setup.fitToHeight = 0
sm.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Daily Report
dr = wb.create_sheet('Daily Report')
for col, w in zip('ABCDEF', (23, 24, 15, 15, 16, 19)):
    dr.column_dimensions[col].width = w
dr.column_dimensions['G'].width = 2
dr.column_dimensions['H'].width = 6

dr.row_dimensions[1].height = 34
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
    dr.merge_cells(start_row=1 + i, start_column=5, end_row=1 + i, end_column=6)
    c = dr.cell(row=1 + i, column=5, value=t)
    c.font = Font(name=FONT, size=8, color=MUTED)
    c.alignment = Alignment(horizontal='right', vertical='center')
dr.row_dimensions[4].height = 5
for col in range(1, 7):
    dr.cell(row=4, column=col).border = Border(bottom=Side(style='medium', color=BRAND_C))

dr.row_dimensions[5].height = 26
dr['A5'] = 'DAILY STOCK REPORT'
dr['A5'].font = Font(name=FONT, size=17, bold=True, color=INK)
dr['A5'].alignment = Alignment(vertical='center')
dr['A6'] = 'PURCHASE & PROCESSING · PRODUCTION'
dr['A6'].font = Font(name=FONT, size=8, bold=True, color=MUTED)
dr.merge_cells('E5:F5'); dr.merge_cells('E6:F6')
dr['E5'] = '=Stock!$E$3'
dr['E5'].number_format = DATE_F
dr['E5'].font = Font(name=FONT, size=13, bold=True, color=DEEP)
dr['E5'].alignment = Alignment(horizontal='right', vertical='center')
dr['E6'] = 'STATEMENT DATE'
dr['E6'].font = Font(name=FONT, size=8, bold=True, color=MUTED)
dr['E6'].alignment = Alignment(horizontal='right')

KPI = 8
kpis = [('Opening (kg)', f'=SUM(Stock!$S${first}:$S${last})', KG_F),
        ('Produced today (kg)', f'=SUM(Stock!$T${first}:$T${last})', KG_F),
        ('Shipped today (kg)', f'=SUM(Stock!$U${first}:$U${last})', KG_F),
        ('Closing (kg)', f'=Stock!$K${tot}', KG_F)]
for i, (lab, f, fmt) in enumerate(kpis):
    k = dr.cell(row=KPI, column=1 + i, value=lab.upper())
    k.font = Font(name=FONT, size=8, bold=True, color=MUTED)
    k.fill = PatternFill('solid', fgColor=WASH)
    k.border = Border(left=thin, right=thin, top=thin)
    k.alignment = Alignment(horizontal='left', indent=1)
    v = dr.cell(row=KPI + 1, column=1 + i, value=f)
    v.font = Font(name=FONT, size=13, bold=True, color=INK)
    v.number_format = fmt
    v.fill = PatternFill('solid', fgColor=WASH)
    v.border = Border(left=thin, right=thin, bottom=thin)
    v.alignment = Alignment(horizontal='left', indent=1)
dr.merge_cells(f'E{KPI}:F{KPI}'); dr.merge_cells(f'E{KPI+1}:F{KPI+1}')
kv = dr.cell(row=KPI, column=5,
             value=f'=IF(Stock!$E$4="","STOCK VALUE (₹)",'
                   f'"STOCK VALUE (₹) AT "&TEXT(Stock!$E$4,"0.00")&" / US$")')
kv.font = Font(name=FONT, size=8, bold=True, color='CFE8F2')
kv.alignment = Alignment(horizontal='left', indent=1)
vv = dr.cell(row=KPI + 1, column=5,
             value=f'=IF(Stock!$E$4="","Exchange rate not set",'
                   f'IF(Stock!$M${tot}="","Export rates not entered",Stock!$M${tot}))')
vv.font = Font(name=FONT, size=13, bold=True, color='FFFFFF')
vv.number_format = CUR_IND
vv.alignment = Alignment(horizontal='left', indent=1)
for c in (5, 6):
    dr.cell(row=KPI, column=c).fill = PatternFill('solid', fgColor=DEEP)
    dr.cell(row=KPI + 1, column=c).fill = PatternFill('solid', fgColor=DEEP)
dr.row_dimensions[KPI].height = 13
dr.row_dimensions[KPI + 1].height = 22

dr.merge_cells(start_row=KPI + 2, start_column=1, end_row=KPI + 2, end_column=6)
dr.cell(row=KPI + 2, column=1,
        value=f'=IF(SUM(Stock!$X${first}:$X${last})=0,"",'
              f'SUM(Stock!$X${first}:$X${last})&" product(s) have no Kg / Unit set — '
              f'their stock is not counted in these totals.")')\
  .font = Font(name=FONT, size=9, bold=True, color=RED)
dr.row_dimensions[KPI + 2].height = 12

def section(row, text):
    dr.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    c = dr.cell(row=row, column=1, value=text.upper())
    c.font = Font(name=FONT, size=9, bold=True, color=DEEP)
    c.alignment = Alignment(vertical='center')
    for col in range(1, 7):
        dr.cell(row=row, column=col).fill = PatternFill('solid', fgColor=BAND)
        dr.cell(row=row, column=col).border = Border(top=thin, bottom=thin)
    dr.row_dimensions[row].height = 17

def sub_head(row, cols, align_from=3):
    for i, t in enumerate(cols, start=1):
        c = dr.cell(row=row, column=i, value=t)
        c.font = Font(name=FONT, size=8, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=DEEP)
        c.alignment = Alignment(horizontal='right' if i >= align_from else 'left', wrap_text=True)
        c.border = box
    dr.row_dimensions[row].height = 22

BSEC = KPI + 4
section(BSEC, 'Position by brand')
sub_head(BSEC + 1, ['Brand', 'Closing Kg', "Today's Production (kg)",
                    "Today's Shipment (kg)", 'Products', 'Stock Value (₹)'], align_from=2)
brow = BSEC + 2
BLOCK = list(BRANDS) + [None] * SPARE_BLOCKS
for i, b in enumerate(BLOCK):
    r = brow + i
    c = dr.cell(row=r, column=1, value=b)
    if b is None:
        c.protection = UNLOCK
    rng = lambda col: f'Stock!$A${first}:$A${last},$A{r},Stock!${col}${first}:${col}${last}'
    for col, sc, fmt in ((2, 'K', KG_F), (3, 'T', KG_F), (4, 'U', KG_F), (6, 'M', CUR_IND)):
        body = f'SUMIF({rng(sc)})'
        cc = dr.cell(row=r, column=col,
                     value=(f'=IF($A{r}="","",IF({body}=0,"",{body}))' if col == 6
                            else f'=IF($A{r}="","",{body})'))
        cc.number_format = fmt
        cc.font = Font(name=FONT, size=9, color=INK, bold=(col == 2))
    dr.cell(row=r, column=5,
            value=f'=IF($A{r}="","",COUNTIF(Stock!$A${first}:$A${last},$A{r}))')\
      .font = Font(name=FONT, size=9, color=MUTED)
    for col in range(1, 7):
        cc = dr.cell(row=r, column=col)
        cc.border = box
        if i % 2:
            cc.fill = PatternFill('solid', fgColor=WASH)
    dr.row_dimensions[r].height = 13
blast = brow + len(BLOCK) - 1
btot = blast + 1
dr.cell(row=btot, column=1, value='GRAND TOTAL')
for col in range(1, 7):
    c = dr.cell(row=btot, column=col)
    c.fill = PatternFill('solid', fgColor=DEEP)
    c.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
    c.border = box
for col, L, fmt in ((2, 'B', KG_F), (3, 'C', KG_F), (4, 'D', KG_F), (6, 'F', CUR_IND)):
    rng2 = f'{L}{brow}:{L}{blast}'
    dr.cell(row=btot, column=col,
            value=(f'=IF(SUM({rng2})=0,"",SUM({rng2}))' if col == 6 else f'=SUM({rng2})'))\
      .number_format = fmt
    dr.cell(row=btot, column=col).font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
dr.conditional_formatting.add(
    f'A{brow}:F{blast}',
    FormulaRule(formula=[f'$A{brow}=""'], border=Border(), stopIfTrue=True))

MSEC = btot + 2
section(MSEC, "Today's movements")
sub_head(MSEC + 1, ['Brand', 'Product', 'Production', 'Shipment', 'Closing Balance', 'Unit'])
mrow = MSEC + 2
for i in range(MOVE_ROWS):
    r = mrow + i
    dr.cell(row=r, column=8,
            value=f'=IFERROR(MATCH({i+1},Stock!$Q${first}:$Q${last},0),"")')\
      .font = Font(name=FONT, size=8, color=MUTED)
    for col, sc in ((1, 'A'), (2, 'B'), (3, 'F'), (4, 'G'), (5, 'J'), (6, 'C')):
        idx = f'INDEX(Stock!${sc}${first}:${sc}${last},$H{r})'
        c = dr.cell(row=r, column=col,
                    value=(f'=IF($H{r}="","",IF({idx}=0,"",{idx}))' if col == 6
                           else f'=IF($H{r}="","",{idx})'))
        c.font = Font(name=FONT, size=9, color=INK, bold=(col == 5))
        c.number_format = N_IND if col == 5 else (N_DASH if col in (3, 4) else 'General')
        if col == 6:
            c.alignment = Alignment(horizontal='center')
    dr.row_dimensions[r].height = 13
mlast = mrow + MOVE_ROWS - 1
dr.conditional_formatting.add(
    f'A{mrow}:F{mlast}',
    FormulaRule(formula=[f'$B{mrow}<>""'], border=box, stopIfTrue=False))
mtot = mlast + 1
dr.cell(row=mtot, column=1, value='TOTAL MOVED TODAY')
for col in range(1, 7):
    c = dr.cell(row=mtot, column=col)
    c.fill = PatternFill('solid', fgColor=WASH)
    c.border = Border(top=Side(style='medium', color=DEEP), bottom=thin, left=thin, right=thin)
    c.font = Font(name=FONT, size=9, bold=True, color=INK)
dr.cell(row=mtot, column=2,
        value=f'=COUNTIF(Stock!$P${first}:$P${last},1)&" product(s)"')\
  .font = Font(name=FONT, size=9, color=MUTED)
dr.cell(row=mtot, column=3, value=f'=SUM(Stock!$T${first}:$T${last})').number_format = KG_F
dr.cell(row=mtot, column=4, value=f'=SUM(Stock!$U${first}:$U${last})').number_format = KG_F
dr.cell(row=mtot, column=5, value='kg').alignment = Alignment(horizontal='right')
dr.cell(row=mtot, column=5).font = Font(name=FONT, size=8, color=MUTED)
dr.row_dimensions[mtot].height = 15

ovf = mtot + 1
dr.merge_cells(start_row=ovf, start_column=1, end_row=ovf, end_column=6)
dr.cell(row=ovf, column=1,
        value=f'=IF(COUNTIF(Stock!$P${first}:$P${last},1)>{MOVE_ROWS},'
              f'"Note: "&(COUNTIF(Stock!$P${first}:$P${last},1)-{MOVE_ROWS})&'
              f'" further product(s) moved today — see the Stock sheet.","")')\
  .font = Font(name=FONT, size=8, bold=True, color=RED)

foot = ovf + 2
dr.merge_cells(start_row=foot, start_column=1, end_row=foot, end_column=6)
dr.cell(row=foot, column=1,
        value='Prepared from the stock register. Balances are in each product\'s own unit; '
              'totals are in kilos. Value = Closing Kg × Rate / Kg (US$) × exchange rate.')\
  .font = Font(name=FONT, size=8, italic=True, color=MUTED)
for col in range(1, 7):
    dr.cell(row=foot, column=col).border = Border(top=thin)

dr.print_area = f'A1:F{foot}'
dr.page_setup.orientation = 'portrait'
dr.page_setup.fitToWidth = 1
dr.page_setup.fitToHeight = 1
dr.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
dr.page_margins.left = dr.page_margins.right = 0.4
dr.page_margins.top = dr.page_margins.bottom = 0.4

# ================================================================ Read me
rm = wb['Sheet']
rm.title = 'Read me'
wb.move_sheet('Read me', offset=-4)
title_block(rm, 'STOCK REGISTER — HOW IT WORKS')
rm.column_dimensions['A'].width = 2
rm.column_dimensions['B'].width = 112
lines = [
    ('h', 'Every day, in one minute'),
    ('p', "1.  On 'Daily Entry' add one row per product that moved — date, product from the"),
    ('p', "    dropdown, and the quantity in that product's own unit."),
    ('p', "2.  On 'Stock' set the Statement date and the US$ → ₹ rate for today."),
    ('p', "3.  Open 'Daily Report' — it has already rebuilt itself."),
    ('p', '4.  File ▸ Export ▸ Create PDF/XPS, "Selected sheet" — that is the directors\' copy.'),
    ('n', ''),
    ('h', 'Units: slabs and kilos are not the same thing'),
    ('p', 'Every product carries a Unit and a Kg / Unit on the Stock sheet.'),
    ('p', '    Slab, 2       the shrimp brands — Blue Dolphin, Primus, AMF, THT, China, Deep Sea'),
    ('p', '    Kg, 1         tuna and mackerel, where the balance is already in kilos'),
    ('p', '    Slab, other   squid, cuttlefish and shrimp whole, whose slab weight varies by buyer'),
    ('p', 'Because a slab and a kilo cannot be added together, every total in this workbook is'),
    ('p', 'taken in kilos: Closing Kg = Closing Balance × Kg / Unit.'),
    ('p', 'A product with no Kg / Unit is shaded pink and left out of the totals, and the count'),
    ('p', 'of such rows is shown on Stock, Summary and the Daily Report until you fill it in.'),
    ('n', ''),
    ('h', 'How the stock is valued'),
    ('p', 'Rate / Kg is the export price per kilo, in US dollars.'),
    ('p', 'Stock Value (₹) = Closing Kg × Rate / Kg (US$) × the US$ → ₹ rate in E4.'),
    ('p', 'The exchange rate lives in that one cell, so changing it revalues everything at once.'),
    ('n', ''),
    ('h', 'Why you never re-type an opening balance'),
    ('p', 'Opening Balance is a fixed baseline, not a daily figure. Closing Balance = Opening'),
    ('p', "+ production to date − shipment to date, counted from every row on 'Daily Entry'."),
    ('n', ''),
    ('h', 'What can be typed into — everything else is locked'),
    ('b', "'Daily Entry' — the whole grid."),
    ('y', "'Stock' — Statement date, exchange rate, Unit, Kg / Unit and Rate / Kg."),
    ('k', 'Every other cell is protected, and formatting is locked, so widths, fonts, colours'),
    ('k', 'and number formats cannot be changed by accident. The owner holds the password.'),
    ('n', ''),
    ('h', 'Adding a product or a brand'),
    ('p', "'Stock' keeps 25 blank rows under the last product. Type Brand, Product, Unit,"),
    ('p', 'Kg / Unit and Opening Balance into the first blank row — no unprotecting needed.'),
    ('p', "For a new brand, also type its name into a blank brand band on 'Summary' and into a"),
    ('p', "blank row on the Daily Report's brand table. Until you do, a red line on both says"),
    ('p', 'the totals understate the stock. Both clear themselves.'),
    ('n', ''),
    ('h', 'Dates must be real dates'),
    ('p', 'Type 22/08/2026, not text. A date stored as text never matches the Statement date.'),
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
    else:
        c.font = Font(name=FONT, size=10, color=INK)
    r += 1
rm.print_area = f'A1:B{r}'
rm.page_setup.fitToWidth = 1
rm.page_setup.fitToHeight = 0
rm.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ---------------------------------------------------------------- protection
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.protection.password = PASSWORD
    ws.protection.sheet = True
    ws.protection.enable()
    for flag in ('formatCells', 'formatColumns', 'formatRows', 'insertRows',
                 'insertColumns', 'deleteRows', 'deleteColumns', 'insertHyperlinks',
                 'sort', 'pivotTables', 'objects', 'scenarios'):
        setattr(ws.protection, flag, True)      # True == locked, i.e. not allowed
    ws.protection.autoFilter = False
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False
wb.security = WorkbookProtection(workbookPassword=PASSWORD, lockStructure=True)

wb.active = wb.index(dr)
wb.save(OUT)
print(f'wrote {OUT}: {len(ROWS)} products, {len(BRANDS)} brands, {len(LOG)} log rows, '
      f'{len(SUB_ROWS)} brand tables on Summary')
