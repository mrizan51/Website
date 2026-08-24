#!/usr/bin/env python3
"""Builds PEI-Lot-Analysis.xlsx — yield and cost analysis per raw-material lot,
with a one-page Lot Report for the directors.

    python tools/build-lot-xlsx.py out.xlsx
    python tools/build-lot-xlsx.py out.xlsx --from live.xlsx
    python tools/build-lot-xlsx.py --relock out.xlsx        # after recalc

The model, read off statement C 1374:

    Kgs      = Slabs × Net kg/slab          net packed weight, what is sold
    F.WT     = Slabs × Gross wt/slab        final weight, what yield is measured on
    Amount   = Kgs × Sales price (US$/kg)
    Yield %  = total F.WT ÷ raw kg          note: F.WT, not net kg
    Raw cost = purchase amount + peeling rate × raw kg
    Expenses = expense rate ₹/kg × total net kg
    Profit   = Amount × exchange rate − raw cost − expenses
    Per kg   = profit ÷ RAW kg, not finished kg

Two weights per slab matter and are not the same: net (2 kg, what the buyer
pays for) and gross/F.WT (2.15, 2.20 …, what the yield is calculated on).
"""
import sys, os, json, datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.drawing.image import Image as XLImage
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO = os.path.join(ROOT, 'stock-statement', 'assets', 'pei-mark.png')
DATA = os.path.join(ROOT, 'tools', 'lot-data.json')

args = [a for a in sys.argv[1:] if not a.startswith('--')]
OUT = args[0] if args else 'PEI-Lot-Analysis.xlsx'
SRC = sys.argv[sys.argv.index('--from') + 1] if '--from' in sys.argv else None
PASSWORD = (sys.argv[sys.argv.index('--password') + 1]
            if '--password' in sys.argv else 'PEI2026')
UNLOCK = Protection(locked=False)

LOT_ROWS = 200          # lots the register holds
LINE_ROWS = 2000        # grade lines across all lots
LINE_FMT = 400          # of those, pre-ruled
GRADE_ROWS = 40         # master grade list, incl. spares
REPORT_LINES = 12       # grade lines the report prints; more triggers a note

# ---------------------------------------------------------------- relock
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
    targets = {}
    for m in _re.finditer(r'<sheet\b[^>]*/>', wbx):
        a = dict(_re.findall(r'(?:r:)?(\w+)="([^"]*)"', m.group(0)))
        if a.get('name') in ('Lot Entry', 'Lots'):
            targets[a['name']] = 'xl/' + rid.get(a.get('id', ''), '').lstrip('/')
    st_xml = blob['xl/styles.xml'].decode('utf-8')
    head, rest = st_xml.split('<cellXfs', 1)
    attrs, body_and_tail = rest.split('>', 1)
    body, tail = body_and_tail.split('</cellXfs>', 1)
    xfs = _re.findall(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', body, _re.S)
    added = []

    def unlocked_copy(xf):
        if '<protection' in xf:
            xf = _re.sub(r'<protection\b[^>]*/>', '<protection locked="0" hidden="0"/>', xf, count=1)
            o = xf.split('>', 1)[0]
            if 'applyProtection' not in o:
                xf = xf.replace(o, o + ' applyProtection="1"', 1)
            return xf
        if xf.endswith('/>'):
            b = xf[:-2]
            if 'applyProtection' not in b:
                b += ' applyProtection="1"'
            return b + '><protection locked="0" hidden="0"/></xf>'
        b = xf[:-len('</xf>')]
        if 'applyProtection' not in b.split('>', 1)[0]:
            o, i2 = b.split('>', 1)
            b = o + ' applyProtection="1">' + i2
        return b + '<protection locked="0" hidden="0"/></xf>'

    for name, f in targets.items():
        if f not in blob:
            continue
        sx = blob[f].decode('utf-8')
        limit = 7 if name == 'Lot Entry' else 21    # entry columns on each sheet

        def repoint(m, limit=limit):
            c = m.group(0)
            a = dict(_re.findall(r'(\w+)="([^"]*)"', c))
            lo = int(a.get('min', 0))
            if lo > limit:
                return c
            nid = len(xfs) + len(added)
            added.append(unlocked_copy(xfs[int(a.get('style', 0))]))
            if 'style=' in c:
                return _re.sub(r'style="\d+"', f'style="{nid}"', c)
            return c[:-2] + f' style="{nid}"/>'

        sx = _re.sub(r'<col\b[^>]*/>', repoint, sx)
        top = 7 if name == 'Lot Entry' else 8
        sx = _re.sub(rf'sqref="([A-Z]+){top}:([A-Z]+)(\d+)"',
                     lambda m: f'sqref="{m.group(1)}{top}:{m.group(2)}'
                               f'{LINE_ROWS if name == "Lot Entry" else LOT_ROWS}"', sx)
        blob[f] = sx.encode('utf-8')
    if added:
        body += ''.join(added)
        attrs = _re.sub(r'count="\d+"', f'count="{len(xfs) + len(added)}"', attrs)
        blob['xl/styles.xml'] = (head + '<cellXfs' + attrs + '>' + body +
                                 '</cellXfs>' + tail).encode('utf-8')
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
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    for f in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.datetime.strptime(str(v).strip(), f).date()
        except (ValueError, TypeError):
            pass
    return None

seed = json.load(open(DATA))
GRADES = [tuple(g) for g in seed['grades']]
EXPENSES = [tuple(e) for e in seed['expenses']]
LOTS = seed['lots']
LINES = [tuple(l) for l in seed['lines']]

if SRC:
    src = load_workbook(SRC)
    s_lo, s_li = src['Lots'], src['Lot Entry']
    LOTS = []
    for r in range(8, LOT_ROWS + 1):
        if not s_lo[f'A{r}'].value:
            continue
        LOTS.append({'lot': s_lo[f'A{r}'].value,
                     'statement_date': str(as_date(s_lo[f'B{r}'].value) or ''),
                     'purchase_date': str(as_date(s_lo[f'C{r}'].value) or ''),
                     'party': s_lo[f'D{r}'].value or '',
                     'supervisor': s_lo[f'E{r}'].value or '',
                     'raw_kg': s_lo[f'F{r}'].value or 0,
                     'raw_amount': s_lo[f'G{r}'].value or 0,
                     'peeling_rate': s_lo[f'H{r}'].value or 0,
                     'fx': s_lo[f'I{r}'].value or 0})
    LINES = []
    for r in range(7, LINE_ROWS + 1):
        if not s_li[f'A{r}'].value:
            continue
        LINES.append((s_li[f'A{r}'].value, s_li[f'B{r}'].value, s_li[f'C{r}'].value,
                      s_li[f'D{r}'].value or 0, s_li[f'E{r}'].value or 2,
                      s_li[f'F{r}'].value or 0, s_li[f'G{r}'].value or 0))
    if 'Grades' in src.sheetnames:
        g = src['Grades']
        GRADES = [(g[f'A{r}'].value, g[f'B{r}'].value, g[f'C{r}'].value, g[f'D{r}'].value)
                  for r in range(6, GRADE_ROWS + 6) if g[f'A{r}'].value]
    if 'Expenses' in src.sheetnames:
        e = src['Expenses']
        EXPENSES = [(e[f'A{r}'].value, e[f'B{r}'].value)
                    for r in range(6, 40) if e[f'A{r}'].value and e[f'A{r}'].value != 'TOTAL']
    print(f'read {SRC}: {len(LOTS)} lots, {len(LINES)} grade lines')

# ---------------------------------------------------------------- style
FONT = 'Arial'
DEEP, BAND, WASH, LINE_C = '065E7A', 'E7F3F9', 'F2F8FB', 'CBDFE8'
INK, MUTED, BRAND_C, RED, GREEN = '0B2C3A', '5E7C8B', '00A2D3', 'B3261E', '1B7A43'
BLUE_IN, YELLOW, GREY = '0000FF', 'FFFF00', 'D9D9D9'
thin = Side(style='thin', color=LINE_C)
box = Border(left=thin, right=thin, top=thin, bottom=thin)

N0 = r'#,##0'
N2 = r'#,##0.00'
N3 = r'#,##0.000'
KG = r'[>=100000]##\,##\,##0.0;#,##0.0'
INR = r'[>=10000000]"₹" ##\,##\,##\,##0.00;[>=100000]"₹" ##\,##\,##0.00;"₹" #,##0.00'
INR0 = r'[>=10000000]"₹" ##\,##\,##\,##0;[>=100000]"₹" ##\,##\,##0;"₹" #,##0'
USD = r'"$" #,##0.00'
PCT = r'0.00%'
DATE_F = 'DD/MM/YYYY'

wb = Workbook()

def head_row(ws, row, cols, widths=None, align_from=3, size=9, height=28):
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

def title(ws, sub):
    ws['A1'] = 'PREMIER EXPORTS INTERNATIONAL'
    ws['A1'].font = Font(name=FONT, size=14, bold=True, color=DEEP)
    ws['A2'] = sub
    ws['A2'].font = Font(name=FONT, size=10, bold=True, color=MUTED)

# ================================================================ Grades
gr = wb.create_sheet('Grades')
title(gr, 'GRADE MASTER — SPEC, GRADE AND THE USUAL WEIGHT AND PRICE')
gr['A4'] = ('The dropdown on Lot Entry reads this list. Gross wt and price here are the '
            'usual figures — enter the actual ones on each lot line.')
gr['A4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)
GH = 5
head_row(gr, GH, ['Spec', 'Grade', 'Gross wt / slab (kg)', 'Usual price (US$/kg)', '',
                  'Key (do not edit)'], [20, 18, 16, 17, 2, 30], align_from=3)
gr.cell(row=GH, column=6).fill = PatternFill('solid', fgColor=GREY)
gr.cell(row=GH, column=6).font = Font(name=FONT, size=9, bold=True, color=MUTED)
g0 = GH + 1
for i in range(GRADE_ROWS):
    r = g0 + i
    src = GRADES[i] if i < len(GRADES) else (None, None, None, None)
    gr.cell(row=r, column=1, value=src[0])
    gr.cell(row=r, column=2, value=src[1])
    gr.cell(row=r, column=3, value=src[2]).number_format = N3
    gr.cell(row=r, column=4, value=src[3]).number_format = USD
    gr.cell(row=r, column=6, value=f'=IF($A{r}="","",$A{r}&"  "&$B{r})')\
      .font = Font(name=FONT, size=9, color=MUTED)
    for col in range(1, 5):
        c = gr.cell(row=r, column=col)
        c.border = box
        c.font = Font(name=FONT, size=10, color=BLUE_IN)
        c.protection = UNLOCK
        if i % 2:
            c.fill = PatternFill('solid', fgColor=WASH)
g1 = g0 + GRADE_ROWS - 1
gr.print_area = f'A1:D{g1}'
gr.page_setup.fitToWidth = 1
gr.page_setup.fitToHeight = 1      # the master is a one-page reference, not a run of rows
gr.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

wb.defined_names.add(DefinedName('SpecList', attr_text=f'OFFSET(Grades!$A${g0},0,0,'
                                                      f'MAX(1,COUNTA(Grades!$A${g0}:$A${g1})),1)'))
wb.defined_names.add(DefinedName('GradeList', attr_text=f'OFFSET(Grades!$B${g0},0,0,'
                                                        f'MAX(1,COUNTA(Grades!$A${g0}:$A${g1})),1)'))

# ================================================================ Expenses
ex = wb.create_sheet('Expenses')
title(ex, 'EXPENSE HEADS — STANDARD RATE PER KG OF PACKED PRODUCT')
ex['A4'] = ('These rates apply to every lot. A lot that differs can override any head on '
            'the Lots sheet without changing them here.')
ex['A4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)
EH = 5
head_row(ex, EH, ['Expense head', 'Rate ₹ / kg'], [30, 14], align_from=2)
e0 = EH + 1
for i, (name, rate) in enumerate(EXPENSES):
    r = e0 + i
    ex.cell(row=r, column=1, value=name)
    c = ex.cell(row=r, column=2, value=rate)
    c.number_format = N2
    c.font = Font(name=FONT, size=10, color=BLUE_IN)
    c.protection = UNLOCK
    c.fill = PatternFill('solid', fgColor=YELLOW)
    for col in (1, 2):
        ex.cell(row=r, column=col).border = box
e1 = e0 + len(EXPENSES) - 1
etot = e1 + 1
ex.cell(row=etot, column=1, value='TOTAL')
ex.cell(row=etot, column=2, value=f'=SUM(B{e0}:B{e1})').number_format = N2
for col in (1, 2):
    c = ex.cell(row=etot, column=col)
    c.fill = PatternFill('solid', fgColor=DEEP)
    c.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
    c.border = box
ex.print_area = f'A1:B{etot}'
ex.page_setup.fitToWidth = 1
ex.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Lots
lo = wb.create_sheet('Lots')
title(lo, 'LOTS — ONE ROW PER RAW-MATERIAL PURCHASE')
lo['A4'] = ('Type the eight blue columns. Everything from Peeling Amount rightwards is '
            'calculated. Leave an expense override blank to use the standard rate.')
lo['A4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)
LH = 7
base = ['Lot No', 'Statement Date', 'Purchase Date', 'Party', 'Supervisor',
        'Raw Qty (kg)', 'Raw Amount (₹)', 'Peeling ₹/kg', 'Exch. rate US$→₹']
calc = ['Peeling Amount (₹)', 'Total Raw Cost (₹)', 'Avg Price /kg (₹)',
        'Packed (kg)', 'F.WT (kg)', 'Yield %', 'Sales (US$)', 'Sales Value (₹)',
        'Expense ₹/kg', 'Expenses (₹)', 'Total Cost (₹)', 'Profit (₹)', 'Profit / raw kg (₹)']
ovr = [f'{n} ovr' for n, _ in EXPENSES]
head_row(lo, LH, base + calc + ovr,
         [12, 14, 14, 26, 14, 12, 15, 11, 13] + [14] * len(calc) + [11] * len(ovr),
         align_from=6)
OVR0 = 10 + len(calc)          # first override column index
for i, c in enumerate(range(OVR0, OVR0 + len(ovr))):
    cc = lo.cell(row=LH, column=c)
    cc.fill = PatternFill('solid', fgColor=GREY)
    cc.font = Font(name=FONT, size=8, bold=True, color=MUTED)

l0 = LH + 1
for i in range(LOT_ROWS - LH):
    r = l0 + i
    d = LOTS[i] if i < len(LOTS) else None
    lo.cell(row=r, column=1, value=d['lot'] if d else None)
    lo.cell(row=r, column=2, value=as_date(d['statement_date']) if d else None).number_format = DATE_F
    lo.cell(row=r, column=3, value=as_date(d['purchase_date']) if d else None).number_format = DATE_F
    lo.cell(row=r, column=4, value=d['party'] if d else None)
    lo.cell(row=r, column=5, value=d['supervisor'] if d else None)
    lo.cell(row=r, column=6, value=d['raw_kg'] if d else None).number_format = N0
    lo.cell(row=r, column=7, value=d['raw_amount'] if d else None).number_format = INR
    lo.cell(row=r, column=8, value=d['peeling_rate'] if d else None).number_format = N2
    lo.cell(row=r, column=9, value=d['fx'] if d else None).number_format = N2
    blank = f'$A{r}=""'
    li = lambda col: f"'Lot Entry'!${col}$7:${col}${LINE_ROWS}"
    rate_terms = '+'.join(
        f'IF({get_column_letter(OVR0 + k)}{r}="",Expenses!$B${e0 + k},'
        f'{get_column_letter(OVR0 + k)}{r})' for k in range(len(EXPENSES)))
    F = [
        (10, f'=IF({blank},"",$H{r}*$F{r})', INR),
        (11, f'=IF({blank},"",$G{r}+$J{r})', INR),
        (12, f'=IF(OR({blank},$F{r}=0),"",$K{r}/$F{r})', N2),
        (13, f'=IF({blank},"",SUMIF({li("A")},$A{r},{li("H")}))', KG),
        (14, f'=IF({blank},"",SUMIF({li("A")},$A{r},{li("I")}))', N3),
        (15, f'=IF(OR({blank},$F{r}=0),"",$N{r}/$F{r})', PCT),
        (16, f'=IF({blank},"",SUMIF({li("A")},$A{r},{li("J")}))', USD),
        (17, f'=IF({blank},"",$P{r}*$I{r})', INR),
        (18, f'=IF({blank},"",{rate_terms})', N2),
        (19, f'=IF({blank},"",$R{r}*$M{r})', INR),
        (20, f'=IF({blank},"",$K{r}+$S{r})', INR),
        (21, f'=IF({blank},"",$Q{r}-$T{r})', INR),
        (22, f'=IF(OR({blank},$F{r}=0),"",$U{r}/$F{r})', N2),
    ]
    for col, f, fmt in F:
        c = lo.cell(row=r, column=col, value=f)
        c.number_format = fmt
        c.font = Font(name=FONT, size=10, color=INK, bold=(col in (21, 22)))
    for k in range(len(EXPENSES)):
        c = lo.cell(row=r, column=OVR0 + k)
        c.number_format = N2
        c.font = Font(name=FONT, size=9, color=BLUE_IN)
        c.protection = UNLOCK
        c.border = box
    for col in range(1, 23):
        c = lo.cell(row=r, column=col)
        c.border = box
        if col <= 9:
            c.font = Font(name=FONT, size=10, color=BLUE_IN)
            c.protection = UNLOCK
            c.fill = PatternFill('solid', fgColor=YELLOW if col in (6, 7, 8, 9) else 'FFFFFF')
        elif i % 2:
            c.fill = PatternFill('solid', fgColor=WASH)
l1 = l0 + (LOT_ROWS - LH) - 1
lo.conditional_formatting.add(
    f'U{l0}:U{l1}',
    FormulaRule(formula=[f'AND($A{l0}<>"",$U{l0}<0)'],
                fill=PatternFill('solid', fgColor='F8D7DA'),
                font=Font(name=FONT, size=10, bold=True, color=RED), stopIfTrue=False))
lo.freeze_panes = f'B{l0}'
lo.auto_filter.ref = f'A{LH}:V{l1}'
lo.print_area = f'A1:V{l1}'
lo.print_title_rows = f'{LH}:{LH}'
lo.page_setup.orientation = 'landscape'
lo.page_setup.fitToWidth = 1
lo.page_setup.fitToHeight = 0
lo.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
wb.defined_names.add(DefinedName('LotList', attr_text=f'OFFSET(Lots!$A${l0},0,0,'
                                                      f'MAX(1,COUNTA(Lots!$A${l0}:$A${l1})),1)'))

# ================================================================ Lot Entry
li = wb.create_sheet('Lot Entry')
title(li, 'LOT ENTRY — ONE ROW PER GRADE PRODUCED FROM A LOT')
li['A4'] = ('Pick the lot and the spec/grade, then enter slabs, the net kg per slab (2 unless '
            'the buyer differs), the gross weight per slab and the sale price per kg.')
li['A4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)
IH = 6
head_row(li, IH, ['Lot No', 'Spec', 'Grade', 'Slabs', 'Net kg / slab', 'Gross wt / slab',
                  'Price (US$/kg)', 'Packed kg', 'F.WT (kg)', 'Amount (US$)', '',
                  'Lot seq (do not edit)'],
         [12, 20, 18, 10, 12, 14, 13, 12, 12, 14, 2, 22], align_from=4)
for c in (11, 12):
    cc = li.cell(row=IH, column=c)
    cc.fill = PatternFill('solid', fgColor=GREY if c == 12 else 'FFFFFF')
    cc.font = Font(name=FONT, size=9, bold=True, color=MUTED)
i0 = IH + 1
for i in range(LINE_FMT - IH):
    r = i0 + i
    d = LINES[i] if i < len(LINES) else (None,) * 7
    for col, v in enumerate(d, start=1):
        li.cell(row=r, column=col, value=v)
    li.cell(row=r, column=4).number_format = N0
    li.cell(row=r, column=5).number_format = N3
    li.cell(row=r, column=6).number_format = N3
    li.cell(row=r, column=7).number_format = USD
    blank = f'$A{r}=""'
    li.cell(row=r, column=8, value=f'=IF({blank},"",$D{r}*$E{r})').number_format = N0
    li.cell(row=r, column=9, value=f'=IF({blank},"",$D{r}*$F{r})').number_format = N3
    li.cell(row=r, column=10, value=f'=IF({blank},"",$H{r}*$G{r})').number_format = USD
    li.cell(row=r, column=12,
            value=f'=IF({blank},"",$A{r}&"#"&COUNTIF($A${i0}:$A{r},$A{r}))')\
      .font = Font(name=FONT, size=9, color=MUTED)
    for col in range(1, 11):
        c = li.cell(row=r, column=col)
        c.border = box
        c.font = Font(name=FONT, size=10, color=BLUE_IN if col <= 7 else INK)
        if col <= 7:
            c.protection = UNLOCK
        elif i % 2:
            c.fill = PatternFill('solid', fgColor=WASH)
for col, fmt in (('A', 'General'), ('B', 'General'), ('C', 'General'), ('D', N0),
                 ('E', N3), ('F', N3), ('G', USD)):
    cd = li.column_dimensions[col]
    cd.protection = Protection(locked=False)
    cd.font = Font(name=FONT, size=10, color=BLUE_IN)
    cd.number_format = fmt
i1 = LINE_FMT
for name, f1, rng in (('lot', '=LotList', f'A{i0}:A{LINE_ROWS}'),
                      ('spec', '=SpecList', f'B{i0}:B{LINE_ROWS}'),
                      ('grade', '=GradeList', f'C{i0}:C{LINE_ROWS}')):
    dv = DataValidation(type='list', formula1=f1, allow_blank=True)
    dv.error = f'Pick a {name} from the list.'
    dv.errorTitle = name.title()
    li.add_data_validation(dv)
    dv.add(rng)
li['N4'] = 'Lines whose lot is not on the Lots sheet:'
li['N4'].font = Font(name=FONT, size=9, bold=True, color=MUTED)
li['R4'] = (f'=SUMPRODUCT((A{i0}:A{LINE_ROWS}<>"")*'
            f'(COUNTIF(Lots!$A${l0}:$A${l1},A{i0}:A{LINE_ROWS})=0))')
li['R4'].font = Font(name=FONT, size=10, bold=True, color=RED)
li['S4'] = '← must stay at 0'
li['S4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)
li.column_dimensions['N'].width = 30
li.column_dimensions['R'].width = 6
li.freeze_panes = f'A{i0}'
li.auto_filter.ref = f'A{IH}:J{i1}'
li.print_area = f'A1:J{i1}'
li.print_title_rows = f'{IH}:{IH}'
li.page_setup.orientation = 'landscape'
li.page_setup.fitToWidth = 1
li.page_setup.fitToHeight = 0
li.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Lot Report
rp = wb.create_sheet('Lot Report')
for col, w in zip('ABCDEFG', (19, 17, 11, 12, 13, 11, 15)):
    rp.column_dimensions[col].width = w
rp.column_dimensions['H'].width = 6
rp.row_dimensions[1].height = 34
if os.path.exists(LOGO):
    im = XLImage(LOGO)
    im.height, im.width = 42, 52
    rp.add_image(im, 'A1')
rp['B1'] = 'Premier Exports International'
rp['B1'].font = Font(name=FONT, size=16, color=BRAND_C)
rp['B1'].alignment = Alignment(vertical='center')
for i, t in enumerate(['AP X/453, NH-66 Highway, Chandiroor P.O.,',
                       'Aroor, Alappuzha, Kerala - 688 537, India',
                       'GSTIN 32AADFP3158P1ZZ']):
    rp.merge_cells(start_row=1 + i, start_column=5, end_row=1 + i, end_column=7)
    c = rp.cell(row=1 + i, column=5, value=t)
    c.font = Font(name=FONT, size=8, color=MUTED)
    c.alignment = Alignment(horizontal='right', vertical='center')
rp.row_dimensions[4].height = 5
for col in range(1, 8):
    rp.cell(row=4, column=col).border = Border(bottom=Side(style='medium', color=BRAND_C))
rp.row_dimensions[5].height = 24
rp['A5'] = 'PURCHASE & PROCESSING REPORT'
rp['A5'].font = Font(name=FONT, size=15, bold=True, color=INK)
rp['A5'].alignment = Alignment(vertical='center')
rp['A6'] = 'YIELD AND COST ANALYSIS BY LOT'
rp['A6'].font = Font(name=FONT, size=8, bold=True, color=MUTED)
rp.merge_cells('F5:G5')
rp['F5'] = LOTS[0]['lot'] if LOTS else None
rp['F5'].font = Font(name=FONT, size=14, bold=True, color=BLUE_IN)
rp['F5'].fill = PatternFill('solid', fgColor=YELLOW)
rp['F5'].alignment = Alignment(horizontal='center', vertical='center')
rp['F5'].border = box
rp['F5'].protection = UNLOCK
rp.merge_cells('F6:G6')
rp['F6'] = 'LOT NO — TYPE OR PICK'
rp['F6'].font = Font(name=FONT, size=8, bold=True, color=MUTED)
rp['F6'].alignment = Alignment(horizontal='right')
dvl = DataValidation(type='list', formula1='=LotList', allow_blank=True)
rp.add_data_validation(dvl)
dvl.add('F5')

rp['H1'] = f'=IFERROR(MATCH($F$5,Lots!$A${l0}:$A${l1},0),"")'
rp['H1'].font = Font(name=FONT, size=8, color=MUTED)
LOTREF = lambda col: f'INDEX(Lots!${col}${l0}:${col}${l1},$H$1)'
def lot(col, text=False):
    inner = LOTREF(col)
    if text:      # INDEX on an empty cell yields 0, not blank
        return f'=IF($H$1="","",IF({inner}=0,"",{inner}))'
    return f'=IF($H$1="","",{inner})'

PB = 8
pairs = [('Purchase date', lot('C'), DATE_F), ('Party', lot('D', text=True), 'General'),
         ('Supervisor', lot('E', text=True), 'General'), ('Statement date', lot('B'), DATE_F)]
for i, (lab, f, fmt) in enumerate(pairs):
    r = PB + i
    k = rp.cell(row=r, column=1, value=lab)
    k.font = Font(name=FONT, size=9, bold=True, color=MUTED)
    v = rp.cell(row=r, column=2, value=f)
    v.number_format = fmt
    v.font = Font(name=FONT, size=9, color=INK)
    rp.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    rp.row_dimensions[r].height = 13
money = [('Purchases (kg)', lot('F'), N0), ('Purchases (₹)', lot('G'), INR),
         ('Peeling @ ₹/kg', lot('H'), N2), ('Peeling (₹)', lot('J'), INR),
         ('Total raw cost (₹)', lot('K'), INR), ('Average price / kg (₹)', lot('L'), N2)]
for i, (lab, f, fmt) in enumerate(money):
    r = PB + i
    k = rp.cell(row=r, column=5, value=lab)
    k.font = Font(name=FONT, size=9, bold=True, color=MUTED)
    k.alignment = Alignment(horizontal='right')
    rp.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    v = rp.cell(row=r, column=7, value=f)
    v.number_format = fmt
    v.font = Font(name=FONT, size=9, bold=True, color=INK)

KP = PB + 7
# Four tiles have to land inside the report's own width (A:G) — a tile merged
# past column G bleeds its fill onto the right margin when the sheet prints.
kpis = [('Yield', lot('O'), PCT, 1, 1), ('Net profit / raw kg', lot('V'), N2, 2, 3),
        ('Sales value (₹)', lot('Q'), INR0, 4, 5), ('Net profit (₹)', lot('U'), INR0, 6, 7)]
for i, (lab, f, fmt, c0, c1) in enumerate(kpis):
    if c1 > c0:
        rp.merge_cells(start_row=KP, start_column=c0, end_row=KP, end_column=c1)
        rp.merge_cells(start_row=KP + 1, start_column=c0, end_row=KP + 1, end_column=c1)
    accent = (i == 3)
    k = rp.cell(row=KP, column=c0, value=lab.upper())
    k.font = Font(name=FONT, size=8, bold=True, color='CFE8F2' if accent else MUTED)
    k.alignment = Alignment(horizontal='left', indent=1)
    v = rp.cell(row=KP + 1, column=c0, value=f)
    v.number_format = fmt
    v.font = Font(name=FONT, size=13, bold=True, color='FFFFFF' if accent else INK)
    v.alignment = Alignment(horizontal='left', indent=1)
    for cc in range(c0, c1 + 1):
        rp.cell(row=KP, column=cc).fill = PatternFill('solid', fgColor=DEEP if accent else WASH)
        rp.cell(row=KP + 1, column=cc).fill = PatternFill('solid', fgColor=DEEP if accent else WASH)
        rp.cell(row=KP, column=cc).border = Border(left=thin, right=thin, top=thin)
        rp.cell(row=KP + 1, column=cc).border = Border(left=thin, right=thin, bottom=thin)
rp.row_dimensions[KP].height = 13
rp.row_dimensions[KP + 1].height = 22

def section(row, text):
    rp.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    c = rp.cell(row=row, column=1, value=text.upper())
    c.font = Font(name=FONT, size=9, bold=True, color=DEEP)
    c.alignment = Alignment(vertical='center')
    for col in range(1, 8):
        rp.cell(row=row, column=col).fill = PatternFill('solid', fgColor=BAND)
        rp.cell(row=row, column=col).border = Border(top=thin, bottom=thin)
    rp.row_dimensions[row].height = 16

def sub_head(row, cols, align_from=3):
    for i, t in enumerate(cols, start=1):
        c = rp.cell(row=row, column=i, value=t)
        c.font = Font(name=FONT, size=8, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=DEEP)
        c.alignment = Alignment(horizontal='right' if i >= align_from else 'left', wrap_text=True)
        c.border = box
    rp.row_dimensions[row].height = 20

GS = KP + 3
section(GS, 'Production by grade')
sub_head(GS + 1, ['Spec', 'Grade', 'Slabs', 'Packed kg', 'F.WT (kg)', 'US$/kg', 'Amount (US$)'])
gp0 = GS + 2
for i in range(REPORT_LINES):
    r = gp0 + i
    rp.cell(row=r, column=8,
            value=f'=IFERROR(MATCH($F$5&"#"&{i+1},\'Lot Entry\'!$L${i0}:$L${LINE_ROWS},0),"")')\
      .font = Font(name=FONT, size=8, color=MUTED)
    for col, sc, fmt in ((1, 'B', 'General'), (2, 'C', 'General'), (3, 'D', N0),
                         (4, 'H', N0), (5, 'I', N3), (6, 'G', USD), (7, 'J', USD)):
        c = rp.cell(row=r, column=col,
                    value=f'=IF($H{r}="","",INDEX(\'Lot Entry\'!${sc}${i0}:${sc}${LINE_ROWS},$H{r}))')
        c.number_format = fmt
        c.font = Font(name=FONT, size=9, color=INK, bold=(col == 7))
    rp.row_dimensions[r].height = 13
gp1 = gp0 + REPORT_LINES - 1
rp.conditional_formatting.add(f'A{gp0}:G{gp1}',
    FormulaRule(formula=[f'$A{gp0}<>""'], border=box, stopIfTrue=False))
gtot = gp1 + 1
rp.cell(row=gtot, column=1, value='TOTAL')
for col in range(1, 8):
    c = rp.cell(row=gtot, column=col)
    c.fill = PatternFill('solid', fgColor=WASH)
    c.border = Border(top=Side(style='medium', color=DEEP), bottom=thin, left=thin, right=thin)
    c.font = Font(name=FONT, size=9, bold=True, color=INK)
for col, L, fmt in ((3, 'C', N0), (4, 'D', N0), (5, 'E', N3), (7, 'G', USD)):
    rp.cell(row=gtot, column=col, value=f'=SUM({L}{gp0}:{L}{gp1})').number_format = fmt
    rp.cell(row=gtot, column=col).font = Font(name=FONT, size=9, bold=True, color=INK)
ovf = gtot + 1
rp.merge_cells(start_row=ovf, start_column=1, end_row=ovf, end_column=7)
rp.cell(row=ovf, column=1,
        value=f'=IF($H$1="","",IF(COUNTIF(\'Lot Entry\'!$A${i0}:$A${LINE_ROWS},$F$5)>{REPORT_LINES},'
              f'"Note: "&(COUNTIF(\'Lot Entry\'!$A${i0}:$A${LINE_ROWS},$F$5)-{REPORT_LINES})&'
              f'" further grade line(s) — see Lot Entry.",""))')\
  .font = Font(name=FONT, size=8, bold=True, color=RED)

ES = ovf + 1
section(ES, 'Expenses and result')
er0 = ES + 1
sub_head(er0, ['Expense head', '', 'Rate ₹/kg', 'Amount (₹)', '', '', ''], align_from=3)
for i, (name, _) in enumerate(EXPENSES):
    r = er0 + 1 + i
    rp.cell(row=r, column=1, value=name).font = Font(name=FONT, size=9, color=INK)
    rp.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ovc = get_column_letter(OVR0 + i)
    rate = (f'=IF($H$1="","",IF(INDEX(Lots!${ovc}${l0}:${ovc}${l1},$H$1)="",'
            f'Expenses!$B${e0 + i},INDEX(Lots!${ovc}${l0}:${ovc}${l1},$H$1)))')
    rp.cell(row=r, column=3, value=rate).number_format = N2
    rp.cell(row=r, column=4,
            value=f'=IF($H$1="","",$C{r}*{LOTREF("M")})').number_format = INR
    for col in range(1, 5):
        c = rp.cell(row=r, column=col)
        c.border = box
        c.font = Font(name=FONT, size=9, color=INK)
        if i % 2:
            c.fill = PatternFill('solid', fgColor=WASH)
    rp.row_dimensions[r].height = 12
elast = er0 + len(EXPENSES)
esum = elast + 1
rp.cell(row=esum, column=1, value='TOTAL EXPENSES')
rp.merge_cells(start_row=esum, start_column=1, end_row=esum, end_column=2)
rp.cell(row=esum, column=3, value=f'=SUM(C{er0+1}:C{elast})').number_format = N2
rp.cell(row=esum, column=4, value=f'=SUM(D{er0+1}:D{elast})').number_format = INR
for col in range(1, 5):
    c = rp.cell(row=esum, column=col)
    c.fill = PatternFill('solid', fgColor=WASH)
    c.border = Border(top=Side(style='medium', color=DEEP), bottom=thin, left=thin, right=thin)
    c.font = Font(name=FONT, size=9, bold=True, color=INK)

res = [('Sales value (₹)', lot('Q'), INR), ('Less: total raw cost (₹)', lot('K'), INR),
       ('Less: expenses (₹)', lot('S'), INR), ('Total cost (₹)', lot('T'), INR),
       ('NET PROFIT (₹)', lot('U'), INR), ('Net profit / raw kg (₹)', lot('V'), N2)]
for i, (lab, f, fmt) in enumerate(res):
    r = er0 + 1 + i
    last = (i == 4)
    k = rp.cell(row=r, column=6, value=lab)
    k.font = Font(name=FONT, size=9, bold=last, color='FFFFFF' if last else MUTED)
    k.alignment = Alignment(horizontal='right')
    v = rp.cell(row=r, column=7, value=f)
    v.number_format = fmt
    v.font = Font(name=FONT, size=10 if last else 9, bold=True,
                  color='FFFFFF' if last else INK)
    for col in (6, 7):
        c = rp.cell(row=r, column=col)
        c.border = box
        if last:
            c.fill = PatternFill('solid', fgColor=DEEP)
        elif i % 2:
            c.fill = PatternFill('solid', fgColor=WASH)

foot = max(esum, er0 + len(res)) + 2
rp.merge_cells(start_row=foot, start_column=1, end_row=foot, end_column=7)
rp.cell(row=foot, column=1,
        value='Yield is F.WT ÷ raw kg. Amount is packed kg × price. Expenses are charged on '
              'packed kg. Net profit per kg is over the raw kg purchased.')\
  .font = Font(name=FONT, size=8, italic=True, color=MUTED)
for col in range(1, 8):
    rp.cell(row=foot, column=col).border = Border(top=thin)
miss = foot + 1
rp.merge_cells(start_row=miss, start_column=1, end_row=miss, end_column=7)
rp.cell(row=miss, column=1,
        value='=IF($H$1="","No lot of that number is on the Lots sheet — check cell F5.","")')\
  .font = Font(name=FONT, size=9, bold=True, color=RED)

rp.print_area = f'A1:G{miss}'
rp.page_setup.orientation = 'portrait'
rp.page_setup.fitToWidth = 1
rp.page_setup.fitToHeight = 1
rp.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
rp.page_margins.left = rp.page_margins.right = 0.4
rp.page_margins.top = rp.page_margins.bottom = 0.4

# ================================================================ Summary
su = wb.create_sheet('Summary')
title(su, 'ALL LOTS — YIELD AND PROFIT')
for col, w in zip('ABCDEFGHI', (13, 14, 26, 12, 11, 14, 15, 15, 14)):
    su.column_dimensions[col].width = w
SH = 6
head_row(su, SH, ['Lot No', 'Purchase date', 'Party', 'Raw (kg)', 'Yield %',
                  'Avg price /kg (₹)', 'Sales value (₹)', 'Net profit (₹)',
                  'Profit / raw kg (₹)'], None, align_from=4)
s0 = SH + 1
NSUM = 60
for i in range(NSUM):
    r = s0 + i
    src = l0 + i
    for col, sc, fmt in ((1, 'A', 'General'), (2, 'C', DATE_F), (3, 'D', 'General'),
                         (4, 'F', N0), (5, 'O', PCT), (6, 'L', N2), (7, 'Q', INR),
                         (8, 'U', INR), (9, 'V', N2)):
        c = su.cell(row=r, column=col,
                    value=f'=IF(Lots!$A${src}="","",Lots!${sc}${src})')
        c.number_format = fmt
        c.font = Font(name=FONT, size=9, color=INK, bold=(col in (5, 8)))
        c.border = box
        if i % 2:
            c.fill = PatternFill('solid', fgColor=WASH)
    su.row_dimensions[r].height = 13
s1 = s0 + NSUM - 1
su.conditional_formatting.add(f'A{s0}:I{s1}',
    FormulaRule(formula=[f'$A{s0}=""'], border=Border(), stopIfTrue=True))
su.conditional_formatting.add(f'H{s0}:H{s1}',
    FormulaRule(formula=[f'AND($A{s0}<>"",$H{s0}<0)'],
                font=Font(name=FONT, size=9, bold=True, color=RED), stopIfTrue=False))
stot = s1 + 1
su.cell(row=stot, column=1, value='ALL LOTS')
for col in range(1, 10):
    c = su.cell(row=stot, column=col)
    c.fill = PatternFill('solid', fgColor=DEEP)
    c.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
    c.border = box
su.cell(row=stot, column=4, value=f'=SUM(D{s0}:D{s1})').number_format = N0
su.cell(row=stot, column=5,
        value=f'=IF(SUM(D{s0}:D{s1})=0,"",SUMPRODUCT(D{s0}:D{s1},'
              f'IF(E{s0}:E{s1}="",0,E{s0}:E{s1}))/SUM(D{s0}:D{s1}))').number_format = PCT
su.cell(row=stot, column=7, value=f'=SUM(G{s0}:G{s1})').number_format = INR
su.cell(row=stot, column=8, value=f'=SUM(H{s0}:H{s1})').number_format = INR
su.cell(row=stot, column=9,
        value=f'=IF(SUM(D{s0}:D{s1})=0,"",H{stot}/D{stot})').number_format = N2
for col in range(1, 10):
    su.cell(row=stot, column=col).font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
su.cell(row=stot + 2, column=1,
        value='Yield across all lots is weighted by raw kg, not a simple average.')\
  .font = Font(name=FONT, size=9, italic=True, color=MUTED)
su.print_area = f'A1:I{stot + 2}'      # keep the weighting note on the printed page
su.print_title_rows = f'{SH}:{SH}'
su.page_setup.orientation = 'landscape'
su.page_setup.fitToWidth = 1
su.page_setup.fitToHeight = 0
su.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Read me
rm = wb['Sheet']
rm.title = 'Read me'
wb.move_sheet('Read me', offset=-7)
title(rm, 'LOT ANALYSIS — HOW IT WORKS')
rm.column_dimensions['A'].width = 2
rm.column_dimensions['B'].width = 110
lines = [
    ('h', 'For each lot'),
    ('p', "1.  On 'Lots' add a row: lot no, dates, party, raw kg, raw amount, peeling ₹/kg,"),
    ('p', '    and the exchange rate. Those eight are all you type.'),
    ('p', "2.  On 'Lot Entry' add one row per grade produced: slabs, net kg per slab,"),
    ('p', '    gross weight per slab and the sale price per kg.'),
    ('p', "3.  On 'Lot Report' put the lot number in the yellow cell."),
    ('p', '4.  File ▸ Export ▸ Create PDF/XPS, "Selected sheet" — the directors\' copy.'),
    ('n', ''),
    ('h', 'The two weights are not the same'),
    ('p', 'Net kg per slab (usually 2) is what the buyer pays for:  Packed kg = slabs × net.'),
    ('p', 'Gross weight per slab (2.150, 2.200 …) is the final weight:  F.WT = slabs × gross.'),
    ('p', 'Yield is measured on F.WT, but the sale is on packed kg — so both are carried.'),
    ('n', ''),
    ('h', 'The arithmetic'),
    ('p', '    Amount (US$)   = Packed kg × Price per kg'),
    ('p', '    Yield %        = total F.WT ÷ raw kg purchased'),
    ('p', '    Total raw cost = raw amount + peeling ₹/kg × raw kg'),
    ('p', '    Expenses       = expense ₹/kg × packed kg'),
    ('p', '    Net profit     = sales value ₹ − total raw cost − expenses'),
    ('p', '    Profit per kg  = net profit ÷ RAW kg, not finished kg'),
    ('n', ''),
    ('h', 'Expenses'),
    ('p', "'Expenses' holds the standard rate per kg for each head, used by every lot."),
    ('p', "A lot that differs — a different freight, say — can override any single head in"),
    ('p', "the grey columns at the right of 'Lots'. Leave an override blank to use the"),
    ('p', 'standard rate. The report always shows the rate it actually used.'),
    ('n', ''),
    ('h', 'What can be typed into — everything else is locked'),
    ('b', "'Lots' — the eight blue columns and the grey expense overrides."),
    ('b', "'Lot Entry' — the whole grid.   'Grades' and 'Expenses' — the master lists."),
    ('y', "'Lot Report' — the lot number in the yellow cell F5."),
    ('k', 'Every calculated cell is protected, and formatting is locked.'),
    ('n', ''),
    ('h', 'A note on the printed statement'),
    ('p', 'Seeded from lot C 1374. Every figure on that sheet is reproduced here except the'),
    ('p', 'expense block: the rates printed there are rounded to two decimals, and the tool'),
    ('p', 'that produced it used slightly different underlying rates, so its expense total is'),
    ('p', '₹107,433.37 where these rates multiply out to ₹107,428.20 — a ₹5.17 difference'),
    ('p', 'that carries into the profit. This workbook is arithmetically consistent.'),
]
r = 4
for kind, text in lines:
    c = rm.cell(row=r, column=2, value=text)
    if kind == 'h':
        c.font = Font(name=FONT, size=10, bold=True, color=DEEP)
    elif kind == 'b':
        c.font = Font(name=FONT, size=10, color=BLUE_IN, bold=True)
    elif kind == 'y':
        # a chip in the 2-wide gutter, not a fill on a 110-wide column
        c.font = Font(name=FONT, size=10, color=BLUE_IN, bold=True)
        rm.cell(row=r, column=1).fill = PatternFill('solid', fgColor=YELLOW)
    else:
        c.font = Font(name=FONT, size=10, color=INK)
    r += 1
rm.print_area = f'A1:B{r}'
rm.page_setup.fitToWidth = 1
rm.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ---------------------------------------------------------------- protection
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.protection.password = PASSWORD
    ws.protection.sheet = True
    ws.protection.enable()
    for flag in ('formatCells', 'formatColumns', 'formatRows', 'insertRows', 'insertColumns',
                 'deleteRows', 'deleteColumns', 'insertHyperlinks', 'sort', 'pivotTables',
                 'objects', 'scenarios'):
        setattr(ws.protection, flag, True)
    ws.protection.autoFilter = False
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False
wb.security = WorkbookProtection(workbookPassword=PASSWORD, lockStructure=True)
wb.active = wb.index(rp)
wb.save(OUT)
print(f'wrote {OUT}: {len(LOTS)} lot(s), {len(LINES)} grade line(s), '
      f'{len(GRADES)} grades, {len(EXPENSES)} expense heads')
