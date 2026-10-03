#!/usr/bin/env python3
"""Builds PEI-Bills-Register.xlsx — the export bills on hand, kept day by day,
with a one-page Summary the partners can open or receive as a PDF.

    python tools/build-bills-xlsx.py out.xlsx                    # seeded from bills/index.html
    python tools/build-bills-xlsx.py out.xlsx --from live.xlsx   # keep the bills already entered
    python tools/build-bills-xlsx.py --relock out.xlsx           # after recalc

A bill's life on the register, all in US$ until the last step:

    Received  = Advance + Realised
    Balance   = Amount − Received
    Balance ₹ = Balance × the US$ → ₹ rate on Summary
    Status    = On hand | Advance received | Part realised      (open: on the Summary)
              | Realised   (balance nil, or closed as Settled)
              | Cancelled  (closed as Cancelled: kept for the record, counted nowhere)

A bill is never deleted to close it — it is realised, settled or cancelled, so
the register keeps its whole history. Clearing a row is for a line typed in
error.
"""
import sys, os, re, datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.drawing.image import Image as XLImage
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO = os.path.join(ROOT, 'stock-statement', 'assets', 'pei-mark.png')
SEED = os.path.join(ROOT, 'bills', 'index.html')

args = [a for i, a in enumerate(sys.argv[1:], start=1)
        if not a.startswith('--') and not sys.argv[i - 1] in ('--from', '--password', '--relock')]
OUT = args[0] if args else 'PEI-Bills-Register.xlsx'
SRC = sys.argv[sys.argv.index('--from') + 1] if '--from' in sys.argv else None
PASSWORD = (sys.argv[sys.argv.index('--password') + 1]
            if '--password' in sys.argv else 'PEI2026')
UNLOCK = Protection(locked=False)

B0 = 7                  # first bill row on Bills
BILL_ROWS = 1000        # bills the register holds, open and closed
B1 = B0 + BILL_ROWS - 1
L0, LIST_ROWS = 5, 100  # Lists: first row, rows per list
L1 = L0 + LIST_ROWS - 1
BUYER_SLOTS = 12        # buyers the Summary names; the rest fold into "Other buyers"
ADV_SLOTS = 4           # advances the Summary lists
LIST_SLOTS = 32         # open bills the Summary lists; more triggers a note
TYPED = 11              # Bills columns A..K are typed; L onwards are calculated

# ---------------------------------------------------------------- relock
# LibreOffice's recalculation drops the workbook-structure lock, the
# column-level unlocking that keeps typing open below the pre-ruled rows, and
# the full extent of data validation. Re-saving through openpyxl would strip
# the cached values again, so this patches the XML in place.
if '--relock' in sys.argv:
    import zipfile, shutil, tempfile
    from openpyxl.utils.protection import hash_password
    RELOCK = {'Bills': (TYPED, B0, B1), 'Lists': (2, L0, L1)}
    target = sys.argv[sys.argv.index('--relock') + 1]
    zin = zipfile.ZipFile(target)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    blob = {i.filename: d for i, d in items}
    el = f'<workbookProtection workbookPassword="{hash_password(PASSWORD)}" lockStructure="1"/>'
    wbx = blob['xl/workbook.xml'].decode('utf-8')
    if re.search(r'<workbookProtection[^>]*/>', wbx):
        wbx = re.sub(r'<workbookProtection[^>]*/>', el, wbx, count=1)
    elif '<workbookProtection' in wbx:
        wbx = re.sub(r'<workbookProtection.*?</workbookProtection>', el, wbx, count=1, flags=re.S)
    else:
        for a in ('<bookViews', '<sheets'):
            if a in wbx:
                wbx = wbx.replace(a, el + a, 1)
                break
    rid = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"',
                          blob['xl/_rels/workbook.xml.rels'].decode('utf-8')))
    targets = {}
    for m in re.finditer(r'<sheet\b[^>]*/>', wbx):
        a = dict(re.findall(r'(?:r:)?(\w+)="([^"]*)"', m.group(0)))
        if a.get('name') in RELOCK:
            targets[a['name']] = 'xl/' + rid.get(a.get('id', ''), '').lstrip('/')
    st_xml = blob['xl/styles.xml'].decode('utf-8')
    head, rest = st_xml.split('<cellXfs', 1)
    attrs, body_and_tail = rest.split('>', 1)
    body, tail = body_and_tail.split('</cellXfs>', 1)
    xfs = re.findall(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', body, re.S)
    added = []

    def unlocked_copy(xf):
        # LibreOffice writes an explicit locked="1", so an existing element is rewritten
        if '<protection' in xf:
            xf = re.sub(r'<protection\b[^>]*/>', '<protection locked="0" hidden="0"/>', xf, count=1)
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
        limit, top, last = RELOCK[name]
        sx = blob[f].decode('utf-8')

        def repoint(m, limit=limit):
            c = m.group(0)
            a = dict(re.findall(r'(\w+)="([^"]*)"', c))
            lo, hi = int(a.get('min', 0)), int(a.get('max', 0))
            if lo > limit:
                return c
            nid = len(xfs) + len(added)
            added.append(unlocked_copy(xfs[int(a.get('style', 0))]))
            open_ = (re.sub(r'style="\d+"', f'style="{nid}"', c) if 'style=' in c
                     else c[:-2] + f' style="{nid}"/>')
            if hi <= limit:
                return open_
            # a range straddling the last typed column is split: typed part open, rest as was
            return (re.sub(r'max="\d+"', f'max="{limit}"', open_) +
                    re.sub(r'min="\d+"', f'min="{limit + 1}"', c))

        sx = re.sub(r'<col\b[^>]*/>', repoint, sx)

        def widen(m, top=top, last=last):
            parts = []
            for rng in m.group(1).split():
                mm = re.fullmatch(rf'([A-Z]+){top}:([A-Z]+)\d+', rng)
                parts.append(f'{mm.group(1)}{top}:{mm.group(2)}{last}' if mm else rng)
            return 'sqref="' + ' '.join(parts) + '"'
        sx = re.sub(r'sqref="([^"]*)"', widen, sx)
        blob[f] = sx.encode('utf-8')
    for f in [k for k in blob if k.startswith('xl/worksheets/sheet') and k.endswith('.xml')]:
        sx = blob[f].decode('utf-8')
        if 'dataBar' in sx:
            sx = re.sub(r'(<x14:dataBar\b[^>]*?)gradient="(?:true|1)"', r'\1gradient="0"', sx)
            sx = re.sub(r'(<(?:x14:)?dataBar\b[^>]*?)minLength="\d+"', r'\1minLength="0"', sx)
            sx = re.sub(r'(<(?:x14:)?dataBar\b[^>]*?)maxLength="\d+"', r'\1maxLength="100"', sx)
            blob[f] = sx.encode('utf-8')
    if added:
        body += ''.join(added)
        attrs = re.sub(r'count="\d+"', f'count="{len(xfs) + len(added)}"', attrs)
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
    print(f'relocked {target}: structure locked, {len(added)} entry column ranges re-unlocked')
    sys.exit(0)

# ---------------------------------------------------------------- data
# Bills: [invoice, invoice date, item, buyer, amount, advance, advance date,
#         realised, realised date, close as, note]
def seed_from_page():
    html = open(SEED, encoding='utf-8').read()
    m = re.search(r'var DATA = \[(.*?)\n\s*\];', html, re.S)
    if not m:
        sys.exit(f'DATA list not found in {SEED}')
    q = r"'((?:[^'\\]|\\.)*)'"
    rows = re.findall(rf"\[\s*{q}\s*,\s*{q}\s*,\s*{q}\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*{q}\s*\]", m.group(1))
    out = []
    for inv, item, buyer, amt, adv, note in rows:
        adv = float(adv)
        out.append([inv, None, item, buyer or None, float(amt), adv or None, None,
                    None, None, None, note or None])
    return out

BILLS = seed_from_page()
ITEMS = ['Tuna', 'Mackerel', 'Shrimp', 'Squid', 'Cuttlefish']
BUYERS = sorted({b[3] for b in BILLS if b[3] and b[3] != 'part'})

if SRC:
    src = load_workbook(SRC)
    s = src['Bills']
    BILLS = []
    for r in range(B0, B1 + 1):
        row = [s.cell(row=r, column=c).value for c in range(1, TYPED + 1)]
        if any(v not in (None, '') for v in row):
            BILLS.append(row)
    if 'Lists' in src.sheetnames:
        ls = src['Lists']
        ITEMS = [ls.cell(row=r, column=1).value for r in range(L0, L1 + 1) if ls.cell(row=r, column=1).value]
        BUYERS = [ls.cell(row=r, column=2).value for r in range(L0, L1 + 1) if ls.cell(row=r, column=2).value]
    print(f'read {SRC}: {len(BILLS)} bills, {len(ITEMS)} items, {len(BUYERS)} buyers')
if len(BILLS) > BILL_ROWS:
    sys.exit(f'{len(BILLS)} bills will not fit in {BILL_ROWS} rows')

# ---------------------------------------------------------------- style
FONT = 'Arial'
DEEP, BAND, WASH, LINE_C = '065E7A', 'E7F3F9', 'F2F8FB', 'CBDFE8'
INK, MUTED, BRAND_C, RED, GREEN = '0B2C3A', '5E7C8B', '00A2D3', 'B3261E', '1B7A43'
BLUE_IN, YELLOW, GREY, PINK = '0000FF', 'FFFF00', 'D9D9D9', 'F9D7D5'
thin = Side(style='thin', color=LINE_C)
box = Border(left=thin, right=thin, top=thin, bottom=thin)

N0 = r'#,##0'
N2 = r'#,##0.00'
N2D = r'#,##0.00;-#,##0.00;"–"'          # zero shows as a dash
INR = r'[>=10000000]"₹" ##\,##\,##\,##0.00;[>=100000]"₹" ##\,##\,##0.00;"₹" #,##0.00'
INR0 = r'[>=10000000]"₹" ##\,##\,##\,##0;[>=100000]"₹" ##\,##\,##0;"₹" #,##0'
USD0 = r'"$" #,##0'
PCT = r'0.0%'
DATE_F = 'DD/MM/YYYY'
DATE_S = 'DD/MM/YY'

wb = Workbook()
wb.remove(wb.active)

def font(size=9, **k):
    return Font(name=FONT, size=size, color=k.pop('color', INK), **k)

def fill(c):
    return PatternFill('solid', fgColor=c)

def head_row(ws, row, cols, align_from=3, height=30, start=1):
    for i, t in enumerate(cols, start=start):
        c = ws.cell(row=row, column=i, value=t)
        c.font = font(8, bold=True, color='FFFFFF')
        c.fill = fill(DEEP)
        c.alignment = Alignment(horizontal='right' if i >= align_from else 'left',
                                vertical='bottom', wrap_text=True)
        c.border = box
    ws.row_dimensions[row].height = height

def title(ws, sub):
    ws['A1'] = 'PREMIER EXPORTS INTERNATIONAL'
    ws['A1'].font = font(14, bold=True, color=DEEP)
    ws['A2'] = sub
    ws['A2'].font = font(10, bold=True, color=MUTED)

R = lambda c: f'Bills!${c}${B0}:${c}${B1}'         # a whole Bills column, absolute
RNG = lambda c: f'Bills!${c}$1:${c}${B1}'           # same, indexed by sheet row

# ================================================================ Summary
su = wb.create_sheet('Summary')
for col, w in zip('ABCDEFGH', (16, 9, 13, 16, 12, 14, 13, 15)):
    su.column_dimensions[col].width = w
for col in 'JKL':                                   # helpers, outside the print area
    su.column_dimensions[col].width = 6
    su.column_dimensions[col].hidden = True

su.row_dimensions[1].height = 34
if os.path.exists(LOGO):
    im = XLImage(LOGO)
    im.height, im.width = 42, 52
    su.add_image(im, 'A1')
su['B1'] = 'Premier Exports International'
su['B1'].font = font(16, color=BRAND_C)
su['B1'].alignment = Alignment(vertical='center')
for i, t in enumerate(['AP X/453, NH-66 Highway, Chandiroor P.O.,',
                       'Aroor, Alappuzha, Kerala - 688 537, India',
                       'GSTIN 32AADFP3158P1ZZ']):
    su.merge_cells(start_row=1 + i, start_column=5, end_row=1 + i, end_column=8)
    c = su.cell(row=1 + i, column=5, value=t)
    c.font = font(8, color=MUTED)
    c.alignment = Alignment(horizontal='right', vertical='bottom' if i == 0 else 'top')
su.row_dimensions[4].height = 5
for col in range(1, 9):
    su.cell(row=4, column=col).border = Border(bottom=Side(style='medium', color=BRAND_C))

su.row_dimensions[5].height = 22
su['A5'] = 'BILLS ON HAND'
su['A5'].font = font(16, bold=True)
su['A5'].alignment = Alignment(vertical='center')
su['A6'] = 'EXPORT BILLS RECEIVABLE · BILLED, RECEIVED AND BALANCE DUE'
su['A6'].font = font(8, bold=True, color=MUTED)
for r, lab, val, fmt, dv in ((5, 'AS AT', '=TODAY()', DATE_F,
                              DataValidation(type='date', operator='between',
                                             formula1='DATE(2020,1,1)', formula2='DATE(2040,12,31)')),
                             (6, 'US$ → ₹', 90, N2,
                              DataValidation(type='decimal', operator='between',
                                             formula1='1', formula2='1000'))):
    k = su.cell(row=r, column=7, value=lab)
    k.font = font(8, bold=True, color=MUTED)
    k.alignment = Alignment(horizontal='right', vertical='center')
    v = su.cell(row=r, column=8, value=val)
    v.number_format = fmt
    v.font = font(11, bold=True, color=BLUE_IN)
    v.fill = fill(YELLOW)
    v.border = box
    v.alignment = Alignment(horizontal='right', vertical='center')
    v.protection = UNLOCK
    dv.error = 'Enter a date' if r == 5 else 'Enter the rupee value of one US dollar'
    dv.showErrorMessage = True
    su.add_data_validation(dv)
    dv.add(f'H{r}')
su.row_dimensions[6].height = 18
wb.defined_names.add(DefinedName('AsAt', attr_text='Summary!$H$5'))
wb.defined_names.add(DefinedName('FX', attr_text='Summary!$H$6'))

def inr_short(expr):
    """₹ in lakh or crore, as text — the tiles' second line."""
    return (f'=IF(FX="","Exchange rate not set","₹ "&IF(ABS({expr})*FX>=1E7,'
            f'TEXT({expr}*FX/1E7,"0.00")&" Cr",IF(ABS({expr})*FX>=1E5,'
            f'TEXT({expr}*FX/1E5,"0.00")&" L",TEXT({expr}*FX,"#,##0"))))')

OPEN_N = f'COUNT({R("R")})'
REG_N = f'SUMPRODUCT(--(LEN({R("O")})>0))'
BILLED = f'SUMIFS({R("E")},{R("V")},1)'
RECEIVED = f'SUMIFS({R("L")},{R("V")},1)'
BALANCE = f'SUM({R("T")})'

TK = 8
su.row_dimensions[7].height = 6
tiles = [('Bills on hand', f'={OPEN_N}', N0, f'="of "&{REG_N}&" on the register"'),
         ('Billed', f'={BILLED}', USD0, inr_short(BILLED)),
         ('Received against them', f'={RECEIVED}', USD0, inr_short(RECEIVED)),
         ('Balance receivable', f'={BALANCE}', USD0, inr_short(BALANCE))]
for i, (lab, f, fmt, sub) in enumerate(tiles):
    c0, c1 = 1 + 2 * i, 2 + 2 * i
    accent = i == 3
    for r in (TK, TK + 1, TK + 2):
        su.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
        for cc in (c0, c1):
            su.cell(row=r, column=cc).fill = fill(DEEP if accent else WASH)
    k = su.cell(row=TK, column=c0, value=lab.upper())
    k.font = font(7, bold=True, color='CFE8F2' if accent else MUTED)
    k.alignment = Alignment(horizontal='left', indent=1, vertical='bottom')
    v = su.cell(row=TK + 1, column=c0, value=f)
    v.number_format = fmt
    v.font = font(15 if accent else 13, bold=True, color='FFFFFF' if accent else INK)
    v.alignment = Alignment(horizontal='left', indent=1, vertical='center')
    u = su.cell(row=TK + 2, column=c0, value=sub)
    u.font = font(8, bold=i > 0, color='FFFFFF' if accent else (INK if i else MUTED))
    u.alignment = Alignment(horizontal='left', indent=1, vertical='top')
    edge = Side(style='thin', color=DEEP if accent else LINE_C)
    gap = Side(style='thick', color='FFFFFF')       # white rule = gutter between tiles
    for r in (TK, TK + 1, TK + 2):
        for cc in (c0, c1):
            su.cell(row=r, column=cc).border = Border(
                top=edge if r == TK else None, bottom=edge if r == TK + 2 else None,
                left=(gap if cc == c0 and i else edge) if cc == c0 else None,
                right=(gap if cc == c1 and i < 3 else edge) if cc == c1 else None)
su.row_dimensions[TK].height = 15
su.row_dimensions[TK + 1].height = 22
su.row_dimensions[TK + 2].height = 15

REGL = TK + 3
su.merge_cells(start_row=REGL, start_column=1, end_row=REGL, end_column=8)
su.cell(row=REGL, column=1, value=(
    f'="Register: "&{REG_N}&" bills · "&COUNTIF({R("O")},"Realised")&" realised (US$ "'
    f'&TEXT(SUMIFS({R("E")},{R("O")},"Realised"),"#,##0")&") · "'
    f'&COUNTIF({R("O")},"Cancelled")&" cancelled"')).font = font(8, color=MUTED)
su.row_dimensions[REGL].height = 16

def section(ws, row, c0, c1, text):
    ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
    c = ws.cell(row=row, column=c0, value=text.upper())
    c.font = font(8, bold=True, color=DEEP)
    c.alignment = Alignment(vertical='center', indent=0)
    for col in range(c0, c1 + 1):
        ws.cell(row=row, column=col).fill = fill(BAND)
    ws.row_dimensions[row].height = 16

def subhead(ws, row, c0, labels, right_from):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=c0 + i, value=t)
        c.font = font(7, bold=True, color='FFFFFF')
        c.fill = fill(DEEP)
        c.alignment = Alignment(horizontal='right' if c0 + i >= right_from else 'left',
                                vertical='bottom', wrap_text=True)
    ws.row_dimensions[row].height = 22

def zebra(ws, rng, key_col, first):
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(${key_col}{first}<>"",MOD(ROW(),2)=0)'], fill=fill(WASH)))

# ---- balance by buyer (A–E) and ageing (F–H)
S1 = REGL + 2
su.row_dimensions[S1 - 1].height = 6
section(su, S1, 1, 5, 'Balance by buyer — US$')
section(su, S1, 6, 8, 'Ageing of balance')
subhead(su, S1 + 1, 1, ['Buyer', 'Bills', 'Balance (US$)', 'Balance (₹)', 'Share'], 2)
subhead(su, S1 + 1, 6, ['Days outstanding', 'Bills', 'Balance (US$)'], 7)
bs0 = S1 + 2
for k in range(BUYER_SLOTS):
    r = bs0 + k
    su[f'J{r}'] = f'=IFERROR(MATCH(LARGE({R("U")},{k + 1}),{R("U")},0),"")'
    su[f'A{r}'] = f'=IF($J{r}="","",INDEX({R("S")},$J{r}))'
    su[f'B{r}'] = f'=IF($J{r}="","",COUNTIFS({R("S")},$A{r},{R("V")},1))'
    su[f'C{r}'] = f'=IF($J{r}="","",SUMIFS({R("T")},{R("S")},$A{r}))'
    su[f'D{r}'] = f'=IF(OR($J{r}="",FX=""),"",$C{r}*FX)'
    su[f'E{r}'] = f'=IF(OR($J{r}="",{BALANCE}=0),"",$C{r}/{BALANCE})'
    for col, fmt in zip('ABCDE', ('General', N0, N2, INR0, PCT)):
        su[f'{col}{r}'].number_format = fmt
        su[f'{col}{r}'].font = font(9)
    su.row_dimensions[r].height = 14
bs1 = bs0 + BUYER_SLOTS - 1
zebra(su, f'A{bs0}:E{bs1}', 'J', bs0)
su.conditional_formatting.add(f'E{bs0}:E{bs1}', DataBarRule(
    start_type='num', start_value=0, end_type='max', color='A9D3E3', showValue=True))
bo = bs1 + 1                                     # buyers beyond the slots
others = f'MAX(0,COUNT({R("U")})-{BUYER_SLOTS})'
su[f'A{bo}'] = f'=IF({others}>0,"Other buyers ("&{others}&")","")'
su[f'B{bo}'] = f'=IF({others}>0,{OPEN_N}-SUM(B{bs0}:B{bs1}),"")'
su[f'C{bo}'] = f'=IF({others}>0,{BALANCE}-SUM(C{bs0}:C{bs1}),"")'
su[f'D{bo}'] = f'=IF(OR(C{bo}="",FX=""),"",C{bo}*FX)'
su[f'E{bo}'] = f'=IF(OR(C{bo}="",{BALANCE}=0),"",C{bo}/{BALANCE})'
for col, fmt in zip('ABCDE', ('General', N0, N2, INR0, PCT)):
    su[f'{col}{bo}'].number_format = fmt
    su[f'{col}{bo}'].font = font(9, italic=True, color=MUTED)
bt = bo + 1
su[f'A{bt}'] = 'TOTAL'
su[f'B{bt}'] = f'={OPEN_N}'
su[f'C{bt}'] = f'={BALANCE}'
su[f'D{bt}'] = f'=IF(FX="","",C{bt}*FX)'
su[f'E{bt}'] = f'=IF(C{bt}=0,"",1)'
for col, fmt in zip('ABCDE', ('General', N0, N2, INR0, PCT)):
    c = su[f'{col}{bt}']
    c.number_format = fmt
    c.font = font(9, bold=True)
    c.border = Border(top=Side(style='medium', color=DEEP), bottom=thin)
    c.alignment = Alignment(horizontal='left' if col == 'A' else 'right')

AGES = ['0–30 days', '31–60 days', '61–90 days', 'Over 90 days', 'No invoice date']
for i, band in enumerate(AGES):
    r = bs0 + i
    su[f'F{r}'] = band
    su[f'G{r}'] = f'=COUNTIFS({R("Q")},$F{r})'
    su[f'H{r}'] = f'=SUMIFS({R("T")},{R("Q")},$F{r})'
    su[f'F{r}'].font = font(9, italic=(band == 'No invoice date'),
                           color=MUTED if band == 'No invoice date' else INK)
    for col, fmt in (('G', N0), ('H', N2)):
        su[f'{col}{r}'].number_format = fmt
        su[f'{col}{r}'].font = font(9)
    if i % 2:
        for col in 'FGH':
            su[f'{col}{r}'].fill = fill(WASH)
at = bs0 + len(AGES)
su[f'F{at}'] = 'TOTAL'
su[f'G{at}'] = f'=SUM(G{bs0}:G{at - 1})'
su[f'H{at}'] = f'=SUM(H{bs0}:H{at - 1})'
for col, fmt in (('F', 'General'), ('G', N0), ('H', N2)):
    c = su[f'{col}{at}']
    c.number_format = fmt
    c.font = font(9, bold=True)
    c.border = Border(top=Side(style='medium', color=DEEP), bottom=thin)
su.conditional_formatting.add(f'H{bs0 + 3}', FormulaRule(
    formula=[f'H{bs0 + 3}>0'], font=Font(name=FONT, size=9, bold=True, color=RED)))

# ---- advances received (F–H), under the ageing
A1 = at + 2
section(su, A1, 6, 8, 'Advances received')
subhead(su, A1 + 1, 6, ['Invoice', 'Buyer', 'Advance (US$)'], 8)
as0 = A1 + 2
for k in range(ADV_SLOTS):
    r = as0 + k
    su[f'K{r}'] = f'=IFERROR(SMALL({R("W")},{k + 1}),"")'
    su[f'F{r}'] = f'=IF($K{r}="","",INDEX({RNG("A")},$K{r})&"")'
    su[f'G{r}'] = f'=IF($K{r}="","",INDEX({RNG("S")},$K{r}))'
    su[f'H{r}'] = f'=IF($K{r}="","",INDEX({RNG("F")},$K{r}))'
    su[f'H{r}'].number_format = N2
    for col in 'FGH':
        su[f'{col}{r}'].font = font(9)
zebra(su, f'F{as0}:H{as0 + ADV_SLOTS - 1}', 'K', as0)
adt = as0 + ADV_SLOTS
ADV_N = f'COUNT({R("W")})'
su[f'F{adt}'] = f'=IF({ADV_N}=0,"None on open bills",{ADV_N}&IF({ADV_N}=1," bill"," bills")' \
                f'&IF({ADV_N}>{ADV_SLOTS},", "&({ADV_N}-{ADV_SLOTS})&" not listed",""))'
su[f'H{adt}'] = f'=SUMIFS({R("F")},{R("V")},1)'
su[f'H{adt}'].number_format = N2
su.merge_cells(f'F{adt}:G{adt}')
for col in 'FGH':
    c = su[f'{col}{adt}']
    c.font = font(9, bold=True)
    c.border = Border(top=Side(style='medium', color=DEEP), bottom=thin)
if adt != bt:
    raise SystemExit(f'layout: advances total row {adt} should line up with buyer total {bt}')
gutter = Side(style='thick', color='FFFFFF')
for r in range(S1, bt + 1):
    c = su.cell(row=r, column=6)
    b = c.border
    c.border = Border(left=gutter, right=b.right, top=b.top, bottom=b.bottom)
    if r not in (S1, A1):
        c.alignment = Alignment(horizontal='left', indent=1, vertical=c.alignment.vertical,
                                wrap_text=c.alignment.wrap_text)
for r in (A1, A1 + 1):              # these rows are buyer rows too — keep them buyer height
    su.row_dimensions[r].height = 14

# ---- warnings — say what is wrong and where, or nothing at all
WR = bt + 1
su.merge_cells(start_row=WR, start_column=1, end_row=WR, end_column=8)
su.cell(row=WR, column=1, value=(
    '=TRIM(IF(SUM(' + R('Z') + ')>0,"Check Bills: "&SUM(' + R('Z') + ')&" rows share an invoice no. ","")'
    '&IF(SUM(' + R('Y') + ')>0,"Check Bills: "&SUM(' + R('Y') + ')&" bills lack an invoice no or amount. ","")'
    '&IF(SUM(' + R('X') + ')>0,"Check Bills: "&SUM(' + R('X') + ')&" bills show more received than billed. ","")'
    '&IF(FX="","Exchange rate not set — rupee figures are blank.",""))')
).font = font(8, bold=True, color=RED)
su.row_dimensions[WR].height = 15

# ---- the open bills, in register order
LS = WR + 2
su.row_dimensions[LS - 1].height = 6
section(su, LS, 1, 8, 'Bills on hand')
subhead(su, LS + 1, 1, ['Invoice no', 'Inv. date', 'Item', 'Buyer', 'Amount (US$)',
                        'Received (US$)', 'Balance (US$)', 'Balance (₹)'], 5)
LT = LS + 2
su.merge_cells(start_row=LT, start_column=1, end_row=LT, end_column=4)
su[f'A{LT}'] = f'="TOTAL — "&{OPEN_N}&IF({OPEN_N}=1," BILL"," BILLS")&" ON HAND"'
su[f'E{LT}'] = f'={BILLED}'
su[f'F{LT}'] = f'={RECEIVED}'
su[f'G{LT}'] = f'={BALANCE}'
su[f'H{LT}'] = f'=IF(FX="","",G{LT}*FX)'
for col, fmt in zip('ABCDEFGH', ('General',) * 4 + (N2, N2D, N2, INR0)):
    c = su[f'{col}{LT}']
    c.number_format = fmt
    c.fill = fill(BAND)
    c.font = font(9, bold=True, color=DEEP)
    c.border = Border(bottom=Side(style='medium', color=DEEP))
    c.alignment = Alignment(horizontal='left' if col == 'A' else 'right', vertical='center')
su.row_dimensions[LT].height = 18
ls0 = LT + 1

def idx(col, r, guard=True):
    inner = f'INDEX({RNG(col)},$L{r})'
    if guard:       # INDEX on an empty cell yields 0, not blank
        return f'=IF($L{r}="","",IF({inner}="","",{inner}))'
    return f'=IF($L{r}="","",{inner})'

for k in range(LIST_SLOTS):
    r = ls0 + k
    su[f'L{r}'] = f'=IFERROR(SMALL({R("R")},{k + 1}),"")'
    for col, sc, fmt, g in (('A', 'A', 'General', True), ('B', 'B', DATE_S, True),
                            ('C', 'C', 'General', True), ('D', 'D', 'General', True),
                            ('E', 'E', N2, True), ('F', 'L', N2D, False),
                            ('G', 'M', N2, False), ('H', 'N', INR0, False)):
        c = su[f'{col}{r}']
        c.value = idx(sc, r, g)
        c.number_format = fmt
        c.font = font(9)
        if col == 'B':
            c.alignment = Alignment(horizontal='left')
    su.row_dimensions[r].height = 13
ls1 = ls0 + LIST_SLOTS - 1
zebra(su, f'A{ls0}:H{ls1}', 'L', ls0)
su.conditional_formatting.add(f'A{ls0}:H{ls1}', FormulaRule(
    formula=[f'AND($L{ls0}<>"",INDEX({RNG("P")},$L{ls0})<>"",INDEX({RNG("P")},$L{ls0})>90)'],
    font=Font(name=FONT, size=9, color=RED)))
ov = ls1 + 1
su.merge_cells(start_row=ov, start_column=1, end_row=ov, end_column=8)
su[f'A{ov}'] = (f'=IF({OPEN_N}>{LIST_SLOTS},"Note: "&({OPEN_N}-{LIST_SLOTS})&" more bills on hand are '
                f'not listed here — they are in the totals and on the Bills sheet.","")')
su[f'A{ov}'].font = font(8, bold=True, color=RED)
fn = ov + 1
su.merge_cells(start_row=fn, start_column=1, end_row=fn, end_column=8)
su[f'A{fn}'] = (f'="Balance = amount − advance − realised. A bill leaves this page once realised or cancelled; '
                f'every bill stays on the Bills sheet. Red: over 90 days. Rupees at ₹ "&IF(FX="","—",TEXT(FX,"0.00"))'
                f'&" per US$, as at "&TEXT(AsAt,"dd/mm/yyyy")&"."')
su[f'A{fn}'].font = font(7, italic=True, color=MUTED)
su[f'A{fn}'].alignment = Alignment(wrap_text=True, vertical='top')
su.row_dimensions[fn].height = 22
for col in range(1, 9):
    su.cell(row=fn, column=col).border = Border(top=thin)

su.print_area = f'A1:H{fn}'
su.page_setup.orientation = 'portrait'
su.page_setup.paperSize = su.PAPERSIZE_A4
su.page_setup.fitToWidth = 1
su.page_setup.fitToHeight = 1
su.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
su.page_margins.left = su.page_margins.right = 0.4
su.page_margins.top = su.page_margins.bottom = 0.4
su.print_options.horizontalCentered = True

# ================================================================ Bills
bi = wb.create_sheet('Bills')
title(bi, 'BILLS REGISTER — ONE ROW PER EXPORT BILL, OPEN OR CLOSED')
bi['A3'] = (f'="As at "&TEXT(AsAt,"dd/mm/yyyy")&"   ·   US$ → ₹ "&IF(FX="","not set",TEXT(FX,"0.00"))'
            f'&"   ·   On hand: "&{OPEN_N}&" bills, US$ "&TEXT({BALANCE},"#,##0.00")&"   (set the date and rate on Summary)"')
bi['A3'].font = font(9, bold=True, color=DEEP)
bi['A4'] = ('Type in the white columns (blue text). Received, balance, status and age are worked out. '
            'Close a bill by realising it, or with Close as: Settled / Cancelled — do not clear it.')
bi['A4'].font = font(9, italic=True, color=MUTED)
bi.merge_cells(start_row=5, start_column=1, end_row=5, end_column=TYPED)
bi['A5'] = 'TYPE HERE'
bi.merge_cells(start_row=5, start_column=TYPED + 1, end_row=5, end_column=17)
bi.cell(row=5, column=TYPED + 1, value='WORKED OUT — LOCKED')
for col in range(1, 18):
    c = bi.cell(row=5, column=col)
    c.fill = fill(BAND if col <= TYPED else GREY)
    c.font = font(8, bold=True, color=DEEP if col <= TYPED else MUTED)
    c.alignment = Alignment(horizontal='center')
HEAD = ['Invoice no', 'Invoice date', 'Item', 'Buyer', 'Amount (US$)', 'Advance received (US$)',
        'Advance date', 'Realised (US$)', 'Realised date', 'Close as', 'Note',
        'Received (US$)', 'Balance (US$)', 'Balance (₹)', 'Status', 'Days out', 'Age']
HELP = ['List key', 'Buyer key', 'Open balance', 'Buyer rank', 'Open', 'Advance key',
        'Over-received', 'Missing data', 'Duplicate inv.']
head_row(bi, 6, HEAD + HELP, align_from=5, height=30)
TEXT_COLS = (3, 4, 10, 11, 15, 17)          # item, buyer, close as, note, status, age
for i in range(1, 18):
    left = i == 1 or i in TEXT_COLS
    bi.cell(row=6, column=i).alignment = Alignment(horizontal='left' if left else 'right',
                                                   indent=1 if i in TEXT_COLS else 0,
                                                   vertical='bottom', wrap_text=True)
for i, w in enumerate((17, 11, 12, 18, 13, 13, 11, 13, 11, 11, 20,
                       13, 14, 17, 16, 7, 14), start=1):
    bi.column_dimensions[get_column_letter(i)].width = w
for i in range(18, 18 + len(HELP)):
    bi.column_dimensions[get_column_letter(i)].width = 9
    bi.column_dimensions[get_column_letter(i)].hidden = True
for i in range(1, TYPED + 1):     # typing stays open below the pre-ruled rows too
    bi.column_dimensions[get_column_letter(i)].protection = UNLOCK

def used(r):
    return f'OR($A{r}<>"",$E{r}<>"")'

for i in range(BILL_ROWS):
    r = B0 + i
    vals = BILLS[i] if i < len(BILLS) else [None] * TYPED
    for col in range(1, TYPED + 1):
        c = bi.cell(row=r, column=col, value=vals[col - 1])
        c.font = font(10, color=BLUE_IN)
        c.protection = UNLOCK
        c.border = box
        if col in (2, 7, 9):
            c.number_format = DATE_F
        elif col in (5, 6, 8):
            c.number_format = N2
        elif col in TEXT_COLS:
            c.alignment = Alignment(horizontal='left', indent=1)
    u = used(r)
    calc = {
        'L': f'=IF({u},N($F{r})+N($H{r}),"")',
        'O': (f'=IF(NOT({u}),"",IF($J{r}="Cancelled","Cancelled",IF($E{r}="","Amount missing",'
              f'IF(OR($J{r}="Settled",$E{r}-N($F{r})-N($H{r})<=0.005),"Realised",'
              f'IF(N($H{r})>0,"Part realised",IF(N($F{r})>0,"Advance received","On hand"))))))'),
        'M': (f'=IF(OR($O{r}="",$O{r}="Cancelled",$O{r}="Amount missing"),"",'
              f'IF($O{r}="Realised",0,$E{r}-N($F{r})-N($H{r})))'),
        'N': f'=IF(OR($M{r}="",FX=""),"",$M{r}*FX)',
        'V': (f'=IF(OR($O{r}="On hand",$O{r}="Advance received",$O{r}="Part realised",'
              f'$O{r}="Amount missing"),1,0)'),
        'P': f'=IF(AND($V{r}=1,ISNUMBER($B{r})),AsAt-$B{r},"")',
        'Q': (f'=IF($V{r}<>1,"",IF($P{r}="","No invoice date",IF($P{r}<=30,"0–30 days",'
              f'IF($P{r}<=60,"31–60 days",IF($P{r}<=90,"61–90 days","Over 90 days")))))'),
        'R': f'=IF($V{r}=1,ROW(),"")',
        'S': f'=IF({u},IF(TRIM($D{r})="","Buyer not noted",TRIM($D{r})),"")',
        'T': f'=IF(AND($V{r}=1,ISNUMBER($M{r})),$M{r},0)',
        'U': (f'=IF($V{r}=1,IF(COUNTIFS($S${B0}:$S{r},$S{r},$V${B0}:$V{r},1)=1,'
              f'SUMIFS($T${B0}:$T${B1},$S${B0}:$S${B1},$S{r})-ROW()/1E9,""),"")'),
        'W': f'=IF(AND($V{r}=1,N($F{r})>0),ROW(),"")',
        'X': f'=IF(AND({u},$O{r}<>"Cancelled",ISNUMBER($E{r}),N($F{r})+N($H{r})>$E{r}+0.005),1,0)',
        'Y': f'=IF(AND({u},OR($A{r}="",$E{r}="")),1,0)',
        'Z': f'=IF(AND($A{r}<>"",COUNTIF($A${B0}:$A${B1},$A{r})>1),1,0)',
    }
    for col, f in calc.items():
        c = bi[f'{col}{r}']
        c.value = f
        c.font = font(10, bold=(col == 'M'))
        if col <= 'Q':
            c.fill = fill(WASH)
            c.border = box
        else:
            c.font = font(8, color=MUTED)
    bi[f'L{r}'].number_format = N2D
    bi[f'M{r}'].number_format = N2
    bi[f'N{r}'].number_format = INR
    bi[f'P{r}'].number_format = N0
    bi[f'P{r}'].alignment = Alignment(horizontal='right')
    bi[f'O{r}'].alignment = Alignment(horizontal='left', indent=1)
    bi[f'Q{r}'].alignment = Alignment(horizontal='left', indent=1)

# dropdowns and type checks
def dv_add(ws, dv, ref):
    ws.add_data_validation(dv)
    dv.add(ref)
dv_item = DataValidation(type='list', formula1='=ItemList', allow_blank=True, errorStyle='warning',
                         error='Not on the Lists sheet. Keep it anyway?', showErrorMessage=True)
dv_buyer = DataValidation(type='list', formula1='=BuyerList', allow_blank=True, errorStyle='warning',
                          error='Not on the Lists sheet. Keep it anyway? (Add new buyers there.)',
                          showErrorMessage=True)
dv_close = DataValidation(type='list', formula1='"Settled,Cancelled"', allow_blank=True,
                          error='Choose Settled or Cancelled, or leave blank', showErrorMessage=True)
dv_money = DataValidation(type='decimal', operator='greaterThanOrEqual', formula1='0', allow_blank=True,
                          error='Enter an amount in US dollars, 0 or more', showErrorMessage=True)
dv_date = DataValidation(type='date', operator='between', formula1='DATE(2020,1,1)',
                         formula2='DATE(2040,12,31)', allow_blank=True,
                         error='Enter a date, e.g. 30/09/2026', showErrorMessage=True)
dv_add(bi, dv_item, f'C{B0}:C{B1}')
dv_add(bi, dv_buyer, f'D{B0}:D{B1}')
dv_add(bi, dv_close, f'J{B0}:J{B1}')
dv_add(bi, dv_money, f'E{B0}:E{B1}')
dv_money.add(f'F{B0}:F{B1}')
dv_money.add(f'H{B0}:H{B1}')
dv_add(bi, dv_date, f'B{B0}:B{B1}')
dv_date.add(f'G{B0}:G{B1}')
dv_date.add(f'I{B0}:I{B1}')

# colour carries the state; the Status column always says it in words
rows = f'A{B0}:Q{B1}'
bi.conditional_formatting.add(f'A{B0}:A{B1}', FormulaRule(
    formula=[f'$Z{B0}=1'], fill=fill(PINK), font=Font(name=FONT, size=10, bold=True, color=RED),
    stopIfTrue=True))
bi.conditional_formatting.add(f'A{B0}:A{B1}', FormulaRule(
    formula=[f'AND($A{B0}="",$E{B0}<>"")'], fill=fill(PINK)))
bi.conditional_formatting.add(f'E{B0}:E{B1}', FormulaRule(
    formula=[f'AND($A{B0}<>"",$E{B0}="")'], fill=fill(PINK)))
bi.conditional_formatting.add(f'F{B0}:H{B1}', FormulaRule(
    formula=[f'$X{B0}=1'], fill=fill(PINK)))
bi.conditional_formatting.add(rows, FormulaRule(
    formula=[f'$O{B0}="Cancelled"'], font=Font(name=FONT, size=10, color='9AA8AF', strike=True)))
bi.conditional_formatting.add(rows, FormulaRule(
    formula=[f'$O{B0}="Realised"'], font=Font(name=FONT, size=10, color='7B8E97')))
bi.conditional_formatting.add(f'O{B0}:O{B1}', FormulaRule(
    formula=[f'$O{B0}="Realised"'], font=Font(name=FONT, size=10, bold=True, color=GREEN)))
bi.conditional_formatting.add(f'P{B0}:Q{B1}', FormulaRule(
    formula=[f'AND(ISNUMBER($P{B0}),$P{B0}>90)'], font=Font(name=FONT, size=10, bold=True, color=RED)))

bi.freeze_panes = f'B{B0}'
bi.auto_filter.ref = f'A6:Q{B1}'
bi.print_title_rows = '6:6'
bi.page_setup.orientation = 'landscape'
bi.page_setup.paperSize = bi.PAPERSIZE_A4
bi.page_setup.fitToWidth = 1
bi.page_setup.fitToHeight = 0
bi.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Lists
li = wb.create_sheet('Lists')
title(li, 'LISTS — THE DROPDOWNS ON THE BILLS SHEET')
li['A3'] = 'Add a new item or buyer in the next empty cell; the dropdowns pick it up at once.'
li['A3'].font = font(9, italic=True, color=MUTED)
head_row(li, 4, ['Item', 'Buyer'], align_from=9, height=20)
li.column_dimensions['A'].width = 24
li.column_dimensions['B'].width = 28
for i in range(LIST_ROWS):
    r = L0 + i
    for col, vals in ((1, ITEMS), (2, BUYERS)):
        c = li.cell(row=r, column=col, value=vals[i] if i < len(vals) else None)
        c.font = font(10, color=BLUE_IN)
        c.border = box
        c.protection = UNLOCK
for col in 'AB':
    li.column_dimensions[col].protection = UNLOCK
wb.defined_names.add(DefinedName('ItemList', attr_text=f'OFFSET(Lists!$A${L0},0,0,MAX(1,COUNTA(Lists!$A${L0}:$A${L1})),1)'))
wb.defined_names.add(DefinedName('BuyerList', attr_text=f'OFFSET(Lists!$B${L0},0,0,MAX(1,COUNTA(Lists!$B${L0}:$B${L1})),1)'))

# ================================================================ Read me
rm = wb.create_sheet('Read me', 0)
title(rm, 'BILLS REGISTER — HOW IT WORKS')
rm.column_dimensions['A'].width = 2
rm.column_dimensions['B'].width = 112
lines = [
    ('h', 'Every day, on the Bills sheet'),
    ('p', '•  New bill raised — next empty row: invoice no, invoice date, item, buyer, amount (US$).'),
    ('p', '•  Advance received — on that bill\'s row: Advance received (US$) and Advance date.'),
    ('p', '•  Payment realised — Realised (US$) and Realised date. When advance + realised reaches the amount,'),
    ('p', '    the bill turns Realised and leaves the Summary. A part payment leaves it open as Part realised.'),
    ('p', '•  Bank paid short and nothing more will come — Close as: Settled. The bill closes at the amount received.'),
    ('p', '•  Bill raised in error or withdrawn — Close as: Cancelled. It stays on the sheet, struck through,'),
    ('p', '    and counts nowhere. Partners can still see that it existed.'),
    ('p', '•  A wrong figure — click the cell and type over it. A line typed by mistake (never a real bill) —'),
    ('p', '    select its cells from Invoice no to Note and press Delete. The empty row is ignored.'),
    ('n', ''),
    ('h', 'For the partners — the Summary sheet'),
    ('p', '"As at" shows today\'s date by itself. Type a date to report as at another day; type =TODAY() to go back.'),
    ('p', 'Set US$ → ₹ to the day\'s rate: every rupee figure in the workbook uses that one cell.'),
    ('p', 'Then File ▸ Export ▸ Create PDF/XPS, "Selected sheet" — one A4 page — or share the workbook itself.'),
    ('p', 'It shows bills on hand, the amount billed, what has come in against them and the balance, in US$ and ₹;'),
    ('p', 'the balance by buyer; its age from the invoice date; the advances received; and every open bill.'),
    ('n', ''),
    ('h', 'Status — worked out, never typed'),
    ('p', 'On hand — nothing received yet.'),
    ('p', 'Advance received — an advance, no realisation yet.'),
    ('p', 'Part realised — some of the bill realised, balance still due.'),
    ('p', 'Amount missing — an invoice no with no amount yet; still listed, so it is not forgotten.'),
    ('p', 'These four are the bills on hand, and the only ones the Summary counts.'),
    ('p', 'Realised — advance + realised covers the amount, or closed as Settled. Shown grey.'),
    ('p', 'Cancelled — closed as Cancelled. Shown grey and struck through.'),
    ('n', ''),
    ('h', 'Checks that flag themselves'),
    ('p', 'Pink on Bills, and a red line on Summary, when: two rows share an invoice no; a row has an amount but'),
    ('p', 'no invoice no, or the reverse; or more has been received than was billed. A bill over 90 days old shows red.'),
    ('p', 'Ageing needs the invoice date — bills without one are counted under "No invoice date".'),
    ('n', ''),
    ('h', 'What can be typed into — everything else is locked'),
    ('b', "'Bills' — the white columns, Invoice no to Note.   'Lists' — items and buyers for the dropdowns."),
    ('y', "'Summary' — the two yellow cells: As at, and US$ → ₹."),
    ('k', 'Formulas, headings and formatting are protected, and rows cannot be deleted, so the register'),
    ('k', 'cannot be broken by accident. Filtering the Bills sheet still works — filter Status to see one kind.'),
    ('k', 'To print the register itself, filter Status and untick (Blanks) first, so the empty rows stay off the paper.'),
    ('n', ''),
    ('h', 'As loaded'),
    ('p', 'The 23 bills on the list dated 30/09/26, with the three advances (invoices 47, 52 and 53).'),
    ('p', 'Invoice and advance dates were not on that list — add them to see the ageing.'),
    ('p', 'Item and buyer are as written: invoice 34 has no buyer, and 34, 44 and 48 show Koyo, Nippon'),
    ('p', 'and KON as the item. Correct them if those are buyers.'),
]
r = 4
for kind, text in lines:
    c = rm.cell(row=r, column=2, value=text)
    if kind == 'h':
        c.font = font(10, bold=True, color=DEEP)
    elif kind == 'b':
        c.font = font(10, color=BLUE_IN, bold=True)
    elif kind == 'y':
        c.font = font(10, color=BLUE_IN, bold=True)
        rm.cell(row=r, column=1).fill = fill(YELLOW)   # a chip in the gutter
    elif kind == 'k':
        c.font = font(10)
    else:
        c.font = font(10)
    r += 1
rm.print_area = f'A1:B{r}'
rm.page_setup.fitToWidth = 1
rm.page_setup.fitToHeight = 1
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
wb.active = wb.index(su)
wb.save(OUT)
print(f'wrote {OUT}: {len(BILLS)} bills, {len(ITEMS)} items, {len(BUYERS)} buyers, '
      f'{BILL_ROWS} register rows')
