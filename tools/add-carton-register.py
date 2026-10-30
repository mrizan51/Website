#!/usr/bin/env python3
"""Adds a carton stock register to the live stock register, leaving every
existing sheet exactly as it is.

    python tools/add-carton-register.py live.xlsx out.xlsx [--count-date 2026-10-30]

Three sheets are added after 'Daily Report':

  Carton Stock   one row per carton: opening count, received, used, adjusted,
                 closing, reorder level, status and days left
  Carton Entry   one row per delivery of cartons, or an adjustment
  Packing        which carton each product goes into, and how many units fill it

Cartons used are never typed. Production already entered on 'Daily Entry'
takes its cartons off: for each product, the production since the carton count
divided by the units packed per carton, rounded up — a part-filled carton is a
carton used. Rounding is on the running total, so two half cartons on two days
make one carton, not two.

The live file is changed through openpyxl alone — no LibreOffice round trip —
so the user's own edits, styles, protection and logo are kept, and Excel works
the new formulas out when the file opens (fullCalcOnLoad).
"""
import sys, re, datetime as dt
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import PageSetupProperties

args = [a for i, a in enumerate(sys.argv[1:], start=1)
        if not a.startswith('--') and sys.argv[i - 1] != '--count-date']
SRC, OUT = args[0], args[1]
COUNT_ARG = sys.argv[sys.argv.index('--count-date') + 1] if '--count-date' in sys.argv else None

CS, CE, PK = 'Carton Stock', 'Carton Entry', 'Packing'
C0, C1 = 8, 47                 # carton rows on Carton Stock (40)
E0, E1 = 7, 5006               # entry rows on Carton Entry (5,000)
UNLOCK, LOCK = Protection(locked=False), Protection(locked=True)

wb = load_workbook(SRC)
for name in (CS, CE, PK):
    if name in wb.sheetnames:
        sys.exit(f'{SRC} already has a "{name}" sheet — run this on a register without cartons.')
st, de, su = wb['Stock'], wb['Daily Entry'], wb['Summary']

# ---------------------------------------------------------------- read the live layout
# Product rows: whatever 'Products' spans (Stock!$A$8:$A$101 in the live file).
m = re.search(r"\$A\$(\d+):\$A\$(\d+)", wb.defined_names['Products'].attr_text)
P0, P1 = int(m.group(1)), int(m.group(2))
# Daily Entry rows: as far as its product dropdown reaches.
de_last = max(int(re.search(r'(\d+)$', str(d.sqref).split(':')[-1]).group(1))
              for d in de.data_validations.dataValidation if 'B' in str(d.sqref))
DA, DB, DC = (f"'Daily Entry'!${c}$7:${c}${de_last}" for c in 'ABC')
STMT, COUNT = 'Stock!$E$3', f"'{CS}'!$E$4"

products = [(r, st.cell(row=r, column=1).value, st.cell(row=r, column=2).value)
            for r in range(P0, P1 + 1) if st.cell(row=r, column=1).value not in (None, '')]
brands = list(dict.fromkeys(b for _, b, _ in products))

# Case size: the live Summary divides a brand's balance by its units per case in
# "No. of Cases" (=+$D11/10 for the shrimp brands). Those sizes seed Packing.
case = {}
cur = None
for r in range(1, su.max_row + 1):
    a = su.cell(row=r, column=1).value
    if isinstance(a, str) and not a.startswith('=') and su.cell(row=r + 1, column=1).value == 'Product':
        cur = a
        heads = [su.cell(row=r + 1, column=c).value for c in range(1, 12)]
        col = heads.index('No. of Cases') + 1 if 'No. of Cases' in heads else None
        continue
    if cur and col:
        f = su.cell(row=r, column=col).value
        mm = re.fullmatch(r'=\+?\$?D\$?\d+/(\d+(?:\.\d+)?)', str(f or ''))
        if mm and cur not in case:
            case[cur] = float(mm.group(1))

dates = [de.cell(row=r, column=1).value for r in range(7, de_last + 1)]
dates = [d for d in dates if isinstance(d, dt.datetime)]
if COUNT_ARG:
    count_date = dt.datetime.strptime(COUNT_ARG, '%Y-%m-%d')
else:                           # the day after the last production already entered
    count_date = max(d for d in dates if d <= dt.datetime(2100, 1, 1)) + dt.timedelta(days=1)

# ---------------------------------------------------------------- styles, taken from the live sheets
def like(dst, src, **over):
    dst._style = copy(src._style)
    for k, v in over.items():
        setattr(dst, k, v)
    return dst

S = {k: st[k] for k in ('A1', 'A2', 'A3', 'E3', 'F3', 'A5', 'A7', 'O7', 'A8', 'D8', 'F8', 'H8',
                        'J8', 'O8', 'A102', 'K102', 'A104')}
INT = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;#,##0'
DASH = '#,##0;\\-#,##0;\\–'
CTN = '#,##0;\\-#,##0'                  # cartons: can run negative, never reach a lakh
SIGNED = '+#,##0;\\-#,##0;\\–'
blue_txt = Font(name='Arial', size=10, color='FF0000FF')
thin = Side(style='thin', color='FFCBDFE8')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
PINK = PatternFill(bgColor='FFFDE7E9')           # the live register's "fill this in" pink
RED_FONT = Font(name='Arial', size=10, bold=True, color='FFB3261E')
RED_FILL = PatternFill(bgColor='FFF8D7DA')
AMBER = PatternFill(bgColor='FFFCE8B2')
MUTED_FONT = Font(name='Arial', size=10, italic=True, color='FF5E7C8B')

def title(ws, sub):
    like(ws['A1'], S['A1']).value = 'PREMIER EXPORTS INTERNATIONAL'
    like(ws['A2'], S['A2']).value = sub
    ws.row_dimensions[1].height = 18

def header(ws, row, labels, helpers=()):
    for i, t in enumerate(labels, start=1):
        like(ws.cell(row=row, column=i), S['A7']).value = t
    for i, t in helpers:
        like(ws.cell(row=row, column=i), S['O7']).value = t
    ws.row_dimensions[row].height = 30

def typed(c, fmt='General', yellow=False):
    like(c, S['D8'] if yellow else S['A8'], number_format=fmt, protection=UNLOCK)
    if not yellow:
        c.font = blue_txt
    return c

def calc(c, f, fmt=INT, bold=False):
    like(c, S['J8'] if bold else S['H8'], number_format=fmt)
    c.value = f
    return c

def helper(c, f, fmt='General'):
    like(c, S['O8'], number_format=fmt).value = f
    return c

def warn(ws, row, f, muted=False):
    c = like(ws.cell(row=row, column=1), S['A104'])
    c.value = f
    if muted:
        c.font = MUTED_FONT
    return c

def protect(ws):
    p, q = st.protection, ws.protection
    for k in ('sheet', 'objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows',
              'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows',
              'selectLockedCells', 'selectUnlockedCells', 'sort', 'autoFilter', 'pivotTables'):
        setattr(q, k, getattr(p, k))
    if p.password:                     # the live sheets carry none; matched, not added
        q.set_password(p.password, already_hashed=True)
    ws.sheet_view.showGridLines = False

def page(ws, area, titles):
    ws.print_area = area
    ws.print_title_rows = titles
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

after = wb.sheetnames.index('Daily Report') + 1
cs = wb.create_sheet(CS, after)
ce = wb.create_sheet(CE, after + 1)
pk = wb.create_sheet(PK, after + 2)
CEA, CEB, CEC, CED = (f"'{CE}'!${c}${E0}:${c}${E1}" for c in 'ABCD')
PKD, PKG, PKI, PKQ = (f"{PK}!${c}${P0}:${c}${P1}" for c in 'DGIQ')

# ================================================================ Carton Stock
title(cs, 'CARTON STOCK — LIVE POSITION')
like(cs['A3'], S['A3']).value = 'Statement Date'
like(cs['E3'], S['H8'], number_format='dd/mm/yyyy', alignment=Alignment(horizontal='center', vertical='center')).value = f'={STMT}'
cs['E3'].font = Font(name='Arial', size=10, bold=True, color='FF0B2C3A')
like(cs['F3'], S['F3']).value = '← from Stock; change it there'
like(cs['A4'], S['A3']).value = 'Count Date'
like(cs['E4'], S['E3'], number_format='dd/mm/yyyy').value = count_date
like(cs['F4'], S['F3']).value = '← the day the opening cartons were counted, before that day\'s packing'
like(cs['A5'], S['A5']).value = ('Cartons used come from production on Daily Entry and the carton sizes on Packing, '
                                  'rounded up per product. Closing = Opening + Received − Used ± Adjusted.')
header(cs, 7, ['Brand', 'Carton', 'Opening Stock', 'Received Today', 'Used Today', 'Received to Date',
               'Used to Date', 'Adjusted to Date', 'Closing Stock', 'Reorder Level', 'Status', 'Days Left'],
       helpers=[(15, 'Used / day (last 30 days)'), (16, 'Duplicate name')])
like(cs['O4'], S['O7']).value = 'Days averaged'
helper(cs['P4'], f'=MAX(1,MIN(30,{STMT}-{COUNT}+1))', '0')

seed = [(b, f'{b} carton') for b in brands]
for i in range(C1 - C0 + 1):
    r = C0 + i
    b, name = seed[i] if i < len(seed) else (None, None)
    typed(cs.cell(row=r, column=1)).value = b
    typed(cs.cell(row=r, column=2)).value = name
    typed(cs.cell(row=r, column=3), CTN, yellow=True)
    typed(cs.cell(row=r, column=10), CTN, yellow=True)
    k = f'$B{r}'
    calc(cs.cell(row=r, column=4), f'=IF({k}="","",IF({STMT}<{COUNT},0,SUMIFS({CEC},{CEB},{k},{CEA},{STMT})))', DASH)
    calc(cs.cell(row=r, column=5), f'=IF({k}="","",SUMIF({PKD},{k},{PKI}))', DASH)
    calc(cs.cell(row=r, column=6), f'=IF({k}="","",IF({STMT}<{COUNT},0,SUMIFS({CEC},{CEB},{k},{CEA},">="&{COUNT},{CEA},"<="&{STMT})))', CTN)
    calc(cs.cell(row=r, column=7), f'=IF({k}="","",SUMIF({PKD},{k},{PKG}))', CTN)
    calc(cs.cell(row=r, column=8), f'=IF({k}="","",IF({STMT}<{COUNT},0,SUMIFS({CED},{CEB},{k},{CEA},">="&{COUNT},{CEA},"<="&{STMT})))', SIGNED)
    calc(cs.cell(row=r, column=9), f'=IF(OR({k}="",$C{r}=""),"",$C{r}+$F{r}-$G{r}+$H{r})', CTN, bold=True)
    calc(cs.cell(row=r, column=11),
         f'=IF({k}="","",IF($C{r}="","Count needed",IF($I{r}<0,"Negative — check",'
         f'IF(AND($J{r}<>"",$I{r}<=$J{r}),"Reorder","OK"))))', 'General')
    cs.cell(row=r, column=11).alignment = Alignment(horizontal='center', vertical='center')
    calc(cs.cell(row=r, column=12), f'=IF($I{r}="","",IF($I{r}<=0,0,IF(N($O{r})<=0,"—",ROUNDDOWN($I{r}/$O{r},0))))', CTN)
    cs.cell(row=r, column=12).alignment = Alignment(horizontal='right', vertical='center')
    helper(cs.cell(row=r, column=15), f'=IF({k}="","",SUMIF({PKD},{k},{PKQ})/$P$4)', '0.0')
    helper(cs.cell(row=r, column=16), f'=IF({k}="",0,IF(COUNTIF($B${C0}:$B${C1},{k})>1,1,0))', '0')
    cs.row_dimensions[r].height = 15

T = C1 + 1
for c in range(1, 13):
    like(cs.cell(row=T, column=c), S['A102'])
cs.cell(row=T, column=1).value = 'TOTAL (cartons)'
for c in 'CDEFGHI':
    like(cs[f'{c}{T}'], S['K102'], number_format=SIGNED if c == 'H' else (DASH if c in 'DE' else CTN)).value = f'=SUM({c}{C0}:{c}{C1})'
cs[f'K{T}'].value = f'=IF(COUNTIF($K${C0}:$K${C1},"Reorder")=0,"",COUNTIF($K${C0}:$K${C1},"Reorder")&" to reorder")'
cs[f'K{T}'].alignment = Alignment(horizontal='center', vertical='center')
W = T + 2
warns = [
    (f'=IF(SUM({PK}!$R${P0}:$R${P1})=0,"",SUM({PK}!$R${P0}:$R${P1})&" product(s) packed since the count have no carton or no '
     f'Packed per Carton on Packing — their cartons are not taken off. Fill them in on Packing.")', False),
    (f'=IF(SUM({PK}!$S${P0}:$S${P1})=0,"",SUM({PK}!$S${P0}:$S${P1})&" product(s) on Packing name a carton that is not on this list — '
     f'check the spelling.")', False),
    (f'=IF(SUMPRODUCT(({CEB}<>"")*(COUNTIF($B${C0}:$B${C1},{CEB})=0))=0,"",SUMPRODUCT(({CEB}<>"")*(COUNTIF($B${C0}:$B${C1},{CEB})=0))'
     f'&" row(s) on Carton Entry name a carton that is not on this list — they are not counted.")', False),
    (f'=IF(SUM($P${C0}:$P${C1})=0,"","Two rows share a carton name — deliveries and usage would be counted on both. Rename one.")', False),
    (f'=IF({STMT}<{COUNT},"The Statement Date on Stock is before the Count Date — nothing after the count is counted yet.","")', False),
    (f'=IF(SUM({PK}!$T${P0}:$T${P1})=0,"",SUM({PK}!$T${P0}:$T${P1})&" product(s) on Packing have no carton size yet — '
     f'fill them in before they are packed.")', True),
]
for i, (f, muted) in enumerate(warns):
    warn(cs, W + i, f, muted)

for col, w in zip('ABCDEFGHIJKL', (18, 26, 12, 11, 11, 12, 12, 12, 13, 11, 20, 9)):
    cs.column_dimensions[col].width = w
cs.column_dimensions['N'].width = 2
cs.column_dimensions['O'].width = 13
cs.column_dimensions['P'].width = 11
cs.freeze_panes = f'C{C0}'
cs.auto_filter.ref = f'A7:L{C1}'
page(cs, f'A1:L{W + len(warns) - 1}', '7:7')
cs.page_setup.fitToHeight = 1

cs.conditional_formatting.add(f'C{C0}:C{C1}', FormulaRule(formula=[f'AND($B{C0}<>"",$C{C0}="")'], fill=PINK))
cs.conditional_formatting.add(f'I{C0}:I{C1}', CellIsRule(operator='lessThan', formula=['0'], font=RED_FONT, fill=RED_FILL))
cs.conditional_formatting.add(f'K{C0}:K{C1}', FormulaRule(formula=[f'$K{C0}="Negative — check"'], font=RED_FONT, fill=RED_FILL, stopIfTrue=True))
cs.conditional_formatting.add(f'K{C0}:K{C1}', FormulaRule(formula=[f'$K{C0}="Reorder"'],
                              font=Font(name='Arial', size=10, bold=True, color='FF7A4B00'), fill=AMBER, stopIfTrue=True))
cs.conditional_formatting.add(f'K{C0}:K{C1}', FormulaRule(formula=[f'$K{C0}="Count needed"'], fill=PINK, font=MUTED_FONT))

def dv(ws, ref, **k):
    v = DataValidation(allow_blank=True, showErrorMessage=True, **k)
    ws.add_data_validation(v)
    v.add(ref)
dv(cs, f'C{C0}:C{C1}', type='whole', operator='greaterThanOrEqual', formula1='0',
   error='Opening stock is a number of cartons, 0 or more.')
dv(cs, f'J{C0}:J{C1}', type='whole', operator='greaterThanOrEqual', formula1='0',
   error='Reorder level is a number of cartons, 0 or more.')
dv(cs, 'E4', type='date', operator='greaterThanOrEqual', formula1='DATE(2000,1,1)', error='Enter the count date.')
cs['E4'].protection = UNLOCK
protect(cs)

wb.defined_names.add(DefinedName('CartonList', attr_text=
    f"OFFSET('{CS}'!$B${C0},0,0,MAX(1,COUNTA('{CS}'!$B${C0}:$B${C1})),1)"))

# ================================================================ Carton Entry
title(ce, 'CARTON ENTRY — ONE ROW PER DELIVERY OR ADJUSTMENT')
like(ce['A4'], S['A5']).value = 'When cartons arrive, add a row: date, carton from the dropdown, number received.'
like(ce['A5'], S['A5']).value = ('Damaged, or a recount differs: Adjustment, minus or plus. Packing is taken off by itself '
                                  'from production — never enter it here.')
like(ce['G4'], de['G4']).value = 'Rows that will not reach Carton Stock:'
like(ce['J4'], de['J4']).value = (f'=SUMPRODUCT((B{E0}:B{E1}<>"")*(COUNTIF(\'{CS}\'!$B${C0}:$B${C1},B{E0}:B{E1})=0))')
like(ce['K4'], de['K4']).value = '← must stay at 0; anything else is a mistyped carton'
header(ce, 6, ['Date', 'Carton (pick from list)', 'Received (cartons)', 'Adjustment (+ / −)', 'Supplier / Invoice No.', 'Note'])
fmts = ['dd/mm/yyyy', 'General', DASH, SIGNED, 'General', 'General']
src = [de.cell(row=7, column=c) for c in (1, 2, 3, 3, 5, 5)]
for r in range(E0, E1 + 1):
    for c in range(1, 7):
        like(ce.cell(row=r, column=c), src[c - 1], number_format=fmts[c - 1], protection=UNLOCK)
for i, w in enumerate((14, 30, 16, 16, 26, 36), start=1):
    ce.column_dimensions[get_column_letter(i)].width = w
    ce.column_dimensions[get_column_letter(i)].protection = UNLOCK
ce.column_dimensions['J'].width = 6
ce.freeze_panes = f'A{E0}'
dv(ce, f'A{E0}:A{E1}', type='date', operator='greaterThanOrEqual', formula1='DATE(2000,1,1)', error='Enter a date, e.g. 30/10/2026.')
dv(ce, f'B{E0}:B{E1}', type='list', formula1='CartonList', error='Pick a carton from the list. New cartons are added on Carton Stock first.')
dv(ce, f'C{E0}:C{E1}', type='whole', operator='greaterThanOrEqual', formula1='0', error='Cartons received: a whole number, 0 or more.')
dv(ce, f'D{E0}:D{E1}', type='whole', operator='between', formula1='-1000000', formula2='1000000',
   error='Adjustment: a whole number of cartons, minus for damaged, plus for found.')
ce.conditional_formatting.add(f'B{E0}:B{E1}', FormulaRule(
    formula=[f'AND($B{E0}<>"",COUNTIF(\'{CS}\'!$B${C0}:$B${C1},$B{E0})=0)'], fill=PINK))
ce.print_title_rows = '6:6'
protect(ce)

# ================================================================ Packing
title(pk, 'PACKING — WHICH CARTON EACH PRODUCT GOES INTO')
like(pk['A4'], S['A5']).value = ('Set once per product: the carton it is packed in, and how many units — slabs, or kilos for '
                                  'products counted in kilos — fill one carton. Products come from Stock by themselves.')
like(pk['A5'], S['A5']).value = ('Cartons Used = Production since Count ÷ Packed per Carton, rounded up: a part-filled carton is a '
                                  'carton used.')
header(pk, 7, ['Brand', 'Product', 'Unit', 'Carton (pick from list)', 'Packed per Carton', 'Production since Count',
               'Cartons Used', "Today's Production", 'Cartons Used Today'],
       helpers=[(15, 'Key (do not edit)'), (16, 'Production, last 30 days'), (17, 'Cartons, last 30 days'),
                (18, 'Packed, no carton'), (19, 'Carton not on list'), (20, 'No carton set')])
brand_of = {r: b for r, b, _ in products}
for r in range(P0, P1 + 1):
    a = f'Stock!$A{r}'
    like(pk.cell(row=r, column=1), S['A8']).value = f'=IF({a}="","",{a})'
    like(pk.cell(row=r, column=2), S['A8']).value = f'=IF({a}="","",Stock!$B{r})'
    like(pk.cell(row=r, column=3), S['A8']).value = f'=IF({a}="","",IF(Stock!$C{r}="","",Stock!$C{r}))'
    b = brand_of.get(r)
    typed(pk.cell(row=r, column=4)).value = f'{b} carton' if b else None
    typed(pk.cell(row=r, column=5), '#,##0.##', yellow=True).value = case.get(b) if b else None
    k = f'$O{r}'
    calc(pk.cell(row=r, column=6), f'=IF($A{r}="","",IF({STMT}<{COUNT},0,SUMIFS({DC},{DB},{k},{DA},">="&{COUNT},{DA},"<="&{STMT})))')
    calc(pk.cell(row=r, column=7), f'=IF(OR($A{r}="",$D{r}="",N($E{r})<=0),"",ROUNDUP($F{r}/$E{r},0))', CTN, bold=True)
    calc(pk.cell(row=r, column=8), f'=IF($A{r}="","",IF({STMT}<{COUNT},0,SUMIFS({DC},{DB},{k},{DA},{STMT})))', DASH)
    calc(pk.cell(row=r, column=9), f'=IF($G{r}="","",$G{r}-ROUNDUP(($F{r}-$H{r})/$E{r},0))', DASH)
    helper(pk.cell(row=r, column=15), f'=IF({a}="","",Stock!$O{r})')
    helper(pk.cell(row=r, column=16), f'=IF($A{r}="",0,SUMIFS({DC},{DB},{k},{DA},">="&MAX({COUNT},{STMT}-29),{DA},"<="&{STMT}))', '#,##0')
    helper(pk.cell(row=r, column=17), f'=IF(OR($A{r}="",$D{r}="",N($E{r})<=0),0,$P{r}/$E{r})', '0.0')
    helper(pk.cell(row=r, column=18), f'=IF(AND($A{r}<>"",N($F{r})>0,OR($D{r}="",N($E{r})<=0)),1,0)', '0')
    helper(pk.cell(row=r, column=19), f'=IF(AND($A{r}<>"",$D{r}<>""),IF(COUNTIF(\'{CS}\'!$B${C0}:$B${C1},$D{r})=0,1,0),0)', '0')
    helper(pk.cell(row=r, column=20), f'=IF(AND($A{r}<>"",OR($D{r}="",N($E{r})<=0)),1,0)', '0')
    pk.row_dimensions[r].height = 15
for col, w in zip('ABCDEFGHI', (18, 26, 8, 26, 11, 13, 11, 12, 12)):
    pk.column_dimensions[col].width = w
pk.column_dimensions['N'].width = 2
for col, w in zip('OPQRST', (32, 12, 12, 10, 10, 10)):
    pk.column_dimensions[col].width = w
pk.freeze_panes = f'C{P0}'
pk.auto_filter.ref = f'A7:I{P1}'
page(pk, f'A1:I{P1}', '7:7')
dv(pk, f'D{P0}:D{P1}', type='list', formula1='CartonList', error='Pick a carton from the list. New cartons are added on Carton Stock first.')
dv(pk, f'E{P0}:E{P1}', type='decimal', operator='greaterThan', formula1='0',
   error='Units in one carton: slabs, or kilos for products counted in kilos.')
pk.conditional_formatting.add(f'D{P0}:E{P1}', FormulaRule(formula=[f'AND($A{P0}<>"",D{P0}="")'], fill=PINK))
pk.conditional_formatting.add(f'D{P0}:D{P1}', FormulaRule(formula=[f'$S{P0}=1'], font=RED_FONT, fill=RED_FILL))
protect(pk)

# ================================================================ Read me
rm = wb['Read me']
last = max(r for r in range(1, rm.max_row + 1) if rm.cell(row=r, column=2).value not in (None, ''))
head_src = next(rm.cell(row=r, column=2) for r in range(1, last + 1)
                if rm.cell(row=r, column=2).value == 'Dates must be real dates')
body_src = rm.cell(row=last, column=2)
lines = [
    ('h', 'Cartons — taken off by production, counted in once'),
    ('p', "Once: on 'Carton Stock' count each carton, type the count in Opening Stock and the day you counted in Count Date."),
    ('p', "    A Reorder Level is optional; at or below it the Status turns amber."),
    ('p', "Once: on 'Packing' check the carton each product goes into and how many slabs (or kilos) fill one carton."),
    ('p', "    The shrimp brands start at 10 per carton, from the No. of Cases column; fill in the others."),
    ('p', "Every day: nothing extra. Production typed on 'Daily Entry' takes its cartons off by itself."),
    ('p', "When cartons arrive: one row on 'Carton Entry' — date, carton, number received."),
    ('p', "    Damaged, or a recount differs: the same row, Adjustment, minus or plus."),
    ('p', "Cartons used are rounded up for each product: a part-filled carton is a carton used."),
]
r = last + 2
for kind, text in lines:
    like(rm.cell(row=r, column=2), head_src if kind == 'h' else body_src).value = text
    r += 1
rm.print_area = f'A1:B{r}'

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print(f'wrote {OUT}: {len(seed)} cartons seeded ({len(brands)} brands), Packing rows {P0}-{P1} '
      f'({len(products)} products; {sum(1 for _, b, _ in products if b in case)} with a case size from Summary), '
      f'count date {count_date:%d/%m/%Y}, Daily Entry read to row {de_last}')
