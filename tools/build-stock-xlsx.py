#!/usr/bin/env python3
"""Builds PEI-Stock-Register.xlsx — a daily stock register driven by formulas.

    python tools/build-stock-xlsx.py stock-statement/PEI-Stock-Register.xlsx

Design note. The paper statement re-keys every opening balance each morning
from yesterday's total, which is both tedious and the easiest place to make a
mistake. Here the opening balance is a FROZEN baseline (the position recorded
on 20/08/2026) and every later movement is a dated row in 'Daily Entry'.
Closing = baseline + production to date - shipment to date, so the register
rolls itself forward and nothing is ever re-typed.
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule

OUT = sys.argv[1] if len(sys.argv) > 1 else 'PEI-Stock-Register.xlsx'

ENTRY_FIRST = 7      # first data row on 'Daily Entry'
ENTRY_MAX = 5000     # how far the SUMIFS reach — roughly a year of entries
ENTRY_FMT = 250      # how many rows are pre-ruled and ready to type into

# ---------------------------------------------------------------- source data
# Opening balance = the "Opening Balance" column of the purchase & processing
# statement dated 20/08/26. Production for that day is logged in 'Daily Entry',
# so Closing reproduces the sheet's own "Total Slabs" column.
STOCK = [
    ('Blue Dolphin', 'PD 100/200 P', 11459), ('Blue Dolphin', 'PD 100/200 K', 517),
    ('Blue Dolphin', 'PUD 200/300 P', 2262), ('Blue Dolphin', 'PUD 200/300 K', 2835),
    ('Blue Dolphin', 'PUD 300/500 P', 1673), ('Blue Dolphin', 'PUD 300/500 K', 7542),
    ('Blue Dolphin', 'PD 200/300 P', 2357),
    ('Primus PUD Japan', 'PUD 40/60 KZN', 64), ('Primus PUD Japan', 'PUD 60/80 KZN', 191),
    ('Primus PUD Japan', 'PUD 80/120 P', 16849), ('Primus PUD Japan', 'PUD 80/120 KZN', 466),
    ('Primus PUD Japan', 'PUD 80/120 N', 69), ('Primus PUD Japan', 'PUD 100/200 P', 2617),
    ('Primus PUD Japan', 'PUD 100/200 N', 13), ('Primus PUD Japan', 'PUD 100/200 KZN', 4209),
    ('Primus PUD Japan', 'PUD 200/300 P', 0), ('Primus PUD Japan', 'PUD 200/300 K', 2757),
    ('Primus PUD Japan', 'PUD 200/300 KZN', 1080), ('Primus PUD Japan', 'PUD 300/500 P', 421),
    ('Primus PUD Japan', 'PUD 300/500 K', 580),
    ('Primus EU', 'PUD 20/40', 185), ('Primus EU', 'PUD 40/60', 368),
    ('Primus EU', 'PUD 60/80', 380), ('Primus EU', 'PUD 80/120', 1151),
    ('Primus EU', 'PUD 100/200', 3797), ('Primus EU', 'PUD 200/300', 4118),
    ('Primus EU', 'PUD 300/500', 1490), ('Primus EU', 'BKN', 15130),
    ('AMF Brand', 'PD 100/200 P', 0), ('AMF Brand', 'PD 100/200 K', 0),
    ('AMF Brand', 'PUD 200/300 P', 1516), ('AMF Brand', 'PUD 200/300 K', 0),
    ('AMF Brand', 'PUD 300/500 P', 18), ('AMF Brand', 'PUD 300/500 K', 0),
    ('THT Brand PUD PVN', '80/120', 1389), ('THT Brand PUD PVN', '100/200', 7055),
    ('THT Brand PUD PVN', '200/300', 4682),
    ('Deepsea', '300/500', 670),
    ('China PUD', '300/500 P', 451), ('China PUD', '300/500 K', 188),
    ('China PUD', '500/800 P', 401), ('China PUD', '500/800 K', 1920),
]
# Day's production recorded on the 20/08/26 statement
DAY1 = '20/08/2026'
ENTRIES = [
    ('Blue Dolphin', 'PD 100/200 P', 1068), ('Blue Dolphin', 'PD 100/200 K', 65),
    ('Blue Dolphin', 'PUD 200/300 P', 267),
    ('Primus PUD Japan', 'PUD 40/60 KZN', 29), ('Primus PUD Japan', 'PUD 60/80 KZN', 39),
    ('Primus PUD Japan', 'PUD 80/120 KZN', 50), ('Primus PUD Japan', 'PUD 80/120 N', 3),
    ('Primus PUD Japan', 'PUD 100/200 KZN', 50), ('Primus PUD Japan', 'PUD 200/300 KZN', 33),
    ('Primus PUD Japan', 'PUD 300/500 P', 156), ('Primus PUD Japan', 'PUD 300/500 K', 596),
    ('Primus EU', 'BKN', 52),
    ('THT Brand PUD PVN', '80/120', 1872), ('THT Brand PUD PVN', '100/200', 600),
]
BRANDS = []
for b, _, _ in STOCK:
    if b not in BRANDS:
        BRANDS.append(b)
BRANDS.append('Shrimp Whole 5x3kgs')      # tracked on the sheet, no stock held

# ---------------------------------------------------------------- style
FONT = 'Arial'
DEEP, BAND, WASH, LINE = '065E7A', 'E7F3F9', 'F2F8FB', 'CBDFE8'
INK, MUTED = '0B2C3A', '5E7C8B'
BLUE_IN = '0000FF'          # hardcoded input
YELLOW = 'FFFF00'           # fill this in

thin = Side(style='thin', color=LINE)
box = Border(left=thin, right=thin, top=thin, bottom=thin)

# Indian grouping (lakh / crore). Two conditions plus a default is the maximum.
N_IND = r'[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;#,##0'
N_DASH = r'#,##0;-#,##0;"–"'                       # small movements; 0 reads as –
CUR_IND = r'[>=10000000]"₹" ##\,##\,##\,##0.00;[>=100000]"₹" ##\,##\,##0.00;"₹" #,##0.00'
RATE = r'#,##0.00'

def head(ws, row, cols, widths=None):
    for i, t in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = Font(name=FONT, size=9, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=DEEP)
        c.alignment = Alignment(horizontal='right' if i > 2 else 'left',
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
    ('h', 'The daily routine'),
    ('p', "After each day's production, open 'Daily Entry' and add one row per grade that moved."),
    ('p', 'Pick the item from the dropdown, type the date, and enter slabs produced and/or shipped.'),
    ('p', "That is the whole job — 'Stock' and 'Summary' update themselves."),
    ('n', ''),
    ('h', 'Why you never re-type an opening balance'),
    ('p', 'Opening balance is a fixed baseline: the position recorded on 20/08/2026.'),
    ('p', 'Closing slabs = opening balance + production to date − shipment to date,'),
    ('p', "counted from every row in 'Daily Entry'. So the register rolls forward on its own,"),
    ('p', "and yesterday's closing is today's starting point without anyone copying a number."),
    ('n', ''),
    ('h', 'Which cells you edit'),
    ('b', "Blue figures — typed by you. Opening balance on 'Stock'; everything on 'Daily Entry'."),
    ('y', 'Yellow cells — rate per slab. Fill these in to value the stock.'),
    ('k', 'Black figures — calculated. Do not type over them, or the sheet stops adding up.'),
    ('g', 'Grey — the Key column, which links a stock row to its entries. Leave it alone.'),
    ('n', ''),
    ('h', 'A worked entry'),
    ('p', "'Daily Entry' already holds the production recorded on 20/08/2026 — 14 rows,"),
    ('p', "4,880 slabs. Copy the shape of those rows. With them, 'Stock' shows a closing"),
    ('p', 'position of 1,07,750 slabs, which is the total on the paper statement.'),
    ('n', ''),
    ('h', 'Adding a grade'),
    ('p', "Add the row at the bottom of 'Stock' (brand, grade, opening balance), then copy the"),
    ('p', 'formulas down from the row above. The dropdown picks it up automatically.'),
    ('n', ''),
    ('h', 'Source'),
    ('p', 'Opening balances and the 20/08 production are taken from the purchase & processing'),
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
st['C3'] = DAY1
st['C3'].font = Font(name=FONT, size=10, bold=True, color=BLUE_IN)
st['C3'].fill = PatternFill('solid', fgColor=YELLOW)
st['C3'].border = box
st['C3'].alignment = Alignment(horizontal='center')
st['D3'] = "← set this to drive the two \"today\" columns"
st['D3'].font = Font(name=FONT, size=9, italic=True, color=MUTED)
st['A5'] = ("Opening balance is the frozen baseline of 20/08/2026 — never re-key it. "
            "Closing follows from the dated rows on 'Daily Entry'.")
st['A5'].font = Font(name=FONT, size=9, italic=True, color=MUTED)

HDR = 7
head(st, HDR,
     ['Brand', 'Count / grade', 'Opening balance', "Today's production", "Today's shipment",
      'Production to date', 'Shipment to date', 'Closing slabs', 'Rate / slab (₹)',
      'Stock value (₹)', '', 'Key (do not edit)'],
     [22, 18, 14, 13, 13, 13, 13, 13, 12, 17, 2, 34])
st.cell(row=HDR, column=11).fill = PatternFill('solid', fgColor='FFFFFF')
kc = st.cell(row=HDR, column=12)
kc.fill = PatternFill('solid', fgColor='D9D9D9')
kc.font = Font(name=FONT, size=9, bold=True, color=MUTED)

st.cell(row=HDR, column=3).comment = Comment(
    'Position recorded on the purchase & processing statement of 20/08/26.\n'
    'This is a fixed baseline — do not update it daily. Closing slabs is\n'
    'derived from it plus the Daily Entry log.', 'Premier Exports International', height=90, width=340)

first = HDR + 1
for i, (brand, grade, ob) in enumerate(STOCK):
    r = first + i
    st.cell(row=r, column=1, value=brand).font = Font(name=FONT, size=10, color=INK)
    st.cell(row=r, column=2, value=grade).font = Font(name=FONT, size=10, color=INK)

    c = st.cell(row=r, column=3, value=ob)                       # input
    c.font = Font(name=FONT, size=10, color=BLUE_IN)
    c.number_format = N_IND

    ent = lambda col: f"'Daily Entry'!${col}${ENTRY_FIRST}:${col}${ENTRY_MAX}"
    st.cell(row=r, column=4,
            value=f"=SUMIFS({ent('C')},{ent('B')},$L{r},{ent('A')},$C$3)")
    st.cell(row=r, column=5,
            value=f"=SUMIFS({ent('D')},{ent('B')},$L{r},{ent('A')},$C$3)")
    st.cell(row=r, column=6, value=f"=SUMIFS({ent('C')},{ent('B')},$L{r})")
    st.cell(row=r, column=7, value=f"=SUMIFS({ent('D')},{ent('B')},$L{r})")
    st.cell(row=r, column=8, value=f'=$C{r}+$F{r}-$G{r}')
    rate = st.cell(row=r, column=9)                              # fill in
    rate.fill = PatternFill('solid', fgColor=YELLOW)
    rate.number_format = RATE
    rate.font = Font(name=FONT, size=10, color=BLUE_IN)
    st.cell(row=r, column=10, value=f'=IF($I{r}="","",$H{r}*$I{r})').number_format = CUR_IND
    k = st.cell(row=r, column=12, value=f'=$A{r}&" - "&$B{r}')
    k.font = Font(name=FONT, size=9, color=MUTED)

    for col in (4, 5):
        st.cell(row=r, column=col).number_format = N_DASH
    for col in (6, 7, 8):
        st.cell(row=r, column=col).number_format = N_IND
    for col in range(1, 11):
        cell = st.cell(row=r, column=col)
        cell.border = box
        if cell.font.color is None or col in (4, 5, 6, 7, 8, 10):
            cell.font = Font(name=FONT, size=10, color=INK,
                             bold=(col == 8))
        if i % 2 and col != 9:            # leave the yellow rate cells alone
            cell.fill = PatternFill('solid', fgColor=WASH)

last = first + len(STOCK) - 1
tot = last + 1
st.cell(row=tot, column=1, value='GRAND TOTAL').font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
for col in range(1, 11):
    c = st.cell(row=tot, column=col)
    c.fill = PatternFill('solid', fgColor=DEEP)
    c.border = box
    if col >= 3:
        c.font = Font(name=FONT, size=10, bold=True, color='FFFFFF')
        if col != 9:
            L = get_column_letter(col)
            rng = f'{L}{first}:{L}{last}'
            c.value = (f'=IF(SUM({rng})=0,"",SUM({rng}))' if col == 10
                       else f'=SUM({rng})')
            c.number_format = CUR_IND if col == 10 else (N_DASH if col in (4, 5) else N_IND)

st.cell(row=tot + 2, column=1,
        value='Closing slabs = opening balance + production to date − shipment to date        '
              'Stock value = closing slabs × rate per slab').font = \
    Font(name=FONT, size=9, italic=True, color=MUTED)

st.freeze_panes = f'C{first}'
st.auto_filter.ref = f'A{HDR}:J{last}'
# a negative closing means more was shipped than held — flag it loudly
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
title_block(de, 'DAILY ENTRY — ADD ONE ROW PER GRADE THAT MOVED')
de['A4'] = ("After production, add a row below: date, item from the dropdown, slabs produced "
            "and/or shipped. Never delete past rows — they are what the closing position is built from.")
de['A4'].font = Font(name=FONT, size=9, italic=True, color=MUTED)

DH = 6
head(de, DH, ['Date', 'Item (pick from list)', 'Production (slabs)', 'Shipment (slabs)', 'Note'],
     [14, 40, 16, 16, 40])
de.cell(row=DH, column=1).alignment = Alignment(horizontal='left', wrap_text=True)
de.cell(row=DH, column=2).alignment = Alignment(horizontal='left', wrap_text=True)
de.cell(row=DH, column=5).alignment = Alignment(horizontal='left', wrap_text=True)

er = DH + 1
for brand, grade, qty in ENTRIES:
    de.cell(row=er, column=1, value=DAY1).font = Font(name=FONT, size=10, color=BLUE_IN)
    de.cell(row=er, column=2, value=f'{brand} - {grade}').font = Font(name=FONT, size=10, color=BLUE_IN)
    c = de.cell(row=er, column=3, value=qty)
    c.font = Font(name=FONT, size=10, color=BLUE_IN)
    c.number_format = N_DASH
    s = de.cell(row=er, column=4)
    s.font = Font(name=FONT, size=10, color=BLUE_IN)
    s.number_format = N_DASH
    de.cell(row=er, column=5,
            value="Day's production per statement 20/08/26" if er == DH + 1 else '')\
        .font = Font(name=FONT, size=9, italic=True, color=MUTED)
    for col in range(1, 6):
        de.cell(row=er, column=col).border = box
    er += 1

LAST_ENTRY_ROW = ENTRY_FMT
for r in range(er, LAST_ENTRY_ROW + 1):
    for col in range(1, 6):
        cell = de.cell(row=r, column=col)
        cell.border = box
        cell.font = Font(name=FONT, size=10, color=BLUE_IN)
    de.cell(row=r, column=3).number_format = N_DASH
    de.cell(row=r, column=4).number_format = N_DASH

dv_item = DataValidation(type='list', formula1=f'=Stock!$L${first}:$L${last}',
                         allow_blank=True, showDropDown=False)
dv_item.error = 'Pick an item from the dropdown so the entry reaches the Stock sheet.'
dv_item.errorTitle = 'Unknown item'
de.add_data_validation(dv_item)
dv_item.add(f'B{DH+1}:B{ENTRY_MAX}')

dv_qty = DataValidation(type='decimal', operator='greaterThanOrEqual', formula1='0',
                        allow_blank=True)
dv_qty.error = 'Slabs cannot be negative. Record a despatch in the Shipment column.'
dv_qty.errorTitle = 'Negative quantity'
de.add_data_validation(dv_qty)
dv_qty.add(f'C{DH+1}:D{ENTRY_MAX}')

de.freeze_panes = f'A{DH+1}'
de.auto_filter.ref = f'A{DH}:E{LAST_ENTRY_ROW}'
de.print_title_rows = f'{DH}:{DH}'
de.page_setup.fitToWidth = 1
de.page_setup.fitToHeight = 0
de.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ================================================================ Summary
sm = wb.create_sheet('Summary')
title_block(sm, 'SUMMARY BY BRAND')

sm['A4'] = 'Closing stock'
sm['A4'].font = Font(name=FONT, size=9, bold=True, color=MUTED)
sm['A5'] = f'=Stock!$H${tot}'
sm['A5'].font = Font(name=FONT, size=16, bold=True, color=DEEP)
sm['A5'].number_format = N_IND
sm['A6'] = 'slabs'
sm['A6'].font = Font(name=FONT, size=9, color=MUTED)

sm['C4'] = 'Stock value'
sm['C4'].font = Font(name=FONT, size=9, bold=True, color=MUTED)
sm['C5'] = f'=IF(Stock!$J${tot}="","—",Stock!$J${tot})'
sm['C5'].font = Font(name=FONT, size=16, bold=True, color=DEEP)
sm['C5'].number_format = CUR_IND
sm['C6'] = 'at the rates entered on Stock'
sm['C6'].font = Font(name=FONT, size=9, color=MUTED)

SH = 8
head(sm, SH, ['Brand', 'Opening balance', "Today's production", "Today's shipment",
              'Closing slabs', 'Stock value (₹)'], [24, 15, 15, 15, 14, 18])
sr = SH + 1
for i, b in enumerate(BRANDS):
    r = sr + i
    sm.cell(row=r, column=1, value=b).font = Font(name=FONT, size=10, color=INK)
    for col, src in ((2, 'C'), (3, 'D'), (4, 'E'), (5, 'H'), (6, 'J')):
        rng = (f'Stock!$A${first}:$A${last},$A{r},'
               f'Stock!${src}${first}:${src}${last}')
        c = sm.cell(row=r, column=col,
                    value=(f'=IF(SUMIF({rng})=0,"",SUMIF({rng}))' if col == 6
                           else f'=SUMIF({rng})'))
        c.number_format = CUR_IND if col == 6 else (N_DASH if col in (3, 4) else N_IND)
        c.font = Font(name=FONT, size=10, color=INK, bold=(col == 5))
    for col in range(1, 7):
        cell = sm.cell(row=r, column=col)
        cell.border = box
        if i % 2:
            cell.fill = PatternFill('solid', fgColor=WASH)

slast = sr + len(BRANDS) - 1
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
        c.value = (f'=IF(SUM({rng})=0,"",SUM({rng}))' if col == 6 else f'=SUM({rng})')
        c.number_format = CUR_IND if col == 6 else (N_DASH if col in (3, 4) else N_IND)

sm.cell(row=stot + 2, column=1,
        value="Brand totals read straight off 'Stock', so they can never disagree with it.")\
    .font = Font(name=FONT, size=9, italic=True, color=MUTED)

sm.print_area = f'A1:F{stot}'
sm.page_setup.fitToWidth = 1
sm.page_setup.fitToHeight = 0
sm.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

wb.active = wb.index(st)
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
wb.save(OUT)
print(f'wrote {OUT}: {len(STOCK)} grades, {len(ENTRIES)} opening entries, {len(BRANDS)} brands')
