#!/usr/bin/env python3
"""Puts the company's carton list into the live register's carton sheets.

    python tools/fill-carton-list.py live.xlsx out.xlsx

Replaces the carton list on 'Carton Stock' (names and opening counts, as
counted on the Count Date) and points every product on 'Packing' at its
carton. Nothing else in the file is changed: the user's row heights, column
widths and edits are kept, and the file is saved through openpyxl with
fullCalcOnLoad so Excel works the figures out on opening.
"""
import sys, re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule

SRC, OUT = sys.argv[1], sys.argv[2]
C0, C1 = 8, 47

# Brand on Stock, carton name, opening count. "Not in stock" cartons open at 0.
CARTONS = [
    ('Blue Dolphin', 'Blue Dolphin PD 100/200', 600),
    ('Blue Dolphin', 'Blue Dolphin PD 200/300', 1600),
    ('Blue Dolphin', 'Blue Dolphin PUD 200/300', 775),
    ('Blue Dolphin', 'Blue Dolphin PUD 300/500', 525),
    ('Primus Japan', 'Primus PD', 650),
    ('Primus Japan', 'Primus PUD', 525),
    ('Primus EU', 'Primus EU', 200),
    ('AMF Brand', 'AMF PD 100/200', 600),
    ('AMF Brand', 'AMF PD 200/300', 1600),
    ('AMF Brand', 'AMF PUD 200/300', 775),
    ('AMF Brand', 'AMF PUD 300/500', 525),
    ('THT Brand', 'THT Brand carton', 0),
    ('China', 'China carton', 0),
    ('Deep Sea', 'Deep Sea carton', 0),
    ('Shrimp Whole', 'Shrimp Whole carton', 0),
    ('Tuna', 'Tuna carton', 0),
    ('Mackerel', 'Mackerel (10kg)', 2200),
    ('Cuttlefish', 'Cuttlefish carton', 0),
    ('White Squid', 'Squid Whole Blast-frozen (6x2kg)', 700),
    ('White Squid', 'Squid Block-frozen (6x2kg)', 1400),
    ('Semi Needle Squid', 'Semi Needle Squid carton', 0),
    ('Octopus', 'Octopus carton', 0),            # the brand added on Summary
    ('Plain', 'Plain Carton', 300),
]

def carton_for(brand, product):
    """(carton, units per carton or None to keep the current figure) for one product."""
    p = product.upper()
    kind = 'PUD' if p.startswith('PUD') else 'PD' if p.startswith('PD') else None
    grade = (re.findall(r'\d+/\d+', p) or [''])[-1]
    if brand in ('Blue Dolphin', 'AMF Brand') and kind:
        return f"{'AMF' if brand == 'AMF Brand' else 'Blue Dolphin'} {kind} {grade}", None
    if brand == 'Primus Japan' and kind:
        return f'Primus {kind}', None
    if brand == 'Primus EU':
        return 'Primus EU', None
    if brand == 'Mackerel':
        return 'Mackerel (10kg)', 10
    if brand == 'White Squid':          # SQW = whole; SQWC = cleaned — to be confirmed by the user
        return ('Squid Block-frozen (6x2kg)' if p.startswith('SQWC') else 'Squid Whole Blast-frozen (6x2kg)'), 6
    return f'{brand} carton', None

wb = load_workbook(SRC)
cs, pk, ce, st = wb['Carton Stock'], wb['Packing'], wb['Carton Entry'], wb['Stock']
names = {n for _, n, _ in CARTONS}

# ---------------------------------------------------------------- Carton Stock: the list
# A merge inside the list (a draft of the user's) would hold two cartons in one
# cell; it is undone and the freed cells take the style of the row's other cells.
from copy import copy
for m in [m for m in cs.merged_cells.ranges if m.min_row >= C0 and m.max_row <= C1]:
    cs.unmerge_cells(str(m))
    for row in range(m.min_row, m.max_row + 1):
        for col in range(m.min_col, m.max_col + 1):
            cs.cell(row=row, column=col)._style = copy(cs.cell(row=C0, column=col)._style)
if len(CARTONS) > C1 - C0 + 1:
    sys.exit('more cartons than rows')
for i in range(C1 - C0 + 1):
    r = C0 + i
    brand, name, opening = CARTONS[i] if i < len(CARTONS) else (None, None, None)
    cs.cell(row=r, column=1).value = brand
    cs.cell(row=r, column=2).value = name
    cs.cell(row=r, column=3).value = opening
    cs.cell(row=r, column=10).value = None          # reorder levels: the user's to set
    # a carton at nil is "Out of stock", not "OK"
    cs.cell(row=r, column=11).value = (
        f'=IF($B{r}="","",IF($C{r}="","Count needed",IF($I{r}<0,"Negative — check",IF($I{r}=0,"Out of stock",'
        f'IF(AND($J{r}<>"",$I{r}<=$J{r}),"Reorder","OK")))))')
cs.conditional_formatting.add(f'K{C0}:K{C1}', FormulaRule(
    formula=[f'$K{C0}="Out of stock"'], font=Font(name='Arial', size=10, italic=True, color='FF5E7C8B')))

# ---------------------------------------------------------------- Packing: every product to its carton
unmapped, mapped = [], {}
for r in range(8, 102):
    brand, product = st.cell(row=r, column=1).value, st.cell(row=r, column=2).value
    if not brand:
        continue
    carton, per = carton_for(brand, str(product))
    if carton not in names:
        unmapped.append((r, brand, product, carton))
        continue
    pk.cell(row=r, column=4).value = carton
    if per is not None:
        pk.cell(row=r, column=5).value = per
    mapped.setdefault(carton, []).append(product)
if unmapped:
    sys.exit(f'products with no carton on the list: {unmapped}')

# ---------------------------------------------------------------- Carton Entry: a rule Excel moved out of reach
# Excel Online stored the "carton not on the list" highlight in an extension
# openpyxl cannot carry, so it is written again.
ce.conditional_formatting.add('B7:B5006', FormulaRule(
    formula=["AND($B7<>\"\",COUNTIF('Carton Stock'!$B$8:$B$47,$B7)=0)"],
    fill=PatternFill(bgColor='FFFDE7E9')))

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print(f'wrote {OUT}: {len(CARTONS)} cartons, {sum(o for _, _, o in CARTONS):,} cartons counted in')
for n in [n for _, n, _ in CARTONS]:
    if n in mapped:
        print(f'  {n:34} {len(mapped[n]):2} products: {", ".join(mapped[n][:4])}{" …" if len(mapped[n]) > 4 else ""}')
    else:
        print(f'  {n:34}  — no product yet')
