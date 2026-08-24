# Website

## Stock statement

[`stock-statement/index.html`](stock-statement/index.html) is the daily purchase
and processing statement, rebuilt from the printed sheet. Open it in any browser.

- One column of grades instead of two facing halves, grouped by brand with a
  subtotal after each and a grand total at the foot.
- **Closing slabs** and **stock value** are calculated, never typed, so they cannot
  drift from the columns beside them:
  `closing = opening + production − shipment`,
  `value (₹) = closing × rate per slab (US$) × the US$ → ₹ rate`.
- A **rate per slab** column holds the export price in US dollars; with the
  US$ → ₹ rate beside the date, the stock is valued in rupees. Figures are
  grouped in the Indian system and the headline reads in lakh/crore.
- Summary tiles and a by-brand bar chart (slabs or value) sit above the table.
- Grades holding no stock are greyed rather than left blank, so an empty cell
  always means "not recorded" instead of "zero".
- Prints to A4 over two pages with the column header repeated.

The figures ship as recorded on 20/08/26; edits are kept in the browser, and
**Reset to sheet** restores them.

### Daily register (Excel)

The daily loop is four steps, and step 4 is the directors' PDF:

1. On `Daily Entry`, add one row per product that moved — date, product from the
   dropdown, quantity in that product's own unit (slabs, or kilos for tuna and
   mackerel).
2. On `Stock`, set **Statement date** and the **US$ → ₹ exchange rate** for today.
3. Open `Daily Report` — it has already rebuilt itself.
4. **File ▸ Export ▸ Create PDF/XPS**, choose *Selected sheet*, and send it.
   Excel writes the PDF; nothing else needs installing.

`Daily Report` is a one-page A4 sheet carrying the letterhead, the day's
headline figures, the position by brand, and a line for every product that moved
that day — filtered by formula (`MATCH` against a rank column on `Stock`), not by
hand. It prints 20 movement lines; if more moved, a note says how many are not
shown rather than dropping them silently.


[`stock-statement/PEI-Stock-Register.xlsx`](stock-statement/PEI-Stock-Register.xlsx)
is the same stock as a working Excel model — five sheets, 3,262 live formulas,
82 products across 13 brands.

Opening balance is a **frozen baseline** (the position on 20/08/2026), never
re-keyed. Each day's movements go in as dated rows on `Daily Entry`, and

```
Closing Balance = Opening Balance + production to date − shipment to date
```

so the register rolls itself forward. `Stock` also carries two "today" columns
driven by the statement-date cell, which reproduce the paper sheet's *Day's
Production* and *Shipment* columns.

**Units.** Balances are counted in each product's own unit, so a Unit and a
**Kg / Unit** pack size sit on every row:

| | Unit | Kg / Unit |
|---|---|---|
| Blue Dolphin, Primus, AMF, THT, China, Deep Sea | Slab | 2 |
| Tuna, Mackerel | Kg | 1 |
| Squid, cuttlefish, shrimp whole | Slab | varies by buyer |

A slab and a kilo cannot be added together, so **every total in the workbook is
taken in kilos** — `Closing Kg = Closing Balance × Kg / Unit`. A product with no
Kg / Unit is shaded pink, left out of the totals, and counted in a red line on
`Stock`, `Summary` and the `Daily Report` until it is filled in.

**Valuation.** Rate / Kg is the export price per kilo, in US dollars:

```
Stock Value (₹) = Closing Kg × Rate / Kg (US$) × Stock!E4
```

`Stock!E4` is the US$ → ₹ rate, one cell that revalues the register when it
changes. Blank, and the value column stays blank rather than showing a wrong
figure; the report reads *"Exchange rate not set"* and otherwise states the rate
it used. `Summary` carries the same stock in dollars.

**`Summary` is a valuation table per brand** — every product of that brand with
its unit, pack size, closing balance, closing kg, export rate and rupee value,
then a brand subtotal, then a grand total. Each table has 3 spare slots, and
there are 2 spare brand tables.

**Protection.** Every sheet is protected and formatting is locked, so widths,
fonts, colours and number formats cannot be changed by accident. Only two places
accept typing:

| Where | What |
|---|---|
| `Daily Entry` | the whole grid — date, product, production, shipment, note |
| `Stock` | Statement date, exchange rate, Unit, Kg / Unit and Rate / Kg |

Everything else — headings, opening balances, formulas, totals, `Summary` and
`Daily Report` — is locked, as is the workbook structure (no adding, renaming or
deleting sheets). Filtering still works. Excel sheet protection is a guardrail
against accidents, **not security**: the password is trivially removable by
anyone determined.

**Adding a product or a brand** needs no unprotecting. `Stock` carries 25 blank
rows under the last product; every formula on them is already in place.

- **New product** — type Brand, Product, Unit, Kg / Unit and Opening Balance into
  the first blank row on `Stock`. The `Daily Entry` dropdown reads from a defined
  name (`OFFSET`/`COUNTA`), so it grows by itself and never shows blank options,
  and the product appears in its brand's `Summary` table automatically.
- **New brand** — do the above, then type the brand name into a blank brand band
  on `Summary` and a blank row on the `Daily Report` brand table. Until you do, a
  red line on both says the totals understate the stock; both clear themselves.

That warning compares the brand totals against the `Stock` grand total in kilos,
so a brand on `Stock` with no table can never leave silently in a report.

**Rebuilding is three steps, in order:**

```
python tools/build-stock-xlsx.py out.xlsx --from <live.xlsx>   # 1. build
python <xlsx-skill>/scripts/recalc.py out.xlsx 240             # 2. cache values
python tools/build-stock-xlsx.py --relock out.xlsx             # 3. restore locks
```

Step 2 is not optional — openpyxl writes formulas without cached values, so an
unrecalculated workbook reads as empty to pandas and most previewers. Step 3 is
not optional either: LibreOffice silently drops the workbook-structure lock, the
column-level unlocking on `Daily Entry`, and the data-validation ranges when it
recalculates. Re-saving through openpyxl to restore them would strip the cached
values again, so `--relock` patches the XML in place instead.

## Lot analysis

[`lot-analysis/PEI-Lot-Analysis.xlsx`](lot-analysis/PEI-Lot-Analysis.xlsx) is the
step after the stock register: yield and cost for each **lot of raw material**
bought, and a one-page report for the directors. Seven sheets, 4,823 live
formulas, room for 200 lots and 2,000 grade lines.

The loop per lot is four steps, and step 4 is the directors' PDF:

1. On `Lots`, add a row — lot no, statement and purchase dates, party, raw kg,
   raw amount, peeling ₹/kg and the US$ → ₹ rate. Those eight are all you type;
   everything from Peeling Amount rightwards is calculated.
2. On `Lot Entry`, add one row per grade produced — slabs, net kg per slab,
   gross weight per slab and the sale price per kg. Lot no, spec and grade come
   from dropdowns that read the `Lots` and `Grades` sheets, so they grow by
   themselves.
3. Open `Lot Report` and put the lot number in the yellow cell **F5**.
4. **File ▸ Export ▸ Create PDF/XPS**, choose *Selected sheet*, and send it.

**Two weights per slab, and they are not the same.** Net kg per slab (usually 2)
is what the buyer pays for; gross weight per slab (2.150, 2.200 …) is the final
weight. Yield is measured on the gross — F.WT — while the sale is on the net, so
the workbook carries both rather than deriving one from the other:

```
Packed kg = slabs × net kg per slab        Amount (US$) = packed kg × price per kg
F.WT (kg) = slabs × gross wt per slab      Yield %      = total F.WT ÷ raw kg
```

**Cost and profit.** Peeling is charged on the raw kg bought, expenses on the
packed kg produced, and the profit per kg is over the **raw** kg — not the
finished kg, which would flatter a bad yield:

```
Total raw cost = raw amount + peeling ₹/kg × raw kg
Expenses       = expense ₹/kg × packed kg
Net profit     = sales value ₹ − total raw cost − expenses
```

**Expenses.** `Expenses` holds the standard rate per kg for each of 13 heads,
used by every lot. A lot that differs — a different freight, say — overrides any
single head in the grey columns at the right of `Lots`; leave an override blank
and the standard rate applies. The report always prints the rate it actually
used, so an override is never invisible.

`Lot Report` is one A4 page: letterhead, the purchase and the day's headline
figures, four KPI tiles, production by grade, the expense build-up and the
result block. It prints 12 grade lines; if a lot has more, a note says how many
are not shown rather than dropping them silently. `Summary` lists every lot with
its yield and profit, and totals them — with the all-lots yield weighted by raw
kg, not averaged.

**Protection** works as on the stock register. Only `Lots` (the eight blue
columns and the expense overrides), the `Lot Entry` grid, the `Grades` and
`Expenses` master lists, and the lot number on `Lot Report` accept typing;
everything else, including formatting and the workbook structure, is locked.

**Against the printed statement.** The workbook is seeded from lot C 1374 and
reproduces every figure on that sheet — ₹15,27,705 raw cost, 49.20% yield,
$18,581.00 sales, ₹37,156.80 profit, ₹6.20 per raw kg — except the expense
block. The rates printed there are rounded to two decimals and the tool that
produced it used slightly different underlying rates, so its total is
₹1,07,433.37 where these rates multiply out to ₹1,07,428.20. The ₹5.17
difference carries into the profit; this workbook is arithmetically consistent.

**Rebuilding is the same three steps as the register:**

```
python tools/build-lot-xlsx.py lot-analysis/PEI-Lot-Analysis.xlsx
python <xlsx-skill>/scripts/recalc.py lot-analysis/PEI-Lot-Analysis.xlsx 240
python tools/build-lot-xlsx.py --relock lot-analysis/PEI-Lot-Analysis.xlsx
```

## Purchase order

[`purchase-order/index.html`](purchase-order/index.html) is a self-contained,
print-ready purchase order form for Premier Exports International. Open the file
in any browser — no build step, no network access, nothing to install.

- Every field on the sheet is editable in place; line totals, GST, round-off and
  the amount in words (Indian lakh/crore wording) recalculate as you type.
- Supply type switches between **CGST + SGST** (within Kerala), **IGST**
  (other states) and **no GST** (export); the combined rate carries across.
- Prints to A4 via the browser's *Save as PDF*. A two-line order fits one page;
  longer orders flow over with the table header repeated, and the terms and
  signature blocks kept whole.
- The working draft is kept in the browser's local storage, so a reload or an
  accidental close does not lose the order.

[`purchase-order/PEI-Purchase-Order.docx`](purchase-order/PEI-Purchase-Order.docx)
is the same form as a Word document, for when an order has to be edited in Word
or sent as an attachment. It carries the same letterhead, palette, type scale and
terms, laid out on one A4 page with five blank item rows. Word does not
recalculate the totals — type them, or use the HTML form and save that as a PDF.

Rebuild it after changing the design with:

```
node tools/build-docx.js purchase-order/PEI-Purchase-Order.docx
```

## Claude Code skills

### 3d-asset-generator

A project skill that generates 3D assets procedurally from text descriptions
and exports web-ready `.glb` (glTF binary) or `.obj` files. Pure Python
standard library — nothing to install.

Lives in [`.claude/skills/3d-asset-generator/`](.claude/skills/3d-asset-generator/):
a `SKILL.md` workflow plus a mesh toolkit (`meshkit.py`), a GLB validator,
a GLB → PNG software renderer for visual self-checks, a drag-and-drop
Three.js viewer, and worked examples.

Use it in a Claude Code session in this repo — either ask naturally:

> make me a low-poly pine tree as a GLB

or invoke it directly:

```
/3d-asset-generator a toy rocket with three fins
```
