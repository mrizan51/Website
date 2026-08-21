# Website

## Stock statement

[`stock-statement/index.html`](stock-statement/index.html) is the daily purchase
and processing statement, rebuilt from the printed sheet. Open it in any browser.

- One column of grades instead of two facing halves, grouped by brand with a
  subtotal after each and a grand total at the foot.
- **Closing slabs** and **stock value** are calculated, never typed, so they cannot
  drift from the columns beside them:
  `closing = opening + production − shipment`, `value = closing × rate per slab`.
- A **rate per slab** column values the stock. Figures are grouped in the Indian
  system and the headline reads in lakh/crore.
- Summary tiles and a by-brand bar chart (slabs or value) sit above the table.
- Grades holding no stock are greyed rather than left blank, so an empty cell
  always means "not recorded" instead of "zero".
- Prints to A4 over two pages with the column header repeated.

The figures ship as recorded on 20/08/26; edits are kept in the browser, and
**Reset to sheet** restores them.

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
