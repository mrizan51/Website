# Website

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
