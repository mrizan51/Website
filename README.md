# Premier Exports International — premierexport.in

Corporate website for **Premier Exports International** — EU- and Japan-approved,
HACCP and BRCGS (A Grade) certified exporters of frozen seafood from Chandiroor,
Kerala, India. Founded in 1982 by K. M. Abdulla and A. Musthafa; run today by the
second generation.

Built as a dependency-free static site: hand-crafted HTML/CSS/JS assembled by
a 150-line Node build script. No frameworks, no npm packages, nothing to
maintain or patch.

## Pages

| Page | Purpose |
|---|---|
| `/` | Home — hero, highlights, product categories, brands, markets |
| `/about.html` | Founder story, founders & leadership, values, brands, milestones |
| `/products.html` | Full catalogue: 40 items in 4 categories with scientific & local names |
| `/markets.html` | Export regions (Japan, Europe, Middle East) + who we serve + logistics |
| `/quality.html` | Certifications, HACCP process chain, traceability, testing |
| `/contact.html` | Head office details + request-a-quote form |

## Editing content

Almost everything an owner needs to change lives in two JSON files:

- **`src/data/site.json`** — contact email, phones, address, brands, markets.
  Current inquiry email: `premier.pei@gmail.com` (owner-confirmed). The quote
  form, footer and JSON-LD all read from this file.
- **`src/data/products.json`** — the product catalogue. Add a product by adding
  one object (`name`, `sci`, `forms`, optional `note`) to a category's `items`.

Page copy lives in `src/pages/*.html`; shared header/footer in `src/partials/`;
design system in `src/assets/css/style.css`.

## Build & preview

```bash
node build.mjs                       # builds into dist/
python3 -m http.server 8080 -d dist  # preview at http://localhost:8080
```

The build inlines partials, renders the catalogue from JSON, stamps SEO meta /
canonical / JSON-LD organization markup, and emits `sitemap.xml`, `robots.txt`
and `.nojekyll`.

## Deployment (GitHub Pages — free)

`.github/workflows/deploy.yml` builds and publishes `dist/` on every push.
One-time setup:

1. GitHub repo → **Settings → Pages → Source: GitHub Actions**.
2. Push (or re-run the workflow). The site appears at
   `https://<user>.github.io/<repo>/` for preview.
3. **Custom domain:** in Settings → Pages, set custom domain
   `premierexport.in`, then at your DNS provider point the domain at
   GitHub Pages:
   - `A` records for `premierexport.in` → `185.199.108.153`,
     `185.199.109.153`, `185.199.110.153`, `185.199.111.153`
   - `CNAME` record for `www` → `<user>.github.io`
4. Enable **Enforce HTTPS** once the certificate is issued.

> ⚠️ Switching DNS replaces the site currently served at premierexport.in.
> Preview on the github.io URL first and switch DNS only when ready.

(Cloudflare Pages / Netlify work equally well: build command `node build.mjs`,
output directory `dist`.)

## Quote form

The form is static-host friendly: it opens the visitor's email client with a
pre-filled enquiry to the address in `site.json`. To upgrade to silent inline
submission later, point the form at a HubSpot form, Formspree or similar
endpoint — markup is ready in `src/pages/contact.html`.

## Pre-launch checklist

- [x] Inquiry **email** confirmed: `premier.pei@gmail.com`
- [x] **Phone numbers** confirmed: +91 8138 914941, +91 98463 14941
- [x] **Office address** confirmed: AP X/453, NH66 Highway, Chandiroor P.O., Aroor, Alappuzha, Kerala - 688 537
- [x] **Founders** (1982): K. M. Abdulla, A. Musthafa · **Management**: Salim K. A., Mohammed Rizan Salim, Nissam Musthafa
- [x] **Certifications**: EU, Japan, HACCP, BRCGS (A Grade) — confirmed by owner
- [x] **Markets**: Japan, Europe, Middle East — confirmed by owner
- [ ] Add **EU / BRCGS approval numbers** to `/quality.html` if you want them shown publicly
- [ ] Add **WhatsApp** number if used for trade enquiries (`site.json` → `whatsapp`)
- [ ] **Verify one scientific name**: "Kazhanthan" mapped to *Metapenaeus monoceros* (Speckled Shrimp) — confirm or correct in `products.json`
- [ ] Confirm partner **titles** (currently "Partner" for all three — adjust in `site.json` if specific roles preferred)
- [ ] Review product **forms** (HOSO/HLSO/PD/etc.) per species
- [ ] Replace recreated SVG logo with original vector artwork if available
- [ ] Add real facility/product photography when available

---

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
