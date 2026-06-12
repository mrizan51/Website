#!/usr/bin/env node
/**
 * Premier Exports International — static site builder
 * Zero dependencies. Node 18+.
 *
 *   node build.mjs        → builds the site into ./dist
 *
 * Sources:
 *   src/pages/*.html      page templates (first line: <!--meta {...}-->)
 *   src/partials/*.html   shared fragments, included with {{> name}}
 *   src/data/site.json    company details (edit contact info here)
 *   src/data/products.json product catalogue (edit products here)
 *   src/assets/**         copied to dist/assets verbatim
 */

import { readFileSync, writeFileSync, mkdirSync, readdirSync, cpSync, rmSync, existsSync } from "node:fs";
import { join, basename } from "node:path";

const ROOT = new URL(".", import.meta.url).pathname;
const SRC = join(ROOT, "src");
const DIST = join(ROOT, "dist");

const site = JSON.parse(readFileSync(join(SRC, "data/site.json"), "utf8"));
const catalogue = JSON.parse(readFileSync(join(SRC, "data/products.json"), "utf8"));

const esc = (s) =>
  String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");

/* ---------- inline icon set (24px line icons) ---------- */
const ICONS = {
  shrimp:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M19.5 7.5c-4.5-3.6-11-2.6-13.6 1.2-2.2 3.2-.9 7.6 2.6 9.3 2.2 1 4.6.9 6.3-.2"/><path d="M19.5 7.5c1.3.4 2.3 1.3 2.7 2.6M19.5 7.5 18 5M19.5 7.5l-3 .4"/><path d="M9.2 7.6c-.7 2.8-.4 5.8 1.6 8.3M13.3 7c-.3 2.7.2 5.4 2 7.8M6.7 9.8c-.3 2.3.4 4.6 2 6.4"/><path d="M14.8 17.8 13 20.5l3.2-.4"/></svg>',
  crab:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="14" rx="5.5" ry="4"/><path d="M8 10.5 5.5 7M5.5 7 4 8.5M5.5 7 7 5.8M16 10.5 18.5 7M18.5 7 20 8.5M18.5 7 17 5.8"/><path d="M6.5 15.5 3.5 17M6.8 13 3 13.5M17.5 15.5l3 1.5M17.2 13l3.8.5"/><path d="M10 13.2h.01M14 13.2h.01"/></svg>',
  squid:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2.5 15.5 9h-7z"/><path d="M9.6 9h4.8v5.5a2.4 2.4 0 0 1-4.8 0z"/><path d="M9.6 14.5c-.4 2.6-1.6 4.4-3.1 5.7M12 16.9V21.5M14.4 14.5c.4 2.6 1.6 4.4 3.1 5.7"/></svg>',
  fish:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3.5 12c2.6-3.6 6.2-5.3 9.8-3.8 2 .9 3.6 2.4 4.7 3.8-1.1 1.4-2.7 2.9-4.7 3.8-3.6 1.5-7.2-.2-9.8-3.8z"/><path d="m18 12 3.5-3-1 3 1 3z"/><path d="M7 11.2h.01"/></svg>',
};

/* ---------- generated fragments ---------- */
const categoryCards = catalogue.categories
  .map((cat) => {
    const examples = cat.items.slice(0, 2).map((i) => i.name).join(", ");
    return `<a class="card cat-card reveal" href="/products.html#${cat.id}">
          <div class="icon-disc">${ICONS[cat.icon] || ICONS.fish}</div>
          <h3>${esc(cat.name)}</h3>
          <p>${esc(cat.intro)}</p>
          <p class="species-note">${cat.items.length} items · incl. ${esc(examples)}</p>
          <span class="card-link">View range</span>
        </a>`;
  })
  .join("\n        ");

const productSections = catalogue.categories
  .map((cat) => {
    const cards = cat.items
      .map((item) => {
        const chips = (item.forms || []).map((f) => `<span class="chip">${esc(f)}</span>`).join("");
        const note = item.note ? `<p class="note">${esc(item.note)}</p>` : "";
        return `<article class="product-card reveal">
            <h3>${esc(item.name)}</h3>
            <p class="sci">${esc(item.sci)}</p>
            <div class="forms">${chips}</div>
            ${note}
          </article>`;
      })
      .join("\n          ");
    return `<section id="${cat.id}" class="product-category" aria-labelledby="${cat.id}-title">
        <div class="cat-head">
          <div class="icon-disc">${ICONS[cat.icon] || ICONS.fish}</div>
          <h2 id="${cat.id}-title">${esc(cat.name)}</h2>
        </div>
        <p class="cat-intro">${esc(cat.intro)}</p>
        <div class="product-grid">
          ${cards}
        </div>
      </section>`;
  })
  .join("\n\n      ");

const marketChips = site.markets.map((m) => `<span class="chip">${esc(m)}</span>`).join("\n            ");
const brandChips = site.brands
  .map((b) => `<span class="brand-chip"><span>●</span> ${esc(b)}</span>`)
  .join("\n          ");
const certChips = site.certifications.map((c) => `<span class="chip">${esc(c)}</span>`).join("");
const addressInline = site.address.lines.map(esc).join("<br>");
const mapsUrl = esc(
  "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(site.address.mapsQuery)
);

const jsonld = `<script type="application/ld+json">${JSON.stringify({
  "@context": "https://schema.org",
  "@type": "Organization",
  name: site.name,
  url: site.domain + "/",
  logo: site.domain + "/assets/img/logo-full.svg",
  foundingDate: String(site.founded),
  email: site.email,
  address: {
    "@type": "PostalAddress",
    streetAddress: "XI/323A, XI/324, Aroor, Chandiroor, NH-66",
    addressLocality: "Cherthala, Alappuzha",
    addressRegion: "Kerala",
    postalCode: "688537",
    addressCountry: "IN",
  },
})}</script>`;

const GENERATED = {
  CATEGORY_CARDS: categoryCards,
  PRODUCT_SECTIONS: productSections,
  MARKET_CHIPS: marketChips,
  BRAND_CHIPS: brandChips,
  CERT_CHIPS: certChips,
  ADDRESS_INLINE: addressInline,
  ADDRESS_BLOCK: addressInline,
  MAPS_URL: mapsUrl,
  PACKING_NOTE: esc(catalogue.packingNote),
  JSONLD: jsonld,
};

/* ---------- template engine ---------- */
const partials = {};
for (const f of readdirSync(join(SRC, "partials"))) {
  partials[basename(f, ".html")] = readFileSync(join(SRC, "partials", f), "utf8");
}

const lookup = (obj, path) => path.split(".").reduce((o, k) => (o == null ? undefined : o[k]), obj);

function render(tpl, page) {
  let html = tpl;
  // expand partials (two passes allows one level of nesting)
  for (let i = 0; i < 2; i++) {
    html = html.replace(/\{\{>\s*([\w-]+)\s*\}\}/g, (_, name) => {
      if (!(name in partials)) throw new Error(`Unknown partial: ${name}`);
      return partials[name];
    });
  }
  // tokens
  html = html.replace(/\{\{\s*([\w.]+)\s*\}\}/g, (match, token) => {
    if (token in GENERATED) return GENERATED[token];
    if (token === "year") return String(new Date().getFullYear());
    if (token.startsWith("site.")) {
      const v = lookup(site, token.slice(5));
      return v == null ? match : esc(v);
    }
    if (token.startsWith("page.")) {
      const v = lookup(page, token.slice(5));
      return v == null ? match : esc(v);
    }
    return match;
  });
  // mark active nav link
  if (page.nav) {
    html = html.replace(`data-nav="${page.nav}"`, `data-nav="${page.nav}" aria-current="page"`);
  }
  return html;
}

/* ---------- build ---------- */
if (existsSync(DIST)) rmSync(DIST, { recursive: true });
mkdirSync(DIST, { recursive: true });

const pages = [];
for (const f of readdirSync(join(SRC, "pages"))) {
  const raw = readFileSync(join(SRC, "pages", f), "utf8");
  const metaMatch = raw.match(/^<!--meta\s+({[\s\S]*?})\s*-->\s*\n?/);
  if (!metaMatch) throw new Error(`${f}: missing <!--meta {...}--> header`);
  const page = JSON.parse(metaMatch[1]);
  page.canonical = site.domain + (page.path === "/index.html" ? "/" : page.path);
  const html = render(raw.slice(metaMatch[0].length), page);

  const leftover = html.match(/\{\{[^}]*\}\}/);
  if (leftover) throw new Error(`${f}: unresolved token ${leftover[0]}`);

  writeFileSync(join(DIST, f), html);
  pages.push(page);
  console.log(`  ✓ ${f}`);
}

cpSync(join(SRC, "assets"), join(DIST, "assets"), { recursive: true });
console.log("  ✓ assets/");

/* sitemap + robots + .nojekyll */
const today = new Date().toISOString().slice(0, 10);
const sitemap = `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${pages
  .filter((p) => p.path !== "/404.html")
  .map(
    (p) => `  <url><loc>${p.canonical}</loc><lastmod>${today}</lastmod><priority>${
      p.path === "/index.html" ? "1.0" : "0.8"
    }</priority></url>`
  )
  .join("\n")}
</urlset>\n`;
writeFileSync(join(DIST, "sitemap.xml"), sitemap);
writeFileSync(join(DIST, "robots.txt"), `User-agent: *\nAllow: /\n\nSitemap: ${site.domain}/sitemap.xml\n`);
writeFileSync(join(DIST, ".nojekyll"), "");
console.log("  ✓ sitemap.xml, robots.txt, .nojekyll");
console.log(`\nBuilt ${pages.length} pages → dist/`);
