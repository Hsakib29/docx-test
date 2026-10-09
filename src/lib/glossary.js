// Glossary matching, placeholder protection and CSV import/export.
const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const BEFORE = "(?<![\\p{L}\\p{M}\\p{N}_])";
const AFTER = "(?![\\p{L}\\p{M}\\p{N}_])";
const BN = "০১২৩৪৫৬৭৮৯";

function matcher(glossary) {
  const terms = glossary.filter((g) => g.src?.trim() && g.tgt?.trim()).sort((a, b) => b.src.trim().length - a.src.trim().length);
  if (!terms.length) return null;
  const re = new RegExp(`${BEFORE}(${terms.map((g) => esc(g.src.trim())).join("|")})${AFTER}`, "giu");
  const lookup = (m) => terms.find((g) => g.src.trim().toLowerCase() === m.toLowerCase());
  return { re, lookup };
}

// Glossary entries whose source term appears in the text.
export function findTerms(text, glossary) {
  const m = matcher(glossary);
  if (!m) return [];
  const seen = new Set();
  const hits = [];
  for (const x of text.matchAll(m.re)) {
    const g = m.lookup(x[1]);
    if (g && !seen.has(g.id)) { seen.add(g.id); hits.push(g); }
  }
  return hits;
}

// Swaps glossary terms for {0}, {1}… so the engine leaves them alone; restoreTerms puts the approved targets in.
export function protectTerms(text, glossary) {
  const m = matcher(glossary);
  if (!m || /\{\s*\d+\s*\}/.test(text)) return { text, map: [] };
  const map = [];
  const out = text.replace(m.re, (hit) => { map.push(m.lookup(hit).tgt.trim()); return `{${map.length - 1}}`; });
  return { text: out, map };
}

// Returns null when a placeholder got lost, so the caller can fall back.
export function restoreTerms(text, map) {
  let found = 0;
  const out = text.replace(/\{\s*([0-9০-৯]+)\s*\}/g, (hit, d) => {
    const i = Number([...d].map((c) => (BN.includes(c) ? BN.indexOf(c) : c)).join(""));
    if (i >= map.length) return hit;
    found++;
    return map[i];
  });
  return found >= map.length ? out : null;
}

export function parseGlossary(raw) {
  const text = raw.replace(/^\uFEFF/, "");
  const first = text.split(/\r?\n/)[0] || "";
  const d = first.includes("\t") ? "\t" : first.includes(";") && !first.includes(",") ? ";" : ",";
  const rows = [];
  let row = [], cur = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += c;
    } else if (c === '"') q = true;
    else if (c === d) { row.push(cur); cur = ""; }
    else if (c === "\n" || c === "\r") { if (c === "\r" && text[i + 1] === "\n") i++; row.push(cur); rows.push(row); row = []; cur = ""; }
    else cur += c;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows
    .filter((r) => r[0]?.trim() && r[1]?.trim())
    .filter((r, i) => !(i === 0 && /^(source|src|term|english)$/i.test(r[0].trim())))
    .map((r) => ({ id: crypto.randomUUID(), src: r[0].trim(), tgt: r[1].trim(), note: (r[2] || "").trim() }));
}

export function glossaryToCsv(glossary) {
  const q = (s) => `"${(s || "").replace(/"/g, '""')}"`;
  return "\uFEFF" + ["source,target,note", ...glossary.map((g) => [g.src, g.tgt, g.note].map(q).join(","))].join("\r\n");
}
