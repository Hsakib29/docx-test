import JSZip from "jszip";

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const XMLNS = "http://www.w3.org/XML/1998/namespace";
export const BN_DIGITS = "০১২৩৪৫৬৭৮৯";
export const toBn = (s) => s.replace(/[0-9]/g, (d) => BN_DIGITS[d]);

/* ---------- DOCX helpers ---------- */
export const kids = (el, name) => [...el.children].filter((c) => c.namespaceURI === W && c.localName === name);
export const norm = (s) => s.replace(/\s+/g, " ").trim();
export const isBn = (t) => /[\u0980-\u09FF]/.test(t);
export const isSkippable = (t) => {
  const c = t.trim();
  return !c || /^\[Image\s+\d+\]$/i.test(c) || /^[0-9\s.,/#!$%^&*;:{}=\-_`~()]+$/.test(c);
};
// JS \b ignores Bengali letters, so whole-word matching uses lookarounds on letters, marks and digits.
export const INFORMAL = [["আপনি", "তুমি"], ["আপনার", "তোমার"], ["আপনাকে", "তোমাকে"]];
export const toInformal = (t) =>
  INFORMAL.reduce((s, [a, b]) => s.replace(new RegExp(`(?<![\\p{L}\\p{M}\\p{N}_])${a}(?![\\p{L}\\p{M}\\p{N}_])`, "gu"), b), t);
export const fileKey = (f) => `${f.name}|${f.size}|${f.lastModified}`;

export function normalizeMem(obj) {
  const out = {};
  for (const [k, v] of Object.entries(obj || {})) {
    const n = norm(k);
    if (n && typeof v === "string" && !Object.hasOwn(out, n)) out[n] = v;
  }
  return out;
}

export function paraText(p) {
  let t = "";
  for (const n of p.getElementsByTagNameNS(W, "*")) {
    if (n.localName === "t") t += n.textContent;
    else if (n.parentNode.localName === "r" && n.localName === "tab") t += "\t";
    else if (n.parentNode.localName === "r" && (n.localName === "br" || n.localName === "cr")) t += "\n";
  }
  return t;
}
export const cellText = (tc) => [...tc.getElementsByTagNameNS(W, "p")].map(paraText).join("\n");
export const collectParas = (body, all) => (all ? [...body.getElementsByTagNameNS(W, "p")] : kids(body, "p"));
export const textRuns = (p) => [...p.getElementsByTagNameNS(W, "r")].filter((r) => kids(r, "t").length);

export async function loadDocx(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const entry = zip.file("word/document.xml");
  if (!entry) throw new Error("Not a valid .docx file");
  const dom = new DOMParser().parseFromString(await entry.async("string"), "application/xml");
  const body = dom.getElementsByTagNameNS(W, "body")[0];
  const rows = [];
  for (const tbl of kids(body, "tbl"))
    for (const tr of kids(tbl, "tr")) {
      const tcs = kids(tr, "tc");
      if (tcs.length >= 2)
        rows.push({ tc0: tcs[0], tc1: tcs[1], src: cellText(tcs[0]).trim(), tgt: cellText(tcs[1]).trim() });
    }
  return { zip, dom, body, rows };
}

// Two-column file: rows with a source and an empty target outweigh the loose paragraphs.
export function detectMode(d) {
  const todo = d.rows.filter((r) => r.src && !r.tgt).length;
  const paras = kids(d.body, "p").map((p) => paraText(p).trim()).filter((t) => !isSkippable(t) && !isBn(t)).length;
  const score = todo > 0 ? todo : d.rows.length / 2;
  return score > 0 && score >= paras ? "table" : "paragraph";
}

export function appendText(dom, r, text) {
  text.split("\n").forEach((line, i) => {
    if (i) r.appendChild(dom.createElementNS(W, "w:br"));
    if (!line) return;
    const t = dom.createElementNS(W, "w:t");
    t.setAttributeNS(XMLNS, "xml:space", "preserve");
    t.textContent = line;
    r.appendChild(t);
  });
}

// Two-column mode: writes into the target cell, cloning size/bold/italic/underline/color from the source cell's first run.
export function writeCell(dom, row, text) {
  const { tc0, tc1 } = row;
  let [p, ...extra] = kids(tc1, "p");
  extra.forEach((x) => x.remove());
  if (!p) { p = dom.createElementNS(W, "w:p"); tc1.appendChild(p); }
  [...p.childNodes].forEach((n) => { if (!(n.namespaceURI === W && n.localName === "pPr")) n.remove(); });

  const r = dom.createElementNS(W, "w:r");
  const srcRun = tc0.getElementsByTagNameNS(W, "r")[0];
  const srcRpr = srcRun && kids(srcRun, "rPr")[0];
  if (srcRpr) {
    const rpr = dom.createElementNS(W, "w:rPr");
    for (const c of srcRpr.children)
      if (["b", "bCs", "i", "iCs", "u", "sz", "szCs", "color", "highlight"].includes(c.localName)) rpr.appendChild(c.cloneNode(true));
    if (rpr.childNodes.length) r.appendChild(rpr);
  }
  appendText(dom, r, text);
  p.appendChild(r);
}

// Paragraph mode: translated words are shared out across the original text runs in proportion to each run's
// original length, so bold/italic stretches stay roughly where they were. Runs holding images are never touched.
export function writeParagraph(dom, ctx, text) {
  if (!ctx.runs.length) {
    const r = dom.createElementNS(W, "w:r");
    ctx.p.appendChild(r);
    ctx.runs = [r]; ctx.weights = [1];
  }
  const words = text.split(" ");
  const total = ctx.weights.reduce((a, b) => a + b, 0);
  let cum = 0, prev = 0;
  ctx.runs.forEach((r, i) => {
    cum += ctx.weights[i];
    const end = i === ctx.runs.length - 1 ? words.length : Math.round((words.length * cum) / total);
    const part = words.slice(prev, end).join(" ");
    prev = end;
    for (const c of [...r.children]) if (c.namespaceURI === W && ["t", "tab", "br", "cr"].includes(c.localName)) c.remove();
    appendText(dom, r, part && end < words.length ? part + " " : part);
  });
}

export async function saveDocx(zip, dom) {
  let xml = new XMLSerializer().serializeToString(dom);
  if (!xml.startsWith("<?xml")) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
  zip.file("word/document.xml", xml);
  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

export function tableItems(d) {
  return d.rows.filter((r) => r.src).map((r) => {
    const write = (t) => writeCell(d.dom, r, t);
    return r.tgt ? { src: r.src, tgt: r.tgt, kept: true, write } : { src: r.src, write };
  });
}

export function paragraphItems(d, { sl, tl, includeTables }) {
  const skipBn = tl === "bn" && sl !== "bn"; // already-Bengali paragraphs are left alone
  const items = [];
  for (const p of collectParas(d.body, includeTables)) {
    const text = paraText(p).trim();
    if (isSkippable(text) || (skipBn && isBn(text))) continue;
    const ctx = { p, runs: textRuns(p) };
    ctx.weights = ctx.runs.map((r) => Math.max(1, kids(r, "t").reduce((n, t) => n + t.textContent.length, 0)));
    items.push({ src: text, write: (t) => writeParagraph(d.dom, ctx, t) });
  }
  return items;
}
