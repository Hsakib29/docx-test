import { norm, isBn, isSkippable, toBn, toInformal, loadDocx, detectMode, tableItems, paragraphItems, paraText, collectParas } from "./docx";
import { protectTerms, restoreTerms } from "./glossary";

/* ---------- Translation ---------- */
// Unofficial Google endpoint (same engine deep_translator uses). For production, route this through your own proxy.
export async function gtx(text, sl, tl) {
  const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${sl}&tl=${tl}&dt=t&q=${encodeURIComponent(text)}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP " + res.status);
  return (await res.json())[0].map((s) => s[0]).join("");
}

export function chunk(line, max = 1500) {
  if (line.length <= max) return [line];
  const out = []; let cur = "";
  for (const s of line.split(/(?<=[.!?।])\s+/)) {
    if (cur && (cur + " " + s).length > max) { out.push(cur); cur = s; }
    else cur = cur ? cur + " " + s : s;
  }
  if (cur) out.push(cur);
  return out;
}

export async function translateText(text, sl, tl) {
  const lines = text.split("\n");
  for (let i = 0; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const parts = [];
    for (const c of chunk(lines[i])) {
      for (let attempt = 0; ; attempt++) {
        try { parts.push(await gtx(c, sl, tl)); break; }
        catch (e) { if (attempt >= 2) throw e; await new Promise((r) => setTimeout(r, 600 * (attempt + 1))); }
      }
    }
    lines[i] = parts.join(" ");
  }
  return lines.join("\n");
}

export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Shared by both modes: items are { src, write(text), kept?, tgt? } in document order.
export async function translateItems(items, opts, mem, onTick) {
  const { tl, localize, mtpe } = opts;
  const bn = tl === "bn";
  const todo = items.filter((x) => !x.kept);
  const srcByKey = new Map();
  todo.forEach((x) => { const k = norm(x.src); if (!srcByKey.has(k)) srcByKey.set(k, x.src); });
  const fromMemory = new Set([...srcByKey.keys()].filter((k) => Object.hasOwn(mem, k)));
  const need = [...srcByKey.keys()].filter((k) => !fromMemory.has(k));
  const failed = new Set();
  const fresh = {}; // MTPE mode: machine output stays out of memory until a human confirms it
  let next = 0, done = 0;
  const worker = async () => {
    while (next < need.length) {
      const k = need[next++];
      try {
        const t = await machineTranslate(srcByKey.get(k), opts);
        (mtpe ? fresh : mem)[k] = t;
      } catch { failed.add(k); }
      onTick(++done, need.length);
      await sleep(120 + Math.random() * 180);
    }
  };
  await Promise.all(Array.from({ length: Math.min(4, need.length) }, worker));

  return items.map((x) => {
    if (x.kept) { return { src: x.src, tgt: x.tgt, status: "kept", ok: false, write: x.write }; }
    const k = norm(x.src);
    let out = failed.has(k) ? "[Translation Error]" : Object.hasOwn(mem, k) ? mem[k] : fresh[k];
    if (localize && bn) out = toBn(out);
    x.write(out);
    return { src: x.src, tgt: out, status: failed.has(k) ? "error" : fromMemory.has(k) ? "tm" : "mt", ok: false, write: x.write };
  });
}

export async function processFile(file, opts, mem, onTick) {
  const d = await loadDocx(file);
  const mode = opts.mode === "auto" ? detectMode(d) : opts.mode;
  const items = mode === "table" ? tableItems(d) : paragraphItems(d, opts);
  const segs = await translateItems(items, opts, mem, onTick);
  return { zip: d.zip, dom: d.dom, segs, mode };
}

export async function addSamples(files, mem) {
  let added = 0;
  for (const f of files) {
    try {
      const { rows } = await loadDocx(f);
      for (const r of rows) {
        const k = norm(r.src);
        if (k && r.tgt && !Object.hasOwn(mem, k)) { mem[k] = r.tgt; added++; }
      }
    } catch { /* skip unreadable sample */ }
  }
  return added;
}

// Paragraph files: pairs "Name.docx" with its edited "Name_translated.docx" and learns every changed paragraph.
export async function learnFromPairs(files, mem, { includeTables }) {
  const byName = new Map(files.map((f) => [f.name.toLowerCase(), f]));
  let learned = 0, pairs = 0;
  const issues = [];
  for (const f of files) {
    if (!/_translated\.docx$/i.test(f.name)) continue;
    const orig = byName.get(f.name.toLowerCase().replace("_translated", ""));
    if (!orig) { issues.push(`${f.name}: its original English file wasn't included.`); continue; }
    try {
      const a = await loadDocx(orig), b = await loadDocx(f);
      const pa = collectParas(a.body, includeTables), pb = collectParas(b.body, includeTables);
      if (pa.length !== pb.length) { issues.push(`${f.name}: paragraph count differs from the original, so it was skipped.`); continue; }
      pairs++;
      pa.forEach((p, i) => {
        const src = norm(paraText(p)), tgt = norm(paraText(pb[i]));
        if (src && tgt && !isBn(src) && isBn(tgt) && !isSkippable(src) && mem[src] !== tgt) { mem[src] = tgt; learned++; }
      });
    } catch (e) { issues.push(`${f.name}: ${e.message}`); }
  }
  return { learned, pairs, issues };
}

// One machine translation, optionally keeping glossary terms. Falls back to a plain translation
// if the engine mangles a term placeholder, so a term is never silently garbled.
export async function machineTranslate(src, { sl, tl, informal, glossary = [], enforceTerms }) {
  let t = null;
  if (enforceTerms && glossary.length) {
    const p = protectTerms(src, glossary);
    if (p.map.length) t = restoreTerms(await translateText(p.text, sl, tl), p.map);
  }
  if (t == null) t = await translateText(src, sl, tl);
  return informal && tl === "bn" ? toInformal(t) : t;
}
