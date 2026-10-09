// DocxTranslator.jsx — browser port of translate_snake (two-column tables) and megatrans (plain paragraphs),\n// with auto-detection, a segment post-editor and an MTPE review mode.
// Setup:  npm i jszip   →   import DocxTranslator from "./DocxTranslator";  <DocxTranslator />
import { useState, useRef, useEffect } from "react";
import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const XMLNS = "http://www.w3.org/XML/1998/namespace";
const LANGS = { en: "English", bn: "Bengali", hi: "Hindi", ur: "Urdu", ar: "Arabic", es: "Spanish", fr: "French", de: "German" };
const BN_DIGITS = "০১২৩৪৫৬৭৮৯";
const toBn = (s) => s.replace(/[0-9]/g, (d) => BN_DIGITS[d]);
const reviewed = (x) => x.status === "edited" || x.ok;
const STATUS = { tm: "From memory", mt: "Machine", error: "Failed", kept: "Already translated", edited: "Edited" };

/* ---------- DOCX helpers ---------- */
const kids = (el, name) => [...el.children].filter((c) => c.namespaceURI === W && c.localName === name);
const norm = (s) => s.replace(/\s+/g, " ").trim();
const isBn = (t) => /[\u0980-\u09FF]/.test(t);
const isSkippable = (t) => {
  const c = t.trim();
  return !c || /^\[Image\s+\d+\]$/i.test(c) || /^[0-9\s.,/#!$%^&*;:{}=\-_`~()]+$/.test(c);
};
// JS \b ignores Bengali letters, so whole-word matching uses lookarounds on letters, marks and digits.
const INFORMAL = [["আপনি", "তুমি"], ["আপনার", "তোমার"], ["আপনাকে", "তোমাকে"]];
const toInformal = (t) =>
  INFORMAL.reduce((s, [a, b]) => s.replace(new RegExp(`(?<![\\p{L}\\p{M}\\p{N}_])${a}(?![\\p{L}\\p{M}\\p{N}_])`, "gu"), b), t);
const fileKey = (f) => `${f.name}|${f.size}|${f.lastModified}`;

function normalizeMem(obj) {
  const out = {};
  for (const [k, v] of Object.entries(obj || {})) {
    const n = norm(k);
    if (n && typeof v === "string" && !Object.hasOwn(out, n)) out[n] = v;
  }
  return out;
}

function paraText(p) {
  let t = "";
  for (const n of p.getElementsByTagNameNS(W, "*")) {
    if (n.localName === "t") t += n.textContent;
    else if (n.parentNode.localName === "r" && n.localName === "tab") t += "\t";
    else if (n.parentNode.localName === "r" && (n.localName === "br" || n.localName === "cr")) t += "\n";
  }
  return t;
}
const cellText = (tc) => [...tc.getElementsByTagNameNS(W, "p")].map(paraText).join("\n");
const collectParas = (body, all) => (all ? [...body.getElementsByTagNameNS(W, "p")] : kids(body, "p"));
const textRuns = (p) => [...p.getElementsByTagNameNS(W, "r")].filter((r) => kids(r, "t").length);

async function loadDocx(file) {
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
function detectMode(d) {
  const todo = d.rows.filter((r) => r.src && !r.tgt).length;
  const paras = kids(d.body, "p").map((p) => paraText(p).trim()).filter((t) => !isSkippable(t) && !isBn(t)).length;
  const score = todo > 0 ? todo : d.rows.length / 2;
  return score > 0 && score >= paras ? "table" : "paragraph";
}

function appendText(dom, r, text) {
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
function writeCell(dom, row, text) {
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
function writeParagraph(dom, ctx, text) {
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

async function saveDocx(zip, dom) {
  let xml = new XMLSerializer().serializeToString(dom);
  if (!xml.startsWith("<?xml")) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
  zip.file("word/document.xml", xml);
  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

/* ---------- Translation ---------- */
// Unofficial Google endpoint (same engine deep_translator uses). For production, route this through your own proxy.
async function gtx(text, sl, tl) {
  const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${sl}&tl=${tl}&dt=t&q=${encodeURIComponent(text)}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP " + res.status);
  return (await res.json())[0].map((s) => s[0]).join("");
}

function chunk(line, max = 1500) {
  if (line.length <= max) return [line];
  const out = []; let cur = "";
  for (const s of line.split(/(?<=[.!?।])\s+/)) {
    if (cur && (cur + " " + s).length > max) { out.push(cur); cur = s; }
    else cur = cur ? cur + " " + s : s;
  }
  if (cur) out.push(cur);
  return out;
}

async function translateText(text, sl, tl) {
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

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Shared by both modes: items are { src, write(text), kept?, tgt? } in document order.
async function translateItems(items, { sl, tl, localize, informal, mtpe }, mem, onTick) {
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
        let t = await translateText(srcByKey.get(k), sl, tl);
        if (informal && bn) t = toInformal(t);
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

function tableItems(d) {
  return d.rows.filter((r) => r.src).map((r) => {
    const write = (t) => writeCell(d.dom, r, t);
    return r.tgt ? { src: r.src, tgt: r.tgt, kept: true, write } : { src: r.src, write };
  });
}

function paragraphItems(d, { sl, tl, includeTables }) {
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

async function processFile(file, opts, mem, onTick) {
  const d = await loadDocx(file);
  const mode = opts.mode === "auto" ? detectMode(d) : opts.mode;
  const items = mode === "table" ? tableItems(d) : paragraphItems(d, opts);
  const segs = await translateItems(items, opts, mem, onTick);
  return { zip: d.zip, dom: d.dom, segs, mode };
}

async function addSamples(files, mem) {
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
async function learnFromPairs(files, mem, { includeTables }) {
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

function download(blob, name) {
  const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: name });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

/* ---------- UI ---------- */
function Drop({ title, hint, files, onFiles, badge, prepare }) {
  const input = useRef();
  const [over, setOver] = useState(false);
  const pick = async (list) => onFiles(await prepare([...list]));
  return (
    <div
      className={"dt-drop" + (over ? " is-over" : "")}
      role="button" tabIndex={0}
      onClick={() => input.current.click()}
      onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && input.current.click()}
      onDragOver={(e) => { e.preventDefault(); setOver(true); }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => { e.preventDefault(); setOver(false); pick(e.dataTransfer.files); }}
    >
      <input ref={input} type="file" accept=".docx,.doc" multiple hidden onChange={(e) => { pick(e.target.files); e.target.value = ""; }} />
      <strong>{title}</strong>
      <span>{hint}</span>
      {files.length > 0 && <em>{files.length} file{files.length > 1 ? "s" : ""}: {files.map((f) => (badge ? `${f.name} (${badge(f)})` : f.name)).join(", ")}</em>}
    </div>
  );
}

function countStats(segs) {
  const st = { tm: 0, mt: 0, error: 0, kept: 0, edited: 0 };
  segs.forEach((x) => st[x.status]++);
  return st;
}

function Seg({ seg, lang, mtpe, onCommit, onConfirm, onRetry }) {
  const [draft, setDraft] = useState(null); // null means "show the saved text"
  const v = draft ?? seg.tgt;
  const dirty = draft !== null && draft !== seg.tgt;
  const done = reviewed(seg);
  const canConfirm = mtpe && !done && seg.status !== "error";
  const onKey = (e) => {
    if (!((e.ctrlKey || e.metaKey) && e.key === "Enter")) return;
    e.preventDefault();
    if (dirty) { onCommit(draft); setDraft(null); } else if (canConfirm) onConfirm();
    let n = e.currentTarget.closest("tr").nextElementSibling;
    while (n && n.dataset.pending !== "true") n = n.nextElementSibling;
    n?.querySelector("textarea").focus();
  };
  return (
    <tr data-pending={mtpe && !done ? "true" : "false"}>
      <td>{seg.src}</td>
      <td>
        <textarea
          lang={lang} value={v} aria-label={"Translation of: " + seg.src.slice(0, 60)}
          rows={Math.max(2, v.split("\n").length, Math.ceil(v.length / 45))}
          onChange={(e) => setDraft(e.target.value)} onKeyDown={onKey}
          onBlur={() => { if (dirty) onCommit(draft); setDraft(null); }}
        />
      </td>
      <td>
        <span className={"dt-tag dt-" + seg.status}>{STATUS[seg.status]}</span>
        {seg.ok && seg.status !== "edited" && <span className="dt-tag dt-ok">Confirmed</span>}
        {canConfirm && <button className="dt-mini" onClick={onConfirm}>Confirm</button>}
        {seg.status !== "kept" && <button className="dt-mini" onClick={() => { setDraft(null); onRetry(); }}>Retranslate</button>}
      </td>
    </tr>
  );
}

// convertUrl: address of the doc→docx converter (see convert-server/). Defaults to VITE_CONVERT_URL from your .env.
export default function DocxTranslator({ convertUrl = import.meta.env?.VITE_CONVERT_URL }) {
  const [sl, setSl] = useState("en");
  const [tl, setTl] = useState("bn");
  const [localize, setLocalize] = useState(true);
  const [queue, setQueue] = useState([]);
  const [samples, setSamples] = useState([]);
  const [results, setResults] = useState([]);
  const [open, setOpen] = useState(null);
  const [filter, setFilter] = useState("all");
  const [mtpe, setMtpe] = useState(false);
  const [modePref, setModePref] = useState("auto");
  const [includeTables, setIncludeTables] = useState(false);
  const [informal, setInformal] = useState(false);
  const [corrections, setCorrections] = useState([]);
  const [kinds, setKinds] = useState({});
  const resultsRef = useRef([]);
  const update = (fn) => { const next = fn(resultsRef.current); resultsRef.current = next; setResults(next); };
  const [limit, setLimit] = useState(100);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(null);
  const [note, setNote] = useState("");
  const [memSize, setMemSize] = useState(0);
  const mem = useRef({});
  const memInput = useRef();
  const key = `docx-translator-memory:${sl}-${tl}`;

  useEffect(() => {
    try { mem.current = normalizeMem(JSON.parse(localStorage.getItem(key))); } catch { mem.current = {}; }
    setMemSize(Object.keys(mem.current).length);
  }, [key]);

  const persist = () => {
    setMemSize(Object.keys(mem.current).length);
    try { localStorage.setItem(key, JSON.stringify(mem.current)); } catch { /* storage full or blocked */ }
  };

  // Keeps .docx files as they are and converts legacy .doc files through the converter service.
  const prepare = async (list) => {
    const out = [], problems = [];
    let converted = 0;
    for (const f of list) {
      const n = f.name.toLowerCase();
      if (n.endsWith(".docx")) { out.push(f); continue; }
      if (!n.endsWith(".doc")) { problems.push(`${f.name}: only .docx and .doc files are supported.`); continue; }
      if (!convertUrl) { problems.push(`${f.name}: .doc files need the converter service. Convert it to .docx or set it up.`); continue; }
      setNote(`Converting ${f.name}…`);
      try {
        const res = await fetch(convertUrl, { method: "POST", headers: { "Content-Type": "application/octet-stream" }, body: f });
        if (!res.ok) throw new Error(await res.text());
        out.push(new File([await res.blob()], f.name.replace(/\.doc$/i, ".docx"), { type: DOCX_MIME, lastModified: f.lastModified }));
        converted++;
      } catch (e) {
        problems.push(`${f.name}: couldn't convert (${e instanceof TypeError ? "converter not reachable" : e.message}).`);
      }
    }
    setNote([converted ? `Converted ${converted} .doc file${converted > 1 ? "s" : ""} to .docx.` : "", ...problems].filter(Boolean).join(" "));
    return out;
  };

  const pickQueue = (files) => {
    setQueue(files);
    files.forEach(async (f) => {
      try { const d = await loadDocx(f); setKinds((k) => ({ ...k, [fileKey(f)]: detectMode(d) })); }
      catch { setKinds((k) => ({ ...k, [fileKey(f)]: "unreadable" })); }
    });
  };
  const modeName = (f) => {
    const m = modePref === "auto" ? kinds[fileKey(f)] : modePref;
    return m === "table" ? "two-column" : m === "paragraph" ? "paragraphs" : m === "unreadable" ? "can't read" : "detecting…";
  };

  const run = async () => {
    setBusy(true); setNote(""); update(() => []); setOpen(null);
    const notes = [];
    if (samples.length) {
      const n = await addSamples(samples, mem.current);
      notes.push(`Added ${n} segment${n === 1 ? "" : "s"} from samples.`);
      setSamples([]);
    }
    if (corrections.length) {
      const c = await learnFromPairs(corrections, mem.current, { includeTables });
      notes.push(`Learned ${c.learned} correction${c.learned === 1 ? "" : "s"} from ${c.pairs} file pair${c.pairs === 1 ? "" : "s"}.`, ...c.issues);
      setCorrections([]);
    }
    if (notes.length) setNote(notes.join(" "));
    const out = [];
    for (const f of queue) {
      setProgress({ file: f.name, done: 0, total: 0 });
      try {
        const r = await processFile(f, { sl, tl, localize, mtpe, informal, includeTables, mode: modePref }, mem.current, (done, total) => setProgress({ file: f.name, done, total }));
        out.push({ name: f.name.replace(/\.docx$/i, "") + "_translated.docx", ...r });
      } catch (e) {
        out.push({ name: f.name, error: e.message });
      }
      update(() => [...out]);
    }
    persist(); setProgress(null); setQueue([]); setBusy(false); setFilter("all"); setOpen(mtpe && out.length ? 0 : null);
  };

  // Saves an edit into the DOCX and memory. Identical untouched segments get the same text
  // (already reviewed outside MTPE mode; still awaiting confirmation inside it).
  const commit = (ri, si, text, status = "edited") => {
    const r = resultsRef.current[ri];
    const src = r.segs[si].src;
    const segs = r.segs.map((x, j) => {
      if (j !== si && !(norm(x.src) === norm(src) && ["mt", "tm", "error"].includes(x.status))) return x;
      x.write(text);
      if (j === si) return { ...x, tgt: text, status, ok: status === "edited" };
      return mtpe ? { ...x, tgt: text, status: x.status === "error" ? "mt" : x.status, ok: false } : { ...x, tgt: text, status };
    });
    update((all) => all.map((x, i) => (i === ri ? { ...r, segs } : x)));
    if (!mtpe || status === "edited") { mem.current[norm(src)] = text; persist(); }
  };

  const confirm = (ri, si) => {
    const r = resultsRef.current[ri];
    const t = r.segs[si];
    const segs = r.segs.map((x, j) => (j === si || (norm(x.src) === norm(t.src) && x.tgt === t.tgt) ? { ...x, ok: true } : x));
    update((all) => all.map((x, i) => (i === ri ? { ...r, segs } : x)));
    mem.current[norm(t.src)] = t.tgt; persist();
  };

  const retry = async (ri, si) => {
    try {
      let t = await translateText(resultsRef.current[ri].segs[si].src, sl, tl);
      if (localize && tl === "bn") t = toBn(t);
      commit(ri, si, t, "mt");
    } catch { setNote("Retranslate failed. Check your connection and try again."); }
  };

  const match = (x) => filter === "all" || (filter === "review" ? (mtpe ? !reviewed(x) : x.status === "mt" || x.status === "error") : x.status === "edited");

  const exportMem = () => download(new Blob([JSON.stringify(mem.current, null, 2)], { type: "application/json" }), "translation_memory.json");
  const importMem = async (file) => {
    try {
      const data = JSON.parse(await file.text());
      let n = 0;
      for (const [k0, v] of Object.entries(data)) { const k = norm(k0); if (k && typeof v === "string" && !Object.hasOwn(mem.current, k)) { mem.current[k] = v; n++; } }
      persist(); setNote(`Imported ${n} new segment${n === 1 ? "" : "s"}.`);
    } catch { setNote("That file isn't a valid translation memory JSON."); }
  };

  return (
    <section className="dt">
      <style>{CSS}</style>
      <h2>DOCX translator</h2>
      <p className="dt-lead">Handles two-column tables (source on the left, empty target on the right) and plain paragraph documents. Segments already in memory are reused.</p>

      <div className="dt-bar">
        <div className="dt-seg" role="group" aria-label="File type">
          {[["auto", "Auto-detect"], ["table", "Two-column table"], ["paragraph", "Plain paragraphs"]].map(([k, l]) => (
            <button key={k} className={modePref === k ? "is-on" : ""} aria-pressed={modePref === k} onClick={() => setModePref(k)} disabled={busy}>{l}</button>
          ))}
        </div>
      </div>

      <div className="dt-bar">
        <label>From
          <select value={sl} onChange={(e) => setSl(e.target.value)} disabled={busy}>
            {Object.entries(LANGS).map(([c, n]) => <option key={c} value={c}>{n}</option>)}
          </select>
        </label>
        <label>To
          <select value={tl} onChange={(e) => setTl(e.target.value)} disabled={busy}>
            {Object.entries(LANGS).map(([c, n]) => <option key={c} value={c}>{n}</option>)}
          </select>
        </label>
        {tl === "bn" && (
          <label className="dt-check"><input type="checkbox" checked={localize} onChange={(e) => setLocalize(e.target.checked)} /> Use Bengali digits (০–৯)</label>
        )}
        {modePref !== "table" && (
          <label className="dt-check"><input type="checkbox" checked={includeTables} onChange={(e) => setIncludeTables(e.target.checked)} /> Also translate text inside tables (paragraph mode)</label>
        )}
        {tl === "bn" && (
          <label className="dt-check"><input type="checkbox" checked={informal} onChange={(e) => setInformal(e.target.checked)} /> Informal address (আপনি → তুমি)</label>
        )}
        <label className="dt-check"><input type="checkbox" checked={mtpe} onChange={(e) => setMtpe(e.target.checked)} disabled={busy} /> MTPE mode: review every segment before download</label>
      </div>

      <div className="dt-grid">
        <Drop title="Files to translate" hint="Drop .docx or .doc files here, or click to choose" files={queue} onFiles={pickQueue} badge={modeName} prepare={prepare} />
        <Drop title="Sample files (optional)" hint="Finished bilingual .docx files to add to memory" files={samples} onFiles={setSamples} prepare={prepare} />
        <Drop title="Edited paragraph files (optional)" hint="Drop each edited Name_translated.docx together with its original Name.docx" files={corrections} onFiles={setCorrections} prepare={prepare} />
      </div>

      <div className="dt-bar">
        <button className="dt-primary" onClick={run} disabled={busy || (!queue.length && !samples.length && !corrections.length)}>
          {busy ? "Translating…" : queue.length ? `Translate ${queue.length} file${queue.length > 1 ? "s" : ""}` : "Update memory"}
        </button>
        <span className="dt-mem">Memory: <b>{memSize}</b> segments</span>
        <button onClick={exportMem} disabled={!memSize}>Export</button>
        <button onClick={() => memInput.current.click()}>Import</button>
        <input ref={memInput} type="file" accept=".json" hidden onChange={(e) => { e.target.files[0] && importMem(e.target.files[0]); e.target.value = ""; }} />
        <button onClick={() => { if (confirm("Clear all saved segments for this language pair?")) { mem.current = {}; persist(); } }} disabled={!memSize || busy}>Clear</button>
      </div>

      {progress && (
        <div className="dt-progress" role="status">
          <div className="dt-track"><div style={{ width: progress.total ? `${(progress.done / progress.total) * 100}%` : "100%" }} /></div>
          <span>{progress.file}: {progress.total ? `${progress.done} of ${progress.total} new segments` : "reading…"}</span>
        </div>
      )}
      {note && <p className="dt-note" role="status">{note}</p>}

      {results.map((r, i) => {
        if (r.error) return (
          <div key={i} className="dt-result"><div className="dt-head"><div><strong>{r.name}</strong><span className="dt-err">{r.error}</span></div></div></div>
        );
        const st = countStats(r.segs);
        const pending = r.segs.filter((x) => !reviewed(x)).length;
        const vis = r.segs.map((x, j) => [x, j]).filter(([x]) => match(x));
        return (
          <div key={i} className="dt-result">
            <div className="dt-head">
              <div>
                <strong>{r.name}</strong>
                <span>{r.mode === "table" ? "Two-column" : "Paragraphs"} · {st.mt} machine · {st.tm} from memory · {st.edited} edited · {st.kept} already translated{st.error ? ` · ${st.error} failed` : ""}{r.segs.length === 0 ? " · Nothing found to translate. Try the other file type." : ""}</span>
              </div>
              <div className="dt-actions">
                <button onClick={() => { setOpen(open === i ? null : i); setLimit(100); }}>{open === i ? "Close editor" : "Review and edit"}</button>
                <button className="dt-primary" disabled={mtpe && pending > 0} onClick={async () => download(await saveDocx(r.zip, r.dom), r.name)}>{mtpe && pending > 0 ? `${pending} left to review` : "Download"}</button>
              </div>
            </div>
            {open === i && (
              <>
                <div className="dt-filter" role="group" aria-label="Filter segments">
                  {[["all", "All"], ["review", "Needs review"], ["edited", "Edited"]].map(([k, l]) => (
                    <button key={k} className={filter === k ? "is-on" : ""} aria-pressed={filter === k} onClick={() => { setFilter(k); setLimit(100); }}>{l}</button>
                  ))}
                  <span>{mtpe ? `${r.segs.length - pending} of ${r.segs.length} reviewed · Ctrl+Enter confirms and moves on` : "Edits save when you click out of a box."}</span>
                </div>
                <div className="dt-scroll">
                  <table>
                    <thead><tr><th>Source</th><th>Translation</th><th>Origin</th></tr></thead>
                    <tbody>
                      {vis.slice(0, limit).map(([x, j]) => (
                        <Seg key={j} seg={x} lang={tl} mtpe={mtpe} onCommit={(t) => commit(i, j, t)} onConfirm={() => confirm(i, j)} onRetry={() => retry(i, j)} />
                      ))}
                    </tbody>
                  </table>
                  {vis.length === 0 && <p className="dt-note" style={{ padding: "0 12px" }}>No segments match this filter.</p>}
                  {vis.length > limit && <div style={{ padding: 12 }}><button onClick={() => setLimit(limit + 100)}>Show more ({vis.length - limit} left)</button></div>}
                </div>
              </>
            )}
          </div>
        );
      })}
    </section>
  );
}

const CSS = `
.dt{--ink:#16222e;--mute:#5d6b79;--line:#d5dce3;--bg:#f6f8fa;--card:#fff;--acc:#0b6b78;--acc-ink:#fff;--bad:#a8261d;
  font-family:"Hind Siliguri","Noto Sans Bengali",system-ui,sans-serif;color:var(--ink);max-width:960px;margin:0 auto;padding:24px 16px;line-height:1.5}
@media (prefers-color-scheme:dark){.dt{--ink:#e7edf2;--mute:#9aa8b5;--line:#2e3a46;--bg:#10171d;--card:#17212a;--acc:#4cc0cf;--acc-ink:#06242a;--bad:#ff8a80}}
.dt *{box-sizing:border-box}
.dt h2{margin:0 0 4px;font-size:1.6rem}
.dt-lead{margin:0 0 20px;color:var(--mute);max-width:62ch}
.dt-bar{display:flex;flex-wrap:wrap;gap:10px 14px;align-items:center;margin:14px 0}
.dt label{display:flex;gap:6px;align-items:center;font-size:.9rem}
.dt select,.dt button{font:inherit;font-size:.9rem;padding:7px 11px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--ink)}
.dt button{cursor:pointer}
.dt button:hover:not(:disabled){border-color:var(--acc)}
.dt button:disabled{opacity:.5;cursor:not-allowed}
.dt .dt-primary{background:var(--acc);color:var(--acc-ink);border-color:var(--acc);font-weight:600}
.dt :focus-visible{outline:2px solid var(--acc);outline-offset:2px}
.dt-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px}
@media (max-width:640px){.dt-grid{grid-template-columns:1fr}}
.dt-drop{display:flex;flex-direction:column;gap:4px;padding:22px 16px;border:2px dashed var(--line);border-radius:10px;background:var(--card);cursor:pointer;text-align:center}
.dt-drop span{color:var(--mute);font-size:.88rem}
.dt-drop em{font-style:normal;font-size:.85rem;word-break:break-word}
.dt-drop.is-over,.dt-drop:hover{border-color:var(--acc);background:var(--bg)}
.dt-mem{margin-left:auto;color:var(--mute);font-size:.9rem}
.dt-progress{margin:10px 0;font-size:.88rem;color:var(--mute)}
.dt-track{height:6px;background:var(--line);border-radius:3px;overflow:hidden;margin-bottom:4px}
.dt-track div{height:100%;background:var(--acc);transition:width .2s}
.dt-note{font-size:.88rem;color:var(--mute);margin:8px 0}
.dt-result{border:1px solid var(--line);border-radius:10px;background:var(--card);margin-top:14px;overflow:hidden}
.dt-head{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;align-items:center;padding:12px 16px}
.dt-head div:first-child{display:flex;flex-direction:column;min-width:0}
.dt-head span{font-size:.85rem;color:var(--mute)}
.dt-head .dt-err{color:var(--bad)}
.dt-actions{display:flex;gap:8px}
.dt-scroll{overflow-x:auto;border-top:1px solid var(--line)}
.dt-scroll table{width:100%;border-collapse:collapse;font-size:.88rem}
.dt-scroll th,.dt-scroll td{text-align:left;vertical-align:top;padding:8px 12px;border-bottom:1px solid var(--line);white-space:pre-wrap}
.dt-scroll th{background:var(--bg);font-weight:600}
.dt-tag{font-size:.78rem;padding:2px 8px;border-radius:99px;border:1px solid var(--line);white-space:nowrap}
.dt-tag.dt-tm{border-color:var(--acc);color:var(--acc)}
.dt-tag.dt-error{border-color:var(--bad);color:var(--bad)}
.dt-tag.dt-ok{border-color:var(--acc);color:var(--acc);margin-left:4px}
.dt-scroll tr[data-pending="true"] td:first-child{box-shadow:inset 3px 0 0 var(--acc)}
.dt-tag.dt-edited{background:var(--acc);border-color:var(--acc);color:var(--acc-ink)}
.dt-seg{display:inline-flex}
.dt .dt-seg button{border-radius:0;margin-left:-1px}
.dt .dt-seg button:first-child{border-radius:6px 0 0 6px;margin-left:0}
.dt .dt-seg button:last-child{border-radius:0 6px 6px 0}
.dt .dt-seg .is-on{background:var(--ink);color:var(--card);border-color:var(--ink);position:relative}
.dt-filter{display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:8px 16px;border-top:1px solid var(--line)}
.dt-filter span{margin-left:auto;font-size:.82rem;color:var(--mute)}
.dt .dt-filter .is-on{background:var(--ink);color:var(--card);border-color:var(--ink)}
.dt-scroll textarea{width:100%;min-width:240px;font:inherit;color:inherit;background:var(--bg);border:1px solid var(--line);border-radius:6px;padding:6px 8px;resize:vertical}
.dt-scroll textarea:focus{background:var(--card);border-color:var(--acc)}
.dt .dt-mini{display:block;margin-top:6px;padding:3px 8px;font-size:.78rem}
`;
