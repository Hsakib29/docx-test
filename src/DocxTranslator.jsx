// DocxTranslator.jsx — browser port of translate_snake V1
// Setup:  npm i jszip   →   import DocxTranslator from "./DocxTranslator";  <DocxTranslator />
import { useState, useRef, useEffect } from "react";
import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XMLNS = "http://www.w3.org/XML/1998/namespace";
const LANGS = { en: "English", bn: "Bengali", hi: "Hindi", ur: "Urdu", ar: "Arabic", es: "Spanish", fr: "French", de: "German" };
const BN_DIGITS = "০১২৩৪৫৬৭৮৯";
const toBn = (s) => s.replace(/[0-9]/g, (d) => BN_DIGITS[d]);
const STATUS = { tm: "From memory", mt: "Machine", error: "Failed", kept: "Already translated" };

/* ---------- DOCX helpers ---------- */
const kids = (el, name) => [...el.children].filter((c) => c.namespaceURI === W && c.localName === name);

function cellText(tc) {
  const out = [];
  for (const p of tc.getElementsByTagNameNS(W, "p")) {
    let t = "";
    for (const n of p.getElementsByTagNameNS(W, "*")) {
      if (n.localName === "t") t += n.textContent;
      else if (n.localName === "tab") t += "\t";
      else if (n.localName === "br" || n.localName === "cr") t += "\n";
    }
    out.push(t);
  }
  return out.join("\n");
}

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
  return { zip, dom, rows };
}

// Writes text into the target cell, cloning size/bold/italic/underline from the source cell's first run.
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
      if (["b", "bCs", "i", "iCs", "u", "sz", "szCs"].includes(c.localName)) rpr.appendChild(c.cloneNode(true));
    if (rpr.childNodes.length) r.appendChild(rpr);
  }
  text.split("\n").forEach((line, i) => {
    if (i) r.appendChild(dom.createElementNS(W, "w:br"));
    const t = dom.createElementNS(W, "w:t");
    t.setAttributeNS(XMLNS, "xml:space", "preserve");
    t.textContent = line;
    r.appendChild(t);
  });
  p.appendChild(r);
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

async function processFile(file, { sl, tl, localize }, mem, onTick) {
  const { zip, dom, rows } = await loadDocx(file);
  const todo = rows.filter((r) => r.src && !r.tgt);
  const fromMemory = new Set(todo.filter((r) => Object.hasOwn(mem, r.src)).map((r) => r.src));
  const need = [...new Set(todo.map((r) => r.src).filter((s) => !fromMemory.has(s)))];
  const failed = new Set();
  let next = 0, done = 0;
  const worker = async () => {
    while (next < need.length) {
      const s = need[next++];
      try { mem[s] = await translateText(s, sl, tl); } catch { failed.add(s); }
      onTick(++done, need.length);
    }
  };
  await Promise.all(Array.from({ length: Math.min(4, need.length) }, worker));

  const preview = [];
  for (const r of rows) {
    if (!r.src) continue;
    if (r.tgt) { preview.push({ src: r.src, tgt: r.tgt, status: "kept" }); continue; }
    let out = failed.has(r.src) ? "[Translation Error]" : mem[r.src];
    if (localize && tl === "bn") out = toBn(out);
    writeCell(dom, r, out);
    preview.push({ src: r.src, tgt: out, status: failed.has(r.src) ? "error" : fromMemory.has(r.src) ? "tm" : "mt" });
  }
  const stats = { tm: 0, mt: 0, error: 0, kept: 0 };
  preview.forEach((p) => stats[p.status]++);
  return { blob: await saveDocx(zip, dom), preview, stats };
}

async function addSamples(files, mem) {
  let added = 0;
  for (const f of files) {
    try {
      const { rows } = await loadDocx(f);
      for (const r of rows) if (r.src && r.tgt && !Object.hasOwn(mem, r.src)) { mem[r.src] = r.tgt; added++; }
    } catch { /* skip unreadable sample */ }
  }
  return added;
}

function download(blob, name) {
  const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: name });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

/* ---------- UI ---------- */
function Drop({ title, hint, files, onFiles }) {
  const input = useRef();
  const [over, setOver] = useState(false);
  const pick = (list) => onFiles([...list].filter((f) => f.name.toLowerCase().endsWith(".docx")));
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
      <input ref={input} type="file" accept=".docx" multiple hidden onChange={(e) => { pick(e.target.files); e.target.value = ""; }} />
      <strong>{title}</strong>
      <span>{hint}</span>
      {files.length > 0 && <em>{files.length} file{files.length > 1 ? "s" : ""}: {files.map((f) => f.name).join(", ")}</em>}
    </div>
  );
}

export default function DocxTranslator() {
  const [sl, setSl] = useState("en");
  const [tl, setTl] = useState("bn");
  const [localize, setLocalize] = useState(true);
  const [queue, setQueue] = useState([]);
  const [samples, setSamples] = useState([]);
  const [results, setResults] = useState([]);
  const [open, setOpen] = useState(null);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(null);
  const [note, setNote] = useState("");
  const [memSize, setMemSize] = useState(0);
  const mem = useRef({});
  const memInput = useRef();
  const key = `docx-translator-memory:${sl}-${tl}`;

  useEffect(() => {
    try { mem.current = JSON.parse(localStorage.getItem(key)) || {}; } catch { mem.current = {}; }
    setMemSize(Object.keys(mem.current).length);
  }, [key]);

  const persist = () => {
    setMemSize(Object.keys(mem.current).length);
    try { localStorage.setItem(key, JSON.stringify(mem.current)); } catch { /* storage full or blocked */ }
  };

  const run = async () => {
    setBusy(true); setNote(""); setResults([]); setOpen(null);
    if (samples.length) {
      const n = await addSamples(samples, mem.current);
      setNote(`Added ${n} segment${n === 1 ? "" : "s"} from samples to memory.`);
      setSamples([]);
    }
    const out = [];
    for (const f of queue) {
      setProgress({ file: f.name, done: 0, total: 0 });
      try {
        const r = await processFile(f, { sl, tl, localize }, mem.current, (done, total) => setProgress({ file: f.name, done, total }));
        out.push({ name: f.name.replace(/\.docx$/i, "") + "_translated.docx", ...r });
      } catch (e) {
        out.push({ name: f.name, error: e.message });
      }
      setResults([...out]);
    }
    persist(); setProgress(null); setQueue([]); setBusy(false);
  };

  const exportMem = () => download(new Blob([JSON.stringify(mem.current, null, 2)], { type: "application/json" }), "translation_memory.json");
  const importMem = async (file) => {
    try {
      const data = JSON.parse(await file.text());
      let n = 0;
      for (const [k, v] of Object.entries(data)) if (typeof v === "string" && !Object.hasOwn(mem.current, k)) { mem.current[k] = v; n++; }
      persist(); setNote(`Imported ${n} new segment${n === 1 ? "" : "s"}.`);
    } catch { setNote("That file isn't a valid translation memory JSON."); }
  };

  return (
    <section className="dt">
      <style>{CSS}</style>
      <h2>Bilingual DOCX translator</h2>
      <p className="dt-lead">Fills the empty right-hand column of each table row with a translation of the left-hand column. Segments already in memory are reused.</p>

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
      </div>

      <div className="dt-grid">
        <Drop title="Files to translate" hint="Drop .docx files here, or click to choose" files={queue} onFiles={setQueue} />
        <Drop title="Sample files (optional)" hint="Finished bilingual .docx files to add to memory" files={samples} onFiles={setSamples} />
      </div>

      <div className="dt-bar">
        <button className="dt-primary" onClick={run} disabled={busy || (!queue.length && !samples.length)}>
          {busy ? "Translating…" : queue.length ? `Translate ${queue.length} file${queue.length > 1 ? "s" : ""}` : "Add samples to memory"}
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

      {results.map((r, i) => (
        <div key={i} className="dt-result">
          <div className="dt-head">
            <div>
              <strong>{r.name}</strong>
              {r.error ? <span className="dt-err">{r.error}</span> : (
                <span>{r.stats.mt} machine · {r.stats.tm} from memory · {r.stats.kept} already translated{r.stats.error ? ` · ${r.stats.error} failed` : ""}</span>
              )}
            </div>
            {!r.error && (
              <div className="dt-actions">
                <button onClick={() => setOpen(open === i ? null : i)}>{open === i ? "Hide preview" : "Preview"}</button>
                <button className="dt-primary" onClick={() => download(r.blob, r.name)}>Download</button>
              </div>
            )}
          </div>
          {open === i && (
            <div className="dt-scroll">
              <table>
                <thead><tr><th>Source</th><th>Translation</th><th>Origin</th></tr></thead>
                <tbody>
                  {r.preview.slice(0, 300).map((p, j) => (
                    <tr key={j}><td>{p.src}</td><td lang={tl}>{p.tgt}</td><td><span className={"dt-tag dt-" + p.status}>{STATUS[p.status]}</span></td></tr>
                  ))}
                </tbody>
              </table>
              {r.preview.length > 300 && <p className="dt-note">Showing the first 300 of {r.preview.length} rows. The download contains all of them.</p>}
            </div>
          )}
        </div>
      ))}
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
.dt-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
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
`;
