import { useEffect, useRef, useState } from "react";
import { LANGS, STATUS, MODES, DOCX_MIME, modeLabel, reviewed } from "../lib/constants";
import { loadDocx, detectMode, fileKey, saveDocx, norm, toBn, normalizeMem } from "../lib/docx";
import { processFile, machineTranslate, addSamples, learnFromPairs } from "../lib/engine";
import { findTerms } from "../lib/glossary";
import { loadTM, saveTM } from "../lib/store";
import { download } from "../lib/files";
import { Drop } from "./ui";
import GlossaryPanel from "./GlossaryPanel";

function countStats(segs) {
  const st = { tm: 0, mt: 0, error: 0, kept: 0, edited: 0 };
  segs.forEach((x) => st[x.status]++);
  return st;
}

function Seg({ seg, lang, mtpe, glossary, onCommit, onConfirm, onRetry }) {
  const [draft, setDraft] = useState(null); // null means "show the saved text"
  const v = draft ?? seg.tgt;
  const dirty = draft !== null && draft !== seg.tgt;
  const done = reviewed(seg);
  const canConfirm = mtpe && !done && seg.status !== "error";
  const hits = findTerms(seg.src, glossary);
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
        {hits.length > 0 && (
          <div className="dt-terms">
            {hits.map((g) => {
              const ok = norm(v).toLowerCase().includes(norm(g.tgt).toLowerCase());
              return <span key={g.id} className={"dt-term " + (ok ? "ok" : "bad")} title={g.note || ""}>{ok ? "✓" : "✗"} {g.src} → {g.tgt}{ok ? "" : " (missing)"}</span>;
            })}
          </div>
        )}
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

export default function Workspace({ project, glossary, onGlossaryChange, onProjectChange, onBack, convertUrl }) {
  const { id, sl, tl, mode, localize, informal, includeTables, mtpe, enforceTerms } = project;
  const opts = { sl, tl, localize, informal, includeTables, mtpe, enforceTerms, glossary, mode };

  const [tab, setTab] = useState("translate");
  const [queue, setQueue] = useState([]);
  const [samples, setSamples] = useState([]);
  const [corrections, setCorrections] = useState([]);
  const [kinds, setKinds] = useState({});
  const [results, setResults] = useState([]);
  const [open, setOpen] = useState(null);
  const [filter, setFilter] = useState("all");
  const [limit, setLimit] = useState(100);
  const [busy, setBusy] = useState(false);
  const [ready, setReady] = useState(false);
  const [progress, setProgress] = useState(null);
  const [note, setNote] = useState("");
  const [memSize, setMemSize] = useState(0);
  const resultsRef = useRef([]);
  const mem = useRef({});
  const loaded = useRef(false);
  const timer = useRef();
  const memInput = useRef();
  const update = (fn) => { const next = fn(resultsRef.current); resultsRef.current = next; setResults(next); };

  // The project's translation memory is loaded once and saved shortly after every change.
  useEffect(() => {
    let alive = true;
    loadTM(id).then((tm) => {
      if (!alive) return;
      mem.current = normalizeMem(tm);
      loaded.current = true;
      setMemSize(Object.keys(mem.current).length);
      setReady(true);
    });
    return () => {
      alive = false;
      clearTimeout(timer.current);
      if (loaded.current) saveTM(id, mem.current);
    };
  }, [id]);

  const persist = () => {
    setMemSize(Object.keys(mem.current).length);
    if (!loaded.current) return;
    clearTimeout(timer.current);
    timer.current = setTimeout(() => saveTM(id, mem.current), 400);
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
  const badge = (f) => modeLabel(mode === "auto" ? kinds[fileKey(f)] : mode);

  const run = async () => {
    setBusy(true); setNote(""); update(() => []); setOpen(null);
    const out = [];
    for (const f of queue) {
      setProgress({ file: f.name, done: 0, total: 0 });
      try {
        const r = await processFile(f, opts, mem.current, (done, total) => setProgress({ file: f.name, done, total }));
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
      let t = await machineTranslate(resultsRef.current[ri].segs[si].src, opts);
      if (localize && tl === "bn") t = toBn(t);
      commit(ri, si, t, "mt");
    } catch { setNote("Retranslate failed. Check your connection and try again."); }
  };

  const match = (x) => filter === "all" || (filter === "review" ? (mtpe ? !reviewed(x) : x.status === "mt" || x.status === "error") : x.status === "edited");

  const learn = async () => {
    setBusy(true);
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
    setNote(notes.join(" "));
    persist(); setBusy(false);
  };

  const exportMem = () => download(new Blob([JSON.stringify(mem.current, null, 2)], { type: "application/json" }), `${project.name}-memory.json`);
  const importMem = async (file) => {
    try {
      const data = JSON.parse(await file.text());
      let n = 0;
      for (const [k0, v] of Object.entries(data)) { const k = norm(k0); if (k && typeof v === "string" && !Object.hasOwn(mem.current, k)) { mem.current[k] = v; n++; } }
      persist(); setNote(`Imported ${n} new segment${n === 1 ? "" : "s"}.`);
    } catch { setNote("That file isn't a valid translation memory JSON."); }
  };

  const setOpt = (k, v) => onProjectChange({ ...project, [k]: v });
  const progressBar = progress && (
    <div className="dt-progress" role="status">
      <div className="dt-track"><div style={{ width: progress.total ? `${(progress.done / progress.total) * 100}%` : "100%" }} /></div>
      <span>{progress.file}: {progress.total ? `${progress.done} of ${progress.total} new segments` : "reading…"}</span>
    </div>
  );

  return (
    <div>
      <div className="dt-top">
        <div>
          <button onClick={onBack} disabled={busy}>← Projects</button>
        </div>
        <div style={{ textAlign: "right" }}>
          <strong>{project.name}</strong>
          <div className="dt-meta">{project.client ? `${project.client} · ` : ""}{LANGS[sl]} → {LANGS[tl]} · {MODES.find((m) => m.id === mode)?.label}</div>
        </div>
      </div>

      <div className="dt-tabs" role="tablist">
        {[["translate", "Translate"], ["glossary", `Glossary (${glossary.length})`], ["memory", `Memory (${memSize})`], ["settings", "Settings"]].map(([k, l]) => (
          <button key={k} role="tab" aria-selected={tab === k} className={tab === k ? "is-on" : ""} onClick={() => setTab(k)}>{l}</button>
        ))}
      </div>

      {note && <p className="dt-note" role="status">{note}</p>}

      {tab === "translate" && (
        <>
          <Drop title="Files to translate" hint="Drop .docx or .doc files here, or click to choose" files={queue} onFiles={pickQueue} badge={badge} prepare={prepare} />
          <div className="dt-bar">
            <button className="dt-primary" onClick={run} disabled={busy || !ready || !queue.length}>
              {busy ? "Translating…" : !ready ? "Loading memory…" : `Translate ${queue.length || ""} file${queue.length === 1 ? "" : "s"}`.replace("  ", " ")}
            </button>
            <span className="dt-mem">{memSize} memory segments · {glossary.length} glossary terms{enforceTerms && glossary.length ? " (applied)" : ""}</span>
          </div>
          {progressBar}

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
                    <span>{r.mode === "table" ? "Two-column" : "Paragraphs"} · {st.mt} machine · {st.tm} from memory · {st.edited} edited · {st.kept} already translated{st.error ? ` · ${st.error} failed` : ""}{r.segs.length === 0 ? " · Nothing found to translate. Try the other engine in Settings." : ""}</span>
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
                            <Seg key={j} seg={x} lang={tl} mtpe={mtpe} glossary={glossary} onCommit={(t) => commit(i, j, t)} onConfirm={() => confirm(i, j)} onRetry={() => retry(i, j)} />
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
        </>
      )}

      {tab === "glossary" && <GlossaryPanel glossary={glossary} onChange={onGlossaryChange} sl={sl} tl={tl} />}

      {tab === "memory" && (
        <>
          <p className="dt-note">This memory belongs to this project only, so client terminology never leaks between projects.</p>
          <div className="dt-bar">
            <span className="dt-mem" style={{ marginLeft: 0 }}><b>{memSize}</b> segments</span>
            <button onClick={exportMem} disabled={!memSize}>Export</button>
            <button onClick={() => memInput.current.click()} disabled={!ready}>Import</button>
            <input ref={memInput} type="file" accept=".json" hidden onChange={(e) => { e.target.files[0] && importMem(e.target.files[0]); e.target.value = ""; }} />
            <button onClick={() => { if (confirm("Clear all saved segments in this project's memory?")) { mem.current = {}; persist(); } }} disabled={!memSize || busy}>Clear</button>
          </div>
          <div className="dt-grid">
            <Drop title="Sample files (optional)" hint="Finished two-column .docx files to add to memory" files={samples} onFiles={setSamples} prepare={prepare} />
            <Drop title="Edited paragraph files (optional)" hint="Drop each edited Name_translated.docx together with its original Name.docx" files={corrections} onFiles={setCorrections} prepare={prepare} />
          </div>
          <div className="dt-bar">
            <button className="dt-primary" onClick={learn} disabled={busy || !ready || (!samples.length && !corrections.length)}>Add to memory</button>
          </div>
        </>
      )}

      {tab === "settings" && (
        <div className="dt-form">
          <label>Project name
            <input type="text" defaultValue={project.name} key={project.name} onBlur={(e) => e.target.value.trim() && setOpt("name", e.target.value.trim())} />
          </label>
          <label>Client
            <input type="text" defaultValue={project.client} key={project.client} onBlur={(e) => setOpt("client", e.target.value.trim())} />
          </label>
          <fieldset disabled={busy}>
            <legend>Translation engine</legend>
            {MODES.map((m) => (
              <label key={m.id}>
                <input type="radio" name="mode" checked={mode === m.id} onChange={() => setOpt("mode", m.id)} />
                <span>{m.label}<small>{m.hint}</small></span>
              </label>
            ))}
          </fieldset>
          {tl === "bn" && <label className="dt-check" style={{ flexDirection: "row" }}><input type="checkbox" checked={localize} onChange={(e) => setOpt("localize", e.target.checked)} /> Use Bengali digits (০–৯)</label>}
          {tl === "bn" && <label className="dt-check" style={{ flexDirection: "row" }}><input type="checkbox" checked={informal} onChange={(e) => setOpt("informal", e.target.checked)} /> Informal address (আপনি → তুমি)</label>}
          {mode !== "table" && <label className="dt-check" style={{ flexDirection: "row" }}><input type="checkbox" checked={includeTables} onChange={(e) => setOpt("includeTables", e.target.checked)} /> Also translate text inside tables (paragraph engine)</label>}
          <label className="dt-check" style={{ flexDirection: "row" }}><input type="checkbox" checked={enforceTerms} onChange={(e) => setOpt("enforceTerms", e.target.checked)} /> Apply glossary terms while translating (best effort; missing terms are flagged in the editor)</label>
          <label className="dt-check" style={{ flexDirection: "row" }}><input type="checkbox" checked={mtpe} onChange={(e) => setOpt("mtpe", e.target.checked)} /> MTPE mode: review every segment before download</label>
          <p className="dt-meta">The language pair ({LANGS[sl]} → {LANGS[tl]}) is fixed for a project so its memory and glossary stay consistent.</p>
        </div>
      )}
    </div>
  );
}
