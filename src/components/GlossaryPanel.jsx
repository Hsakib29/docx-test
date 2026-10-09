import { useRef, useState } from "react";
import { parseGlossary, glossaryToCsv } from "../lib/glossary";
import { download } from "../lib/files";
import { LANGS } from "../lib/constants";

export default function GlossaryPanel({ glossary, onChange, sl, tl }) {
  const [q, setQ] = useState("");
  const [src, setSrc] = useState("");
  const [tgt, setTgt] = useState("");
  const [note, setNote] = useState("");
  const [msg, setMsg] = useState("");
  const file = useRef();

  const add = () => {
    const s = src.trim(), t = tgt.trim();
    if (!s || !t) return;
    if (glossary.some((g) => g.src.toLowerCase() === s.toLowerCase())) { setMsg(`"${s}" is already in this glossary.`); return; }
    onChange([{ id: crypto.randomUUID(), src: s, tgt: t, note: note.trim() }, ...glossary]);
    setSrc(""); setTgt(""); setNote(""); setMsg("");
  };
  const edit = (id, patch) => onChange(glossary.map((g) => (g.id === id ? { ...g, ...patch } : g)));
  const remove = (id) => onChange(glossary.filter((g) => g.id !== id));
  const importFile = async (f) => {
    const rows = parseGlossary(await f.text());
    const have = new Set(glossary.map((g) => g.src.toLowerCase()));
    const fresh = rows.filter((r) => !have.has(r.src.toLowerCase()));
    onChange([...fresh, ...glossary]);
    setMsg(`Imported ${fresh.length} new term${fresh.length === 1 ? "" : "s"}; skipped ${rows.length - fresh.length} already here.`);
  };

  const needle = q.trim().toLowerCase();
  const rows = glossary.filter((g) => !needle || [g.src, g.tgt, g.note].some((x) => (x || "").toLowerCase().includes(needle)));

  return (
    <div className="dt-gl">
      <p className="dt-note">Terms in this glossary apply to this project only. Source: {LANGS[sl]}, target: {LANGS[tl]}. Matching ignores letter case.</p>
      <div className="dt-gl-add">
        <input type="text" placeholder="Source term" value={src} onChange={(e) => setSrc(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} />
        <input type="text" placeholder="Approved translation" lang={tl} value={tgt} onChange={(e) => setTgt(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} />
        <input type="text" placeholder="Note (optional)" value={note} onChange={(e) => setNote(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} />
        <button className="dt-primary" onClick={add} disabled={!src.trim() || !tgt.trim()}>Add term</button>
      </div>
      <div className="dt-bar">
        <input type="search" placeholder="Search terms" value={q} onChange={(e) => setQ(e.target.value)} style={{ maxWidth: 260 }} />
        <button onClick={() => file.current.click()}>Import CSV</button>
        <button onClick={() => download(new Blob([glossaryToCsv(glossary)], { type: "text/csv;charset=utf-8" }), "glossary.csv")} disabled={!glossary.length}>Export CSV</button>
        <span className="dt-mem"><b>{glossary.length}</b> term{glossary.length === 1 ? "" : "s"}</span>
        <input ref={file} type="file" accept=".csv,.tsv,.txt" hidden onChange={(e) => { e.target.files[0] && importFile(e.target.files[0]); e.target.value = ""; }} />
      </div>
      {msg && <p className="dt-note" role="status">{msg}</p>}
      {rows.length > 0 && (
        <div className="dt-scroll">
          <table>
            <thead><tr><th>Source</th><th>Translation</th><th>Note</th><th /></tr></thead>
            <tbody>
              {rows.slice(0, 200).map((g) => (
                <tr key={g.id}>
                  {[["src", g.src], ["tgt", g.tgt], ["note", g.note || ""]].map(([f, v]) => (
                    <td key={f}>
                      <input
                        type="text" lang={f === "tgt" ? tl : undefined} aria-label={f} defaultValue={v} key={v}
                        onBlur={(e) => {
                          const nv = e.target.value.trim();
                          if (nv === v) return;
                          if (f !== "note" && !nv) { e.target.value = v; return; }
                          edit(g.id, { [f]: nv });
                        }}
                      />
                    </td>
                  ))}
                  <td><button className="dt-mini" onClick={() => remove(g.id)} aria-label={`Delete ${g.src}`}>Delete</button></td>
                </tr>
              ))}
            </tbody>
          </table>
          {rows.length > 200 && <p className="dt-note" style={{ padding: "0 12px" }}>Showing 200 of {rows.length}. Use search to narrow the list.</p>}
        </div>
      )}
      {glossary.length === 0 && <p className="dt-note">No terms yet. Add one above or import a CSV with the columns source, target, note.</p>}
    </div>
  );
}
