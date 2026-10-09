import { useMemo, useRef, useState } from "react";
import { LANGS, MODES } from "../lib/constants";
import { normalizeMem } from "../lib/docx";

const legacyKey = (sl, tl) => `docx-translator-memory:${sl}-${tl}`;
const modeName = (id) => MODES.find((m) => m.id === id)?.label ?? id;

export default function ProjectList({ projects, onOpen, onCreate, onDelete, onExport, onImport }) {
  const [creating, setCreating] = useState(false);
  const [f, setF] = useState({ name: "", client: "", sl: "en", tl: "bn", mode: "auto", copyLegacy: false });
  const file = useRef();
  const set = (k, v) => setF((x) => ({ ...x, [k]: v }));

  // Memory saved by the earlier single-memory version of the tool, if any.
  const legacyCount = useMemo(() => {
    try { return Object.keys(normalizeMem(JSON.parse(localStorage.getItem(legacyKey(f.sl, f.tl))))).length; } catch { return 0; }
  }, [f.sl, f.tl]);

  const submit = (e) => {
    e.preventDefault();
    if (!f.name.trim() || f.sl === f.tl) return;
    onCreate({ ...f, name: f.name.trim(), client: f.client.trim() }, f.copyLegacy && legacyCount > 0);
    setCreating(false);
    setF({ name: "", client: "", sl: "en", tl: "bn", mode: "auto", copyLegacy: false });
  };

  return (
    <div>
      <div className="dt-top">
        <div>
          <h2>Projects</h2>
          <p className="dt-lead">Each project has its own engine, glossary and translation memory.</p>
        </div>
        <div className="dt-actions">
          <button onClick={() => file.current.click()}>Import project</button>
          <button className="dt-primary" onClick={() => setCreating(true)} disabled={creating}>New project</button>
          <input ref={file} type="file" accept=".json" hidden onChange={(e) => { e.target.files[0] && onImport(e.target.files[0]); e.target.value = ""; }} />
        </div>
      </div>

      {creating && (
        <form className="dt-form" onSubmit={submit}>
          <label>Project name
            <input type="text" value={f.name} onChange={(e) => set("name", e.target.value)} autoFocus required />
          </label>
          <label>Client (optional)
            <input type="text" value={f.client} onChange={(e) => set("client", e.target.value)} />
          </label>
          <div className="dt-bar" style={{ margin: 0 }}>
            <label>From
              <select value={f.sl} onChange={(e) => set("sl", e.target.value)}>
                {Object.entries(LANGS).map(([c, n]) => <option key={c} value={c}>{n}</option>)}
              </select>
            </label>
            <label>To
              <select value={f.tl} onChange={(e) => set("tl", e.target.value)}>
                {Object.entries(LANGS).map(([c, n]) => <option key={c} value={c}>{n}</option>)}
              </select>
            </label>
            {f.sl === f.tl && <span className="dt-err">Choose two different languages.</span>}
          </div>
          <fieldset>
            <legend>Translation engine</legend>
            {MODES.map((m) => (
              <label key={m.id}>
                <input type="radio" name="mode" checked={f.mode === m.id} onChange={() => set("mode", m.id)} />
                <span>{m.label}<small>{m.hint}</small></span>
              </label>
            ))}
          </fieldset>
          {legacyCount > 0 && (
            <label className="dt-check" style={{ flexDirection: "row" }}>
              <input type="checkbox" checked={f.copyLegacy} onChange={(e) => set("copyLegacy", e.target.checked)} />
              Start with my existing {LANGS[f.sl]} → {LANGS[f.tl]} memory ({legacyCount} segments)
            </label>
          )}
          <div className="dt-actions">
            <button type="submit" className="dt-primary" disabled={!f.name.trim() || f.sl === f.tl}>Create project</button>
            <button type="button" onClick={() => setCreating(false)}>Cancel</button>
          </div>
        </form>
      )}

      {projects.length === 0 && !creating && <p className="dt-empty">No projects yet. Create one to choose its engine and start building its glossary and memory.</p>}
      <div className="dt-cards">
        {projects.map((p) => (
          <div key={p.id} className="dt-card">
            <h3>{p.name}</h3>
            <span className="dt-meta">{p.client ? `${p.client} · ` : ""}{LANGS[p.sl]} → {LANGS[p.tl]}</span>
            <span className="dt-meta">Engine: {modeName(p.mode)}</span>
            <div className="dt-actions">
              <button className="dt-primary" onClick={() => onOpen(p.id)}>Open</button>
              <button onClick={() => onExport(p.id)}>Back up</button>
              <button onClick={() => { if (confirm(`Delete "${p.name}" with its glossary and memory? This can't be undone.`)) onDelete(p.id); }}>Delete</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
