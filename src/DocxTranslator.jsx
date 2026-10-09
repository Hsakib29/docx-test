// DocxTranslator: project-based mini CAT tool. Each project has its own engine (two-column table, plain
// paragraphs or auto-detect), glossary and translation memory, all stored in the browser (IndexedDB).
// convertUrl: address of the .doc → .docx converter (see convert-server/); defaults to VITE_CONVERT_URL.
import { useEffect, useState } from "react";
import CSS from "./styles";
import ProjectList from "./components/ProjectList";
import Workspace from "./components/Workspace";
import { download } from "./lib/files";
import { normalizeMem } from "./lib/docx";
import { deleteProject, listProjects, loadGlossary, loadTM, saveGlossary, saveProject, saveTM } from "./lib/store";

const DEFAULTS = { localize: true, informal: false, includeTables: false, mtpe: false, enforceTerms: true };

export default function DocxTranslator({ convertUrl = import.meta.env?.VITE_CONVERT_URL }) {
  const [projects, setProjects] = useState(null);
  const [current, setCurrent] = useState(null);
  const [glossary, setGlossary] = useState([]);
  const [error, setError] = useState("");

  useEffect(() => {
    listProjects().then(setProjects).catch(() => { setProjects([]); setError("This browser is blocking local storage (private mode?), so projects can't be saved."); });
  }, []);

  const open = async (id) => { setGlossary(await loadGlossary(id)); setCurrent(id); };

  const create = async (f, copyLegacy) => {
    const p = { id: crypto.randomUUID(), createdAt: Date.now(), ...DEFAULTS, name: f.name, client: f.client, sl: f.sl, tl: f.tl, mode: f.mode };
    setProjects(await saveProject(p));
    if (copyLegacy) {
      try { await saveTM(p.id, normalizeMem(JSON.parse(localStorage.getItem(`docx-translator-memory:${f.sl}-${f.tl}`)))); } catch { /* nothing to copy */ }
    }
    open(p.id);
  };

  const remove = async (id) => setProjects(await deleteProject(id));

  const exportProject = async (id) => {
    const project = projects.find((p) => p.id === id);
    const bundle = { app: "docx-cat", version: 1, project, tm: await loadTM(id), glossary: await loadGlossary(id) };
    download(new Blob([JSON.stringify(bundle)], { type: "application/json" }), `${project.name}.project.json`);
  };

  const importProject = async (file) => {
    try {
      const b = JSON.parse(await file.text());
      if (b.app !== "docx-cat" || !b.project?.name) throw new Error("not a project file");
      const p = { ...DEFAULTS, ...b.project, id: crypto.randomUUID(), createdAt: Date.now() };
      setProjects(await saveProject(p));
      await saveTM(p.id, normalizeMem(b.tm));
      await saveGlossary(p.id, (b.glossary || []).map((g) => ({ ...g, id: crypto.randomUUID() })));
      setError("");
    } catch { setError("That file isn't a project backup from this tool."); }
  };

  const updateProject = async (p) => setProjects(await saveProject(p));
  const changeGlossary = (next) => { setGlossary(next); saveGlossary(current, next); };
  const project = projects?.find((p) => p.id === current);

  return (
    <section className="dt">
      <style>{CSS}</style>
      {error && <p className="dt-err" role="alert">{error}</p>}
      {projects === null ? <p className="dt-note">Loading projects…</p> : project ? (
        <Workspace key={project.id} project={project} glossary={glossary} onGlossaryChange={changeGlossary} onProjectChange={updateProject} onBack={() => setCurrent(null)} convertUrl={convertUrl} />
      ) : (
        <ProjectList projects={projects} onOpen={open} onCreate={create} onDelete={remove} onExport={exportProject} onImport={importProject} />
      )}
    </section>
  );
}
