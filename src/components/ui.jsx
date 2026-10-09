import { useRef, useState } from "react";

export function Drop({ title, hint, files, onFiles, badge, prepare }) {
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
