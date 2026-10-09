# DOCX translator / mini CAT tool — project spec

Browser-only tool (React + Vite, JSZip). Projects, memories and glossaries are stored in IndexedDB.
The only optional server piece is the .doc → .docx converter in `convert-server/`.

## How it works
- **Projects**: each project has a name, optional client, a fixed language pair, an engine and options.
  Each project has its **own translation memory and its own glossary**, so client terminology never mixes.
- **Engine** (chosen when the project is created, editable in Settings):
  - Two-column table (the `translate_snake` engine): column 1 source, empty column 2 gets the translation.
  - Plain paragraphs (the `megatrans` engine): body paragraphs translated in place.
  - Auto-detect: picks one of the two per file.
- **Glossary**: source term → approved translation (+ note), CSV import/export. While translating, terms are swapped for
  `{0}`, `{1}`… placeholders and replaced by the approved target afterwards; if the engine loses a placeholder the segment
  falls back to a plain translation. The editor flags every segment whose source has a glossary term that is missing from the target.
- **Memory**: whitespace-normalised keys, JSON import/export (compatible with the Python scripts), samples and edited
  paragraph files can be fed in from the Memory tab.
- **Editor**: side-by-side segments, retranslate, filters, term chips; edits are written into the DOCX and the memory.
- **MTPE mode**: every segment must be confirmed or edited before download; unreviewed machine output stays out of memory.
- Project backup/restore as one JSON file (project + memory + glossary).
- Options per project: Bengali digits, informal address (আপনি → তুমি), translate text inside tables, enforce glossary, MTPE.

## Code map
- `src/DocxTranslator.jsx` — top level: project list or workspace, project backup/import.
- `src/components/` — `ProjectList`, `Workspace` (tabs: Translate, Glossary, Memory, Settings), `GlossaryPanel`, `ui` (drop zone).
- `src/lib/docx.js` — DOCX reading/writing, mode detection, item builders (`tableItems`, `paragraphItems`).
- `src/lib/engine.js` — machine translation, `translateItems`, `processFile`, memory learning.
- `src/lib/glossary.js` — term matching, placeholder protect/restore, CSV.
- `src/lib/store.js` — IndexedDB (projects index, `tm:<id>`, `gl:<id>`).
- `src/lib/constants.js`, `src/lib/files.js`, `src/styles.js`.

## Known limits
- Formatting is per run (paragraphs) or from the first run (table cells); there are no inline tags yet.
- Only body paragraphs and top-level tables; no headers, footers, footnotes or text boxes.
- Machine translation uses an unofficial Google endpoint called from the browser (fine for testing; use a backend and an API key for shared use).
- Glossary enforcement relies on the engine keeping `{n}` placeholders. Not verified against the live engine for Bengali yet.
- Data lives in one browser. Use "Back up" on a project to move or keep it safe; browsers can clear site data.
- No sharing between people or devices.

## Roadmap
1. Inline tags and bold/italic inside segments; Bengali-aware sentence splitting (`।`).
2. Fuzzy-match bands, concordance search, repeated-segment propagation across a project.
3. QA checks (tags, numbers) and MQM-style scoring.
4. XLIFF/TMX import and export; pluggable MT engine (official API via a backend).
5. Shared backend for teams, only when needed.

## Working agreement
- Test with real DOCX files and report failures with the file attached.
- Keep this file updated when features or decisions change.
