# DOCX translator / mini CAT tool — project spec

Browser-only tool (React + Vite, JSZip). No backend needed except an optional .doc → .docx converter.

## Current features
- Two file types, auto-detected per file (override with the mode buttons):
  - **Two-column tables**: column 1 source, empty column 2 gets the translation (port of `translate_snake`).
  - **Plain paragraphs**: body paragraphs translated in place (port of `megatrans`).
- Translation memory per language pair in localStorage; keys are whitespace-normalised; JSON import/export
  (flat `{source: target}`, compatible with the Python scripts).
- Samples (finished bilingual tables) and "edited paragraph files" (`Name.docx` + `Name_translated.docx`) feed the memory.
- Post-editor: side-by-side segment list, per-segment retranslate, filters, edits saved to the DOCX and memory.
- MTPE mode: every segment must be confirmed or edited before download; unreviewed machine output is kept out of memory.
- Options: Bengali digits, informal address (আপনি → তুমি), translate text inside tables (paragraph mode).
- `.doc` files are converted through `convert-server/` (LibreOffice) when `VITE_CONVERT_URL` is set.

## Code map (`src/DocxTranslator.jsx`)
- DOCX helpers: `loadDocx`, `detectMode`, `writeCell`, `writeParagraph`, `saveDocx`.
- Pipeline: `tableItems` / `paragraphItems` build `{src, write(text)}` items; `translateItems` is shared by both modes.
- Each segment keeps a `write` closure, so edits and retranslations rewrite the DOM and are saved by `saveDocx`.
- `commit` / `confirm` / `retry` live in the component and update memory with `norm(src)` keys.

## Known limits
- Formatting is per run or per first run; there are no inline tags yet.
- Only top-level tables and body paragraphs; no headers, footers, footnotes or text boxes.
- Machine translation uses an unofficial Google endpoint called from the browser (fine for testing; use a backend and an API key for shared use).
- Memory lives in the browser, so there is no sharing between people or devices.

## Roadmap (agreed order)
1. Inline tags and bold/italic inside segments; Bengali-aware sentence splitting (`।`); project storage in IndexedDB.
2. Fuzzy-match bands, concordance search, glossary, repeated-segment propagation.
3. QA checks (tags, numbers, terminology) and MQM-style scoring.
4. XLIFF/TMX import and export; pluggable MT engine.
5. Shared backend for teams, only when needed.

## Working agreement
- Test with real DOCX files and report failures with the file attached.
- Keep this file updated when features or decisions change.
