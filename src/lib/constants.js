export const LANGS = { en: "English", bn: "Bengali", hi: "Hindi", ur: "Urdu", ar: "Arabic", es: "Spanish", fr: "French", de: "German" };
export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
export const STATUS = { tm: "From memory", mt: "Machine", error: "Failed", kept: "Already translated", edited: "Edited" };
export const MODES = [
  { id: "auto", label: "Auto-detect each file", hint: "Looks at every file and picks the right engine." },
  { id: "table", label: "Two-column table", hint: "Source on the left, empty target on the right (the translate_snake engine)." },
  { id: "paragraph", label: "Plain paragraphs", hint: "Column-less documents translated in place (the megatrans engine)." },
];
export const modeLabel = (m) => (m === "table" ? "two-column" : m === "paragraph" ? "paragraphs" : m === "unreadable" ? "can't read" : "detecting…");
export const reviewed = (x) => x.status === "edited" || x.ok;
