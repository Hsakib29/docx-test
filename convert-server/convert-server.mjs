// Converts legacy .doc files to .docx using LibreOffice.
// Run:  npm init -y && npm i express && node convert-server.mjs
import express from "express";
import { execFile } from "node:child_process";
import { mkdtemp, writeFile, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { promisify } from "node:util";

const run = promisify(execFile);
const SOFFICE = process.env.SOFFICE_PATH || "soffice";
const ORIGIN = process.env.ALLOWED_ORIGIN || "*"; // set to your site's address in production
const PORT = process.env.PORT || 3001;

const app = express();
app.use((req, res, next) => {
  res.set({
    "Access-Control-Allow-Origin": ORIGIN,
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
  });
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.post("/convert", express.raw({ type: () => true, limit: "25mb" }), async (req, res) => {
  if (!req.body?.length) return res.status(400).send("No file received");
  const dir = await mkdtemp(join(tmpdir(), "convert-"));
  try {
    await writeFile(join(dir, "in.doc"), req.body);
    await run(
      SOFFICE,
      [
        "--headless", "--norestore",
        `-env:UserInstallation=${pathToFileURL(join(dir, "profile")).href}`, // separate profile per request, so parallel jobs don't collide
        "--convert-to", "docx", "--outdir", dir, join(dir, "in.doc"),
      ],
      { timeout: 60000 }
    );
    res
      .type("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
      .send(await readFile(join(dir, "in.docx")));
  } catch (e) {
    console.error(e.message);
    res.status(500).send("Conversion failed");
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

app.listen(PORT, () => console.log(`doc→docx converter listening on http://localhost:${PORT}/convert`));
