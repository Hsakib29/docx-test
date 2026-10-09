// Projects, translation memories and glossaries live in the browser's IndexedDB (one store, simple keys).
const req = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
let dbp;
const db = () => (dbp ??= new Promise((res, rej) => {
  const r = indexedDB.open("docx-cat", 1);
  r.onupgradeneeded = () => r.result.createObjectStore("kv");
  r.onsuccess = () => res(r.result);
  r.onerror = () => rej(r.error);
}));
const store = async (mode) => (await db()).transaction("kv", mode).objectStore("kv");
const kvGet = async (k) => req((await store("readonly")).get(k));
const kvSet = async (k, v) => req((await store("readwrite")).put(v, k));
const kvDel = async (k) => req((await store("readwrite")).delete(k));

export const listProjects = async () => (await kvGet("projects")) || [];
export async function saveProject(p) {
  const all = await listProjects();
  const i = all.findIndex((x) => x.id === p.id);
  if (i >= 0) all[i] = p; else all.push(p);
  await kvSet("projects", all);
  return all;
}
export async function deleteProject(id) {
  const all = (await listProjects()).filter((p) => p.id !== id);
  await kvSet("projects", all);
  await kvDel("tm:" + id);
  await kvDel("gl:" + id);
  return all;
}
export const loadTM = async (id) => (await kvGet("tm:" + id)) || {};
export const saveTM = (id, tm) => kvSet("tm:" + id, tm);
export const loadGlossary = async (id) => (await kvGet("gl:" + id)) || [];
export const saveGlossary = (id, g) => kvSet("gl:" + id, g);
