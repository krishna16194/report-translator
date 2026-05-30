// LinguaDoc — free Cloudflare Worker backend.
// Serves the static frontend (via the ASSETS binding) and implements /api/*.
import { LANGUAGES } from "./languages.js";
import { parse, UnsupportedFileError } from "./parse.js";
import { translateMany } from "./translate.js";
import { buildDocx } from "./docx.js";
import { buildPptx } from "./pptx.js";
import { zipSync, strToU8 } from "fflate";

const MAX_UPLOAD_BYTES = 2 * 1024 * 1024; // 2 MB cap (free CPU/subrequest limits)

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/languages") {
      return json(LANGUAGES);
    }
    if (url.pathname === "/api/capabilities") {
      // Lean Cloudflare build: no Excel, no custom templates, 2 MB cap.
      return json({ xlsx: false, templates: false, maxBytes: MAX_UPLOAD_BYTES, build: "cloudflare" });
    }
    if (url.pathname === "/api/translate") {
      if (request.method !== "POST") return json({ detail: "Use POST." }, 405);
      return handleTranslate(request).catch((err) =>
        json({ detail: err.message || "Translation failed." }, err.status || 500)
      );
    }
    // Everything else: static assets (index.html, logo.svg, ...).
    return env.ASSETS.fetch(request);
  },
};

async function handleTranslate(request) {
  const form = await request.formData();
  const file = form.get("file");
  const targetLang = form.get("target_lang");
  const formats = String(form.get("formats") || "docx")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  if (!file || typeof file === "string") throw httpError(400, "No file uploaded.");
  if (!LANGUAGES[targetLang]) throw httpError(400, `Unsupported language '${targetLang}'.`);
  if (!formats.length || !formats.every((f) => f === "docx" || f === "pptx"))
    throw httpError(400, "formats must be 'docx', 'pptx', or both.");
  if (file.size > MAX_UPLOAD_BYTES)
    throw httpError(400, "File too large for the free build (max 2 MB).");

  const bytes = await file.arrayBuffer();
  let parsed;
  try {
    parsed = parse(bytes, file.name || "document");
  } catch (e) {
    if (e instanceof UnsupportedFileError) throw httpError(400, e.message);
    throw httpError(400, "Could not read the file.");
  }
  if (!parsed.blocks.length) throw httpError(400, "No readable text found in the file.");

  // Translate every text segment + the title.
  const strings = [parsed.title, ...parsed.blocks.map((b) => b.text)];
  const map = await translateMany(strings, targetLang);
  const tr = (s) => map.get(s) ?? s;
  const title = tr(parsed.title);
  const blocks = parsed.blocks.map((b) => ({ ...b, text: tr(b.text) }));

  const langName = LANGUAGES[targetLang];
  const dateStr = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
  const base = parsed.title;

  const outputs = {};
  if (formats.includes("docx"))
    outputs[`${base}_${targetLang}.docx`] = buildDocx(title, blocks, langName, dateStr);
  if (formats.includes("pptx"))
    outputs[`${base}_${targetLang}.pptx`] = buildPptx(title, blocks, langName, dateStr);

  const names = Object.keys(outputs);
  if (names.length === 1) return fileResponse(names[0], outputs[names[0]]);

  // Both formats -> zip.
  const zip = {};
  for (const [name, data] of Object.entries(outputs)) zip[name] = data;
  return fileResponse(`${base}_${targetLang}.zip`, zipSync(zip));
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8" },
  });
}

function fileResponse(filename, bytes) {
  return new Response(bytes, {
    headers: {
      "Content-Type": "application/octet-stream",
      "Content-Disposition": `attachment; filename="${filename}"`,
    },
  });
}

function httpError(status, message) {
  const e = new Error(message);
  e.status = status;
  return e;
}
