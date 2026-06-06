// LinguaDoc — free Cloudflare Worker backend.
// Serves the static frontend (via the ASSETS binding) and implements /api/*.
import { LANGUAGES } from "./languages.js";
import { parse, UnsupportedFileError } from "./parse.js";
import { translateMany } from "./translate.js";
import { buildDocx, buildDocxFromTemplate } from "./docx.js";
import { buildPptx, buildPptxFromTemplate } from "./pptx.js";
import { generateReport } from "./report.js";
import { fillTemplate, decideWithWorkersAI } from "./fill.js";
import { zipSync, strToU8 } from "fflate";

const MAX_UPLOAD_BYTES = 5 * 1024 * 1024; // 5 MB cap (free CPU/subrequest limits)

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/languages") {
      return json(LANGUAGES);
    }
    if (url.pathname === "/api/capabilities") {
      // Cloudflare build: Excel input, custom templates, AI report generation.
      return json({
        xlsx: true,
        templates: true,
        report: !!env.AI,
        fill: !!env.AI,
        maxBytes: MAX_UPLOAD_BYTES,
        build: "cloudflare",
      });
    }
    if (url.pathname === "/api/fill") {
      if (request.method !== "POST") return json({ detail: "Use POST." }, 405);
      return handleFill(request, env).catch((err) =>
        json({ detail: err.message || "Fill failed." }, err.status || 500)
      );
    }
    if (url.pathname === "/api/translate") {
      if (request.method !== "POST") return json({ detail: "Use POST." }, 405);
      return handleTranslate(request, env).catch((err) =>
        json({ detail: err.message || "Translation failed." }, err.status || 500)
      );
    }
    // Everything else: static assets (index.html, logo.svg, ...).
    return env.ASSETS.fetch(request);
  },
};

async function handleTranslate(request, env) {
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

  // Report mode (default): synthesize a structured report (summary, themed
  // section headings, bullet points, key takeaways) with Workers AI before
  // translating. Falls back to the original content if AI is off or fails.
  if (form.get("mode") !== "raw" && env && env.AI) {
    try {
      const report = await generateReport(env, parsed);
      parsed = report;
    } catch {
      /* keep the original parsed content */
    }
  }

  // Collect every text segment (title, paragraphs/headings, and table cells).
  const strings = [parsed.title];
  for (const b of parsed.blocks) {
    if (b.type === "table") {
      for (const row of b.rows) for (const cell of row) strings.push(cell);
    } else {
      strings.push(b.text);
    }
  }

  const map = await translateMany(strings, targetLang);
  const tr = (s) => map.get(s) ?? s;
  const title = tr(parsed.title);
  const blocks = parsed.blocks.map((b) =>
    b.type === "table"
      ? { ...b, rows: b.rows.map((row) => row.map(tr)) }
      : { ...b, text: tr(b.text) }
  );

  const langName = LANGUAGES[targetLang];
  const dateStr = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
  const base = parsed.title;

  // Optional custom templates. If a template build fails (templates are varied),
  // fall back to the standard design so the user still gets a report.
  const docxTpl = await templateBytes(form.get("docx_template"));
  const pptxTpl = await templateBytes(form.get("pptx_template"));

  const outputs = {};
  if (formats.includes("docx")) {
    outputs[`${base}_${targetLang}.docx`] = docxTpl
      ? tryTemplate(
          () => buildDocxFromTemplate(title, blocks, langName, dateStr, docxTpl),
          () => buildDocx(title, blocks, langName, dateStr)
        )
      : buildDocx(title, blocks, langName, dateStr);
  }
  if (formats.includes("pptx")) {
    outputs[`${base}_${targetLang}.pptx`] = pptxTpl
      ? tryTemplate(
          () => buildPptxFromTemplate(title, blocks, langName, dateStr, pptxTpl),
          () => buildPptx(title, blocks, langName, dateStr)
        )
      : buildPptx(title, blocks, langName, dateStr);
  }

  const names = Object.keys(outputs);
  if (names.length === 1) return fileResponse(names[0], outputs[names[0]]);

  // Both formats -> zip.
  const zip = {};
  for (const [name, data] of Object.entries(outputs)) zip[name] = data;
  return fileResponse(`${base}_${targetLang}.zip`, zipSync(zip));
}

// Fill an existing .docx template (with checkboxes + labeled blanks) from an
// analyzed input file, using Workers AI to decide values. Inputs (multipart):
//   file     — the source document (.docx/.xlsx/.txt) to read from
//   template — the .docx form to fill (checkboxes are toggled in place)
async function handleFill(request, env) {
  if (!env || !env.AI) throw httpError(503, "Workers AI is not enabled for this Worker.");
  const form = await request.formData();
  const file = form.get("file");
  const template = form.get("template");

  if (!file || typeof file === "string") throw httpError(400, "No input file uploaded.");
  if (!template || typeof template === "string") throw httpError(400, "No .docx template uploaded.");
  if (!/\.docx$/i.test(template.name || "")) throw httpError(400, "Template must be a .docx file.");
  if (file.size > MAX_UPLOAD_BYTES || template.size > MAX_UPLOAD_BYTES)
    throw httpError(400, "File too large for the free build.");

  let parsed;
  try {
    parsed = parse(await file.arrayBuffer(), file.name || "document");
  } catch (e) {
    if (e instanceof UnsupportedFileError) throw httpError(400, e.message);
    throw httpError(400, "Could not read the input file.");
  }
  if (!parsed.blocks.length) throw httpError(400, "No readable text found in the input file.");

  const tplBytes = await template.arrayBuffer();
  let filled;
  try {
    filled = await fillTemplate(tplBytes, parsed, decideWithWorkersAI(env));
  } catch (e) {
    throw httpError(400, e.message || "Could not fill the template.");
  }

  const base = (template.name || "report").replace(/\.docx$/i, "");
  return fileResponse(`${base}_filled.docx`, filled);
}

// Read an optional uploaded template into an ArrayBuffer, or null if absent.
async function templateBytes(field) {
  if (!field || typeof field === "string" || !field.size) return null;
  return field.arrayBuffer();
}

// Build with the template; on any failure, fall back to the standard design.
function tryTemplate(build, fallback) {
  try {
    return build();
  } catch {
    return fallback();
  }
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
