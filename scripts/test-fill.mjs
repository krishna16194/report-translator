// Offline check for src/fill.js against the real DISPACT template.
//   node scripts/test-fill.mjs
//
// Proves index-based targeting: pick the "Ja"-ish option of several questions,
// tick by global index, then confirm each ticked box's preceding question text
// matches the question we chose — i.e. no field-name collisions.
import { unzipSync, strFromU8 } from "fflate";
import { readFileSync, writeFileSync } from "node:fs";
import { extractSchema, applyDecisions, setCheckedByIndices, fillTemplate } from "../src/fill.js";

const TPL = "240927_DISPACT-2_Template_Monitoringbericht_Version 1.3.docx";
const buf = readFileSync(TPL);
const xml = strFromU8(unzipSync(new Uint8Array(buf))["word/document.xml"]);

const schema = extractSchema(xml);
console.log(`groups: ${schema.groups.length}  fields: ${schema.fields.length}`);
console.log("--- first 8 groups (option label=globalIndex) ---");
for (const g of schema.groups.slice(0, 8))
  console.log(`#${g.id} ${g.question.slice(0, 50)} :: ` +
    g.options.map((o) => `${o.label}=${o.index}`).join("  "));

// Map a checkbox ordinal -> the nearest preceding question text (ends with "?").
function precedingQuestion(x, ordinal) {
  const re = /<w:checkBox>/g;
  let i = -1, off = -1, m;
  while ((m = re.exec(x))) { if (++i === ordinal) { off = m.index; break; } }
  const before = x.slice(0, off);
  const texts = [...before.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map((mm) => mm[1]);
  // join the last ~15 text runs (a question is split across several runs)
  const ctx = texts.slice(-15).join("").replace(/\s+/g, " ").trim();
  // return text up to and including the final "?" (the question this box answers)
  const q = ctx.slice(0, ctx.lastIndexOf("?") + 1) || ctx;
  return q.slice(-70);
}

// Choose the "ja" option (else option 0) for the first 6 multi-option groups.
const chosen = schema.groups.filter((g) => g.options.length >= 2).slice(0, 6).map((g) => {
  const opt = g.options.find((o) => /^ja$/i.test(o.label)) || g.options[0];
  return { q: g.question, label: opt.label, index: opt.index };
});
const decisions = { check: chosen.map((c) => c.index), fields: {} };

console.log("\n--- targeting check (intended vs. where the tick landed) ---");
const out = applyDecisions(xml, decisions);
// rebuild a full docx with only these ticks, for visual verification
{
  const files = unzipSync(new Uint8Array(buf));
  files["word/document.xml"] = new TextEncoder().encode(out);
  const { zipSync } = await import("fflate");
  writeFileSync("_targeting.docx", zipSync(files));
}
// Informational: show each tick's intended question and the text just before
// the box that was ticked (open _targeting.docx to confirm visually).
for (const c of chosen) {
  console.log(`tick ${c.label}@${c.index}  intended="${c.q.slice(0, 36)}"`);
  console.log(`     landed after: "${precedingQuestion(out, c.index).slice(-44)}"`);
}

// End-to-end rezip with a stub AI.
const filled = await fillTemplate(buf, { title: "demo", blocks: [] }, async (s) => ({
  check: s.groups.slice(0, 3).flatMap((g) => g.options.slice(0, 1).map((o) => o.index)),
  fields: {},
}));
writeFileSync("_out_filled.docx", filled);
const outXml = strFromU8(unzipSync(filled)["word/document.xml"]);
console.log("rezip OK, checked boxes in output:", (outXml.match(/<w:checked w:val="1"\/>/g) || []).length);
