// Generic, AI-assisted filler for an *existing* .docx template — preserves the
// template's formatting. Two kinds of targets:
//
//   1. Legacy Word form-field checkboxes  (<w:checkBox><w:default w:val="0"/>)
//      — ticked by their DOCUMENT-ORDER INDEX, never by field name. Many real
//        templates reuse field names across questions, so name-based targeting
//        would tick the wrong box; the i-th <w:checkBox> in the file is stable.
//   2. Labeled text blanks  (a paragraph/cell whose text ends in ":")
//      — a value is appended after the label run.
//
// "Document intelligence" is a Workers-AI step (decideWithWorkersAI) that reads
// the parsed input + the template's extracted schema and decides which option
// to tick per question and what to write in each blank. Extraction and editing
// are pure string/XML work, unit-testable offline without any AI.
//
// Flow:  unzip -> extractSchema(xml) -> decide() -> applyDecisions(xml) -> rezip.
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";

const DOC = "word/document.xml";

export async function fillTemplate(templateBytes, parsed, decide) {
  const files = unzipSync(new Uint8Array(templateBytes));
  const xmlBytes = files[DOC];
  if (!xmlBytes) throw new Error("Not a valid .docx (missing document.xml).");
  let xml = strFromU8(xmlBytes);

  const schema = extractSchema(xml);
  if (!schema.groups.length && !schema.fields.length)
    throw new Error("No fillable checkboxes or labeled blanks found in template.");

  const decisions = await decide(schema, serializeInput(parsed));
  xml = applyDecisions(xml, decisions);

  files[DOC] = strToU8(xml);
  return zipSync(files);
}

// --------------------------------------------------------------------------- //
// 1) Schema extraction
// --------------------------------------------------------------------------- //
// schema = {
//   groups: [{ id, question, options: [{ index, label }] }],  // index = global
//   fields: [{ id, label }],                                   // labeled blanks
// }
// `index` is the checkbox's 0-based position among all <w:checkBox> in the file,
// which is exactly what applyDecisions ticks — so targeting is collision-proof.
export function extractSchema(xml) {
  // Canonical document-order index for every checkbox.
  const ordMap = new Map(offsetsOf(xml, /<w:checkBox>/g).map((off, i) => [off, i]));

  const groups = [];
  const seen = new Set(); // ordinals already assigned to a group
  let gid = 0;

  // Every table at every nesting level. We process each table on its *own*
  // rows by blanking descendant tables (length-preserving, so offsets hold),
  // and the nested tables are handled as their own entries — so each checkbox
  // is grouped exactly once, by its innermost table.
  const tables = findTableSpans(xml);
  for (const span of tables) {
    const body = blankInnerTables(xml, span, tables);
    const rows = offsetMatches(body, /<w:tr\b[\s\S]*?<\/w:tr>/g);
    if (!rows.length) continue;
    const headerCells = cellsOf(rows[0].text).map((c) => textOf(c.text));

    for (let ri = 0; ri < rows.length; ri++) {
      const rowBase = span.start + rows[ri].index;
      const cells = cellsOf(rows[ri].text);
      const question = ri === 0 ? "" : textOf(cells[0] ? cells[0].text : "");
      const opts = [];
      for (let ci = 0; ci < cells.length; ci++) {
        const cellBase = rowBase + cells[ci].index;
        const cellText = textOf(cells[ci].text);
        for (const rel of offsetsOf(cells[ci].text, /<w:checkBox>/g)) {
          const ord = ordMap.get(cellBase + rel);
          if (ord === undefined || seen.has(ord)) continue;
          seen.add(ord);
          const label = headerCells[ci] || cellText || `option ${opts.length + 1}`;
          opts.push({ index: ord, label: clip(label, 40) });
        }
      }
      if (opts.length)
        groups.push({
          id: gid++,
          question: clip(question || rowQuestion(rows, ri), 140),
          options: opts,
        });
    }
  }

  // Loose checkboxes (not inside any table): context is their paragraph text.
  for (const p of offsetMatches(xml, /<w:p\b[\s\S]*?<\/w:p>/g)) {
    if (!/<w:checkBox>/.test(p.text)) continue;
    const ptext = textOf(p.text);
    for (const rel of offsetsOf(p.text, /<w:checkBox>/g)) {
      const ord = ordMap.get(p.index + rel);
      if (ord === undefined || seen.has(ord)) continue;
      seen.add(ord);
      groups.push({
        id: gid++,
        question: clip(ptext, 140),
        options: [{ index: ord, label: clip(ptext, 40) || "checked" }],
      });
    }
  }

  return { groups, fields: extractFields(xml) };
}

// When a row's first cell is empty (merged headers etc.), fall back to the
// nearest preceding non-empty first-cell text in the table.
function rowQuestion(rows, ri) {
  for (let i = ri; i >= 0; i--) {
    const t = textOf(cellsOf(rows[i].text)[0]?.text || "");
    if (t) return t;
  }
  return "";
}

// Labeled text blanks: a paragraph whose visible text ends in ":" (a label
// awaiting a value).
function extractFields(xml) {
  const fields = [];
  let fid = 0;
  for (const p of matchAll(xml, /<w:p\b[\s\S]*?<\/w:p>/g)) {
    if (/<w:checkBox>/.test(p)) continue;
    const t = textOf(p).trim();
    if (t.length >= 2 && t.length <= 60 && /[:：]$/.test(t)) {
      fields.push({ id: `f${fid++}`, label: t.replace(/[:：]$/, "").trim() });
    }
  }
  return fields;
}

// --------------------------------------------------------------------------- //
// 2) Apply decisions back into the XML
// --------------------------------------------------------------------------- //
// decisions = { check: number[] /* checkbox indices */, fields: { label -> value } }
export function applyDecisions(xml, decisions) {
  let out = setCheckedByIndices(
    xml,
    new Set((decisions.check || []).map(Number).filter((n) => Number.isInteger(n)))
  );
  for (const [label, value] of Object.entries(decisions.fields || {})) {
    if (value) out = appendAfterLabel(out, label, value);
  }
  return out;
}

// Tick the checkboxes whose document-order index is in `indexSet`. A single
// left-to-right pass over <w:checkBox> elements keeps the ordinal in sync with
// extractSchema's offset scan.
export function setCheckedByIndices(xml, indexSet) {
  if (!indexSet || !indexSet.size) return xml;
  let i = -1;
  return xml.replace(/<w:checkBox>([\s\S]*?)<\/w:checkBox>/g, (m, inner) => {
    i += 1;
    if (!indexSet.has(i)) return m;
    let body = inner.replace(/<w:default w:val="0"\s*\/>/, '<w:default w:val="1"/>');
    if (!/<w:default /.test(body)) body += '<w:default w:val="1"/>';
    if (!/<w:checked/.test(body)) body += '<w:checked w:val="1"/>';
    return `<w:checkBox>${body}</w:checkBox>`;
  });
}

// Append " value" as a new run in the first paragraph whose text equals the
// given label. Reuses the label's run properties so the font matches.
export function appendAfterLabel(xml, label, value) {
  const target = label.trim().toLowerCase();
  let done = false;
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (p) => {
    if (done) return p;
    const t = textOf(p).trim().replace(/[:：]$/, "").trim().toLowerCase();
    if (t !== target) return p;
    done = true;
    const rpr = (p.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [""])[0];
    const run = `<w:r>${rpr}<w:t xml:space="preserve"> ${escapeXml(value)}</w:t></w:r>`;
    return p.replace(/<\/w:p>/, run + "</w:p>");
  });
}

// --------------------------------------------------------------------------- //
// helpers
// --------------------------------------------------------------------------- //
function matchAll(s, re) {
  const out = [];
  let m;
  const r = new RegExp(re.source, re.flags.includes("g") ? re.flags : re.flags + "g");
  while ((m = r.exec(s)) !== null) out.push(m[0]);
  return out;
}
function offsetMatches(s, re) {
  const out = [];
  let m;
  const r = new RegExp(re.source, re.flags.includes("g") ? re.flags : re.flags + "g");
  while ((m = r.exec(s)) !== null) out.push({ text: m[0], index: m.index });
  return out;
}
function offsetsOf(s, re) {
  return offsetMatches(s, re).map((m) => m.index);
}
function cellsOf(rowXml) {
  return offsetMatches(rowXml, /<w:tc\b[\s\S]*?<\/w:tc>/g);
}

// Balanced spans of every <w:tbl>…</w:tbl> at every nesting depth. Only true
// table open/close tags are tokenized (not <w:tblPr>, <w:tblGrid>, …).
function findTableSpans(xml) {
  const spans = [];
  const stack = [];
  const re = /<w:tbl[\s>]|<\/w:tbl>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    if (m[0].charAt(1) === "/") {
      const start = stack.pop();
      if (start !== undefined) spans.push({ start, end: re.lastIndex });
    } else {
      stack.push(m.index);
    }
  }
  return spans;
}

// Return xml[span] with every descendant table replaced by spaces (same length,
// so character offsets are preserved). Leaves this table's own rows intact.
function blankInnerTables(xml, span, allSpans) {
  let body = xml.slice(span.start, span.end);
  for (const t of allSpans) {
    if (t.start > span.start && t.end <= span.end) {
      const a = t.start - span.start;
      const b = t.end - span.start;
      body = body.slice(0, a) + " ".repeat(b - a) + body.slice(b);
    }
  }
  return body;
}

function textOf(chunk) {
  if (!chunk) return "";
  let out = "";
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>/g;
  let m;
  while ((m = re.exec(chunk)) !== null) out += m[1] !== undefined ? decodeXml(m[1]) : " ";
  return out.replace(/\s+/g, " ").trim();
}
function clip(s, n) {
  s = (s || "").trim();
  return s.length > n ? s.slice(0, n - 1) + "…" : s;
}
function escapeXml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function decodeXml(s) {
  return s
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, "&");
}

// --------------------------------------------------------------------------- //
// Workers AI decision maker
// --------------------------------------------------------------------------- //
// Returns a `decide(schema, sourceText)` for fillTemplate(). The model is told
// to ABSTAIN (pick -1 / empty) unless the source clearly supports a choice.
const FILL_MODEL = "@cf/meta/llama-3.3-70b-instruct-fp8-fast";
const MAX_SOURCE_CHARS = 6000;
const GROUPS_PER_CALL = 12;

export function decideWithWorkersAI(env) {
  if (!env || !env.AI) throw new Error("Workers AI is not available.");
  return async (schema, sourceText) => {
    const source = (sourceText || "").slice(0, MAX_SOURCE_CHARS);
    const check = [];
    const fields = {};

    for (let i = 0; i < schema.groups.length; i += GROUPS_PER_CALL) {
      const chunk = schema.groups.slice(i, i + GROUPS_PER_CALL);
      const picks = await runJson(env, source, {
        task:
          "For each question, choose which ONE option the source material supports. " +
          'Return {"picks":[{"id":<question id>,"pick":<option index, or -1 if the ' +
          'source does not clearly answer it>}]}. Prefer -1 over guessing.',
        questions: chunk.map((g) => ({
          id: g.id,
          q: g.question,
          options: g.options.map((o) => o.label),
        })),
      });
      for (const p of (picks && picks.picks) || []) {
        const g = chunk.find((x) => x.id === p.id);
        const idx = Number(p.pick);
        if (g && Number.isInteger(idx) && idx >= 0 && idx < g.options.length)
          check.push(g.options[idx].index); // global checkbox index
      }
    }

    const labels = [...new Set(schema.fields.map((f) => f.label))];
    for (let i = 0; i < labels.length; i += GROUPS_PER_CALL * 2) {
      const chunk = labels.slice(i, i + GROUPS_PER_CALL * 2);
      const vals = await runJson(env, source, {
        task:
          "Fill each labeled blank from the source material. " +
          'Return {"values":{"<label>":"<value or empty string if unknown>"}}. ' +
          "Use empty string when the source does not contain the value.",
        labels: chunk,
      });
      for (const [k, v] of Object.entries((vals && vals.values) || {}))
        if (typeof v === "string" && v.trim()) fields[k] = v.trim();
    }

    return { check, fields };
  };
}

async function runJson(env, source, instruction) {
  const res = await env.AI.run(FILL_MODEL, {
    messages: [
      {
        role: "system",
        content:
          "You map source material onto a report form. Respond with ONLY a JSON " +
          "object, no prose, no markdown fences. Never invent facts.",
      },
      { role: "user", content: JSON.stringify(instruction) + "\n\nSOURCE MATERIAL:\n" + source },
    ],
    max_tokens: 1024,
    temperature: 0,
  });
  // Workers AI may return `response` as a string OR (for JSON-mode models) as an
  // already-parsed object — accept both.
  const raw = res && res.response;
  return raw && typeof raw === "object" ? raw : parseJson(raw);
}

function parseJson(text) {
  if (typeof text !== "string") return null;
  const s = text.trim().replace(/^```(?:json)?/i, "").replace(/```$/, "").trim();
  const a = s.indexOf("{");
  const b = s.lastIndexOf("}");
  if (a === -1 || b === -1) return null;
  try {
    return JSON.parse(s.slice(a, b + 1));
  } catch {
    return null;
  }
}

// Flatten parsed input to plain text for the AI prompt.
function serializeInput(parsed) {
  const lines = [];
  if (parsed.title) lines.push(parsed.title, "");
  for (const b of parsed.blocks || []) {
    if (b.type === "heading") lines.push("# " + b.text);
    else if (b.type === "table") for (const row of b.rows) lines.push(row.join(" | "));
    else lines.push(b.text);
  }
  return lines.join("\n");
}
