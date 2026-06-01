// Turn a parsed content model into a *structured report* using Workers AI.
//
// Instead of transcribing the input verbatim, we ask an LLM to synthesize it
// into a report: a short summary, themed sections with bullet points, and key
// takeaways. The result is returned in the same block model the generators
// already understand, with one new block type:
//   { type: "bullet", text }
//
// The model writes in the *source* language; the caller translates afterwards.
// Any failure (no AI binding, bad JSON, timeout) throws so the caller can fall
// back to plain translation of the original content.

const MODEL = "@cf/meta/llama-3.1-8b-instruct";
const MAX_SOURCE_CHARS = 6000; // keep the prompt within the model's context

export async function generateReport(env, parsed) {
  if (!env || !env.AI) throw new Error("Workers AI is not available.");

  const source = serialize(parsed).slice(0, MAX_SOURCE_CHARS);
  if (!source.trim()) throw new Error("Nothing to summarize.");

  const system =
    "You are a professional report writer. You read raw source material and " +
    "rewrite it as a clear, structured report. Synthesize and organize — do not " +
    "copy sentences verbatim. Write in the SAME language as the source material. " +
    "Respond with ONLY a JSON object, no prose and no markdown fences.";

  const user =
    "Produce a report from the source material below as a JSON object with these keys:\n" +
    '  "title": a short report title (string),\n' +
    '  "summary": a 2-3 sentence overview (string),\n' +
    '  "sections": an array of { "heading": string, "bullets": string[] } — ' +
    "3-6 themed sections, each with 3-6 concise, insightful bullet points,\n" +
    '  "takeaways": an array of 3-5 key-takeaway strings.\n\n' +
    "SOURCE MATERIAL:\n" +
    source;

  const res = await env.AI.run(MODEL, {
    messages: [
      { role: "system", content: system },
      { role: "user", content: user },
    ],
    max_tokens: 1536,
    temperature: 0.3,
  });

  const data = parseJson(res && res.response);
  return toBlocks(data, parsed.title);
}

// Flatten the content model to plain text the model can read.
function serialize(parsed) {
  const lines = [];
  if (parsed.title) lines.push(parsed.title, "");
  for (const b of parsed.blocks) {
    if (b.type === "heading") lines.push("# " + b.text);
    else if (b.type === "table") {
      for (const row of b.rows) lines.push(row.join(" | "));
    } else {
      lines.push(b.text);
    }
  }
  return lines.join("\n");
}

// Models sometimes wrap JSON in prose or code fences; extract the object.
function parseJson(text) {
  if (typeof text !== "string") throw new Error("AI returned no text.");
  let s = text.trim().replace(/^```(?:json)?/i, "").replace(/```$/, "").trim();
  const start = s.indexOf("{");
  const end = s.lastIndexOf("}");
  if (start === -1 || end === -1) throw new Error("AI response was not JSON.");
  return JSON.parse(s.slice(start, end + 1));
}

function toBlocks(data, fallbackTitle) {
  const blocks = [];
  const str = (v) => (typeof v === "string" ? v.trim() : "");

  if (str(data.summary)) {
    blocks.push({ type: "heading", text: "Summary", level: 1 });
    blocks.push({ type: "paragraph", text: str(data.summary) });
  }

  for (const section of Array.isArray(data.sections) ? data.sections : []) {
    if (str(section.heading)) blocks.push({ type: "heading", text: str(section.heading), level: 1 });
    for (const bullet of Array.isArray(section.bullets) ? section.bullets : []) {
      if (str(bullet)) blocks.push({ type: "bullet", text: str(bullet) });
    }
  }

  const takeaways = (Array.isArray(data.takeaways) ? data.takeaways : []).filter((t) => str(t));
  if (takeaways.length) {
    blocks.push({ type: "heading", text: "Key Takeaways", level: 1 });
    for (const t of takeaways) blocks.push({ type: "bullet", text: str(t) });
  }

  if (!blocks.length) throw new Error("AI report was empty.");
  return { title: str(data.title) || fallbackTitle, blocks };
}
