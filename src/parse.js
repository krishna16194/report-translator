// Extract a simple content model from supported uploads.
//
// Block shapes (mirrors the Python app, minus tables):
//   { type: "heading", text, level }
//   { type: "paragraph", text }
//
// Supported here: .txt and .docx. (.xlsx and tables are not supported in the
// free Cloudflare build — see the API handler for the friendly error.)
import { unzipSync, strFromU8 } from "fflate";

export class UnsupportedFileError extends Error {}

export function parse(bytes, filename) {
  const ext = (filename.split(".").pop() || "").toLowerCase();
  const title = filename.replace(/\.[^.]+$/, "");
  let blocks;
  if (ext === "txt") {
    blocks = parseTxt(bytes);
  } else if (ext === "docx") {
    blocks = parseDocx(bytes);
  } else {
    throw new UnsupportedFileError(
      `The free Cloudflare build supports .txt and .docx. ` +
        `For .xlsx or scanned files, run the full local version.`
    );
  }
  return { title, blocks };
}

function parseTxt(bytes) {
  const text = strFromU8(new Uint8Array(bytes));
  return text
    .split(/\n\s*\n/)
    .map((c) => c.trim())
    .filter(Boolean)
    .map((text) => ({ type: "paragraph", text }));
}

function parseDocx(bytes) {
  const files = unzipSync(new Uint8Array(bytes));
  const docXml = files["word/document.xml"];
  if (!docXml) {
    throw new UnsupportedFileError("Not a valid .docx (missing document.xml).");
  }
  const xml = strFromU8(docXml);
  const blocks = [];

  // Each <w:p> ... </w:p> is a paragraph (including those inside tables).
  const paraRe = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
  let m;
  while ((m = paraRe.exec(xml)) !== null) {
    const inner = m[1];
    const text = extractRunText(inner);
    if (!text.trim()) continue;

    const styleMatch = inner.match(/<w:pStyle\b[^>]*w:val="([^"]*)"/);
    const style = (styleMatch ? styleMatch[1] : "").toLowerCase();
    if (style.startsWith("heading")) {
      const digits = style.replace(/\D/g, "");
      blocks.push({ type: "heading", text, level: digits ? +digits : 1 });
    } else if (style === "title") {
      blocks.push({ type: "heading", text, level: 1 });
    } else {
      blocks.push({ type: "paragraph", text });
    }
  }
  return blocks;
}

function extractRunText(paragraphXml) {
  let out = "";
  // <w:t> text, <w:tab/> -> space, <w:br/> -> newline.
  const tokenRe = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>|<w:br\b[^>]*\/>/g;
  let m;
  while ((m = tokenRe.exec(paragraphXml)) !== null) {
    if (m[1] !== undefined) out += decodeXml(m[1]);
    else if (m[0].startsWith("<w:tab")) out += " ";
    else out += "\n";
  }
  return out;
}

function decodeXml(s) {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, "&");
}
