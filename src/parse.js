// Extract a simple content model from supported uploads.
//
// Block shapes (mirrors the Python app):
//   { type: "heading", text, level }
//   { type: "paragraph", text }
//   { type: "table", rows: string[][] }
//
// Supported here: .txt, .docx, .xlsx and .xlsm.
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
  } else if (ext === "xlsx" || ext === "xlsm") {
    blocks = parseXlsx(bytes);
  } else {
    throw new UnsupportedFileError(
      `Unsupported file type '.${ext}'. Supported: .txt, .docx, .xlsx`
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

// ---- XLSX --------------------------------------------------------------------
// Mirrors the Python _parse_xlsx: each non-empty sheet becomes a level-2 heading
// (the sheet name) followed by a table of its used rows. We read cached cell
// values (the equivalent of openpyxl's data_only=True) — formulas aren't
// recomputed, but their stored results are used.

function parseXlsx(bytes) {
  const files = unzipSync(new Uint8Array(bytes));
  const get = (path) => (files[path] ? strFromU8(files[path]) : "");

  const shared = parseSharedStrings(get("xl/sharedStrings.xml"));
  const sheets = workbookSheets(get("xl/workbook.xml"), get("xl/_rels/workbook.xml.rels"));

  const blocks = [];
  for (const sheet of sheets) {
    const xml = get("xl/" + sheet.target);
    if (!xml) continue;
    const rows = sheetRows(xml, shared);
    if (rows.length) {
      blocks.push({ type: "heading", text: sheet.name, level: 2 });
      blocks.push({ type: "table", rows });
    }
  }
  if (!blocks.length) {
    throw new UnsupportedFileError("The workbook has no readable cells.");
  }
  return blocks;
}

// Each <si> is one shared string; rich-text strings hold several <r><t> runs,
// which we concatenate.
function parseSharedStrings(xml) {
  if (!xml) return [];
  const out = [];
  const siRe = /<(?:\w+:)?si\b[^>]*>([\s\S]*?)<\/(?:\w+:)?si>/g;
  let m;
  while ((m = siRe.exec(xml)) !== null) {
    let s = "";
    const tRe = /<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g;
    let tm;
    while ((tm = tRe.exec(m[1])) !== null) s += decodeXml(tm[1]);
    out.push(s);
  }
  return out;
}

// Returns sheets in document order: [{ name, target }] where target is relative
// to the "xl/" folder (e.g. "worksheets/sheet1.xml").
function workbookSheets(wbXml, relsXml) {
  if (!wbXml) return [];
  const rels = {};
  const relRe = /<Relationship\b[^>]*\/>/g;
  let rm;
  while ((rm = relRe.exec(relsXml)) !== null) {
    const tag = rm[0];
    const id = (tag.match(/Id="([^"]+)"/) || [])[1];
    let target = (tag.match(/Target="([^"]+)"/) || [])[1];
    if (id && target) {
      target = target.replace(/^\/?xl\//, "").replace(/^\//, "");
      rels[id] = target;
    }
  }

  const sheets = [];
  const sheetRe = /<(?:\w+:)?sheet\b([^>]*)\/>/g;
  let sm;
  while ((sm = sheetRe.exec(wbXml)) !== null) {
    const attrs = sm[1];
    const name = decodeXml((attrs.match(/name="([^"]*)"/) || [])[1] || "Sheet");
    const rid = (attrs.match(/r:id="([^"]*)"/) || [])[1];
    const target = rid ? rels[rid] : undefined;
    if (target) sheets.push({ name, target });
  }
  return sheets;
}

function sheetRows(xml, shared) {
  const rows = [];
  const rowRe = /<(?:\w+:)?row\b[^>]*>([\s\S]*?)<\/(?:\w+:)?row>/g;
  let rm;
  while ((rm = rowRe.exec(xml)) !== null) {
    const byCol = {};
    let maxCol = -1;
    const cellRe = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
    let cm;
    let autoCol = 0;
    while ((cm = cellRe.exec(rm[1])) !== null) {
      const attrs = cm[1];
      const inner = cm[2] || "";
      const ref = (attrs.match(/r="([A-Za-z]+)\d+"/) || [])[1];
      const col = ref ? colToIndex(ref) : autoCol;
      autoCol = col + 1;
      const type = (attrs.match(/t="([^"]+)"/) || [])[1] || "";
      byCol[col] = cellValue(type, inner, shared);
      if (col > maxCol) maxCol = col;
    }
    const cells = [];
    for (let i = 0; i <= maxCol; i++) cells.push(byCol[i] ?? "");
    // Skip fully empty rows.
    if (cells.some((c) => c.trim())) rows.push(cells);
  }
  return rows;
}

function cellValue(type, inner, shared) {
  if (type === "inlineStr") {
    const m = inner.match(/<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/);
    return m ? decodeXml(m[1]) : "";
  }
  const raw = (inner.match(/<(?:\w+:)?v\b[^>]*>([\s\S]*?)<\/(?:\w+:)?v>/) || [])[1] || "";
  if (type === "s") return shared[+raw] ?? "";
  if (type === "b") return raw === "1" ? "TRUE" : "FALSE";
  return decodeXml(raw);
}

// "A" -> 0, "B" -> 1, ... "Z" -> 25, "AA" -> 26.
function colToIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
