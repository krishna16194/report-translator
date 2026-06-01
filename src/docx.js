// Build a clean, valid .docx from the content model (no external doc library).
import { zipSync, unzipSync, strToU8, strFromU8 } from "fflate";
import { esc } from "./xml.js";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="280" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="5B21B6"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2E1065"/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="2E1065"/><w:sz w:val="24"/></w:rPr></w:style>
</w:styles>`;

// One <w:p>; text may contain \n which becomes <w:br/>.
function para(text, { style, jc, bold, sz, color, italic, bullet } = {}) {
  const pPr = [];
  if (style) pPr.push(`<w:pStyle w:val="${style}"/>`);
  // A bullet item: hanging indent so wrapped lines align past the glyph.
  if (bullet) pPr.push('<w:ind w:left="360" w:hanging="220"/>', '<w:spacing w:after="60"/>');
  if (jc) pPr.push(`<w:jc w:val="${jc}"/>`);
  const pPrXml = pPr.length ? `<w:pPr>${pPr.join("")}</w:pPr>` : "";

  const rPr = [];
  if (bold) rPr.push("<w:b/>");
  if (italic) rPr.push("<w:i/>");
  if (sz) rPr.push(`<w:sz w:val="${sz}"/>`);
  if (color) rPr.push(`<w:color w:val="${color}"/>`);
  const rPrXml = rPr.length ? `<w:rPr>${rPr.join("")}</w:rPr>` : "";

  const runs = String(text)
    .split("\n")
    .map((line, i) => `${i ? "<w:br/>" : ""}<w:t xml:space="preserve">${esc(line)}</w:t>`)
    .join("");
  return `<w:p>${pPrXml}<w:r>${rPrXml}${runs}</w:r></w:p>`;
}

const PAGE_BREAK = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;

// ---- tables ----------------------------------------------------------------
const ACCENT_HEX = "5B21B6";
const LIGHT_ROW_HEX = "F2EEFB";
const TABLE_BORDERS =
  "<w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="D6CDF5"/>`)
    .join("") +
  "</w:tblBorders>";

function shade(fill) {
  return `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
}

const TABLE_WIDTH = 9026; // page text width (A4, 1440twip margins) in dxa

// A styled <w:tbl> that always fits the page: a *fixed* layout with explicit
// column widths summing to the text width, so columns never push past the
// margin and long cell text wraps instead. Header row accent-filled (white
// bold); body rows banded. Font shrinks as the column count grows.
function tableXml(rows) {
  if (!rows || !rows.length) return "";
  const cols = Math.max(...rows.map((r) => r.length));

  // Equal columns; the last one absorbs the rounding remainder so the widths
  // add up to exactly TABLE_WIDTH.
  const colW = Math.floor(TABLE_WIDTH / cols);
  const widths = Array.from({ length: cols }, (_, i) =>
    i === cols - 1 ? TABLE_WIDTH - colW * (cols - 1) : colW
  );
  const grid =
    "<w:tblGrid>" + widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("") + "</w:tblGrid>";

  // 10pt for narrow tables, easing down to 8pt for wide ones.
  const sz = cols <= 4 ? 20 : cols <= 8 ? 18 : 16;

  const trs = rows
    .map((row, ri) => {
      const header = ri === 0;
      const banded = !header && ri % 2 === 0;
      const fill = header ? ACCENT_HEX : banded ? LIGHT_ROW_HEX : null;
      const tcs = [];
      for (let ci = 0; ci < cols; ci++) {
        const text = ci < row.length ? row[ci] : "";
        const rPr = [`<w:sz w:val="${sz}"/>`];
        if (header) rPr.push("<w:b/>", '<w:color w:val="FFFFFF"/>');
        const cellP =
          `<w:p><w:r><w:rPr>${rPr.join("")}</w:rPr>` +
          `<w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
        const tcPr =
          `<w:tcPr><w:tcW w:w="${widths[ci]}" w:type="dxa"/>` +
          `${fill ? shade(fill) : ""}</w:tcPr>`;
        tcs.push(`<w:tc>${tcPr}${cellP}</w:tc>`);
      }
      return `<w:tr>${header ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${tcs.join("")}</w:tr>`;
    })
    .join("");

  return (
    `<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="${TABLE_WIDTH}"/><w:jc w:val="center"/>` +
    `${TABLE_BORDERS}<w:tblLayout w:type="fixed"/>` +
    `<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>` +
    `<w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>` +
    `</w:tblPr>${grid}${trs}</w:tbl>`
  );
}

// Render the content model (headings / paragraphs / tables) to body XML.
function contentXml(blocks) {
  const out = [];
  for (const b of blocks) {
    if (b.type === "heading") {
      const lvl = Math.min(Math.max(b.level || 1, 1), 3);
      out.push(para(b.text, { style: `Heading${lvl}` }));
    } else if (b.type === "bullet") {
      out.push(para("•  " + b.text, { bullet: true }));
    } else if (b.type === "table") {
      out.push(tableXml(b.rows));
      out.push(para("")); // spacer after the table
    } else {
      out.push(para(b.text));
    }
  }
  return out.join("");
}

export function buildDocx(title, blocks, languageName, dateStr) {
  const body = [];

  // Cover.
  for (let i = 0; i < 6; i++) body.push(para(""));
  body.push(para(title, { jc: "center", bold: true, sz: 60, color: "2E1065" }));
  body.push(para("Translated Report", { jc: "center", italic: true, sz: 28, color: "6B7280" }));
  body.push(para(`${languageName}   ·   ${dateStr}`, { jc: "center", sz: 22, color: "6B7280" }));
  body.push(PAGE_BREAK);

  body.push(contentXml(blocks));

  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${body.join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body>
</w:document>`;

  const zip = {
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(RELS),
    "word/document.xml": strToU8(document),
    "word/styles.xml": strToU8(STYLES),
    "word/_rels/document.xml.rels": strToU8(DOC_RELS),
  };
  return zipSync(zip);
}

// Reuse the caller's .docx template: keep its package (styles, theme, fonts,
// any letterhead) and append the translated content before the body's final
// sectPr. Our paragraphs reference the standard Title/Subtitle/HeadingN style
// IDs; where a template lacks one, Word falls back to default formatting.
export function buildDocxFromTemplate(title, blocks, languageName, dateStr, templateBytes) {
  const files = unzipSync(new Uint8Array(templateBytes));
  const docPart = files["word/document.xml"];
  if (!docPart) throw new Error("Not a valid .docx template (missing document.xml).");
  const xml = strFromU8(docPart);

  const bodyOpen = xml.indexOf("<w:body");
  const bodyClose = xml.lastIndexOf("</w:body>");
  if (bodyOpen === -1 || bodyClose === -1) throw new Error("Template document.xml has no <w:body>.");
  const innerStart = xml.indexOf(">", bodyOpen) + 1;
  const inner = xml.slice(innerStart, bodyClose);

  // The body-level <w:sectPr> is the last one; keep it after our content.
  const sectIdx = inner.lastIndexOf("<w:sectPr");
  let existing = inner;
  let sectPr = "";
  if (sectIdx !== -1) {
    const sectEnd = inner.indexOf("</w:sectPr>", sectIdx);
    if (sectEnd !== -1) {
      sectPr = inner.slice(sectIdx, sectEnd + "</w:sectPr>".length);
      existing = inner.slice(0, sectIdx);
    }
  }

  const head =
    para(title, { style: "Title" }) +
    para(`Translated Report — ${languageName} · ${dateStr}`, { style: "Subtitle", italic: true });
  const newInner = existing + head + contentXml(blocks) + sectPr;
  const newXml = xml.slice(0, innerStart) + newInner + xml.slice(bodyClose);

  files["word/document.xml"] = strToU8(newXml);
  return zipSync(files);
}
