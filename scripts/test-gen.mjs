// Offline checks for the JS generators/parser (no network).
import { unzipSync, strFromU8 } from "fflate";
import { buildDocx } from "../src/docx.js";
import { buildPptx } from "../src/pptx.js";
import { parse } from "../src/parse.js";

let failures = 0;
const ok = (cond, msg) => {
  console.log((cond ? "PASS" : "FAIL") + " - " + msg);
  if (!cond) failures++;
};

const blocks = [
  { type: "heading", text: "Executive Summary", level: 1 },
  { type: "paragraph", text: "Revenue grew twenty percent this quarter." },
  { type: "paragraph", text: "Margins improved across all regions & markets <test>." },
  { type: "heading", text: "Outlook", level: 2 },
  { type: "paragraph", text: "We expect continued momentum into the next half." },
];

// ---- DOCX ----
const docx = buildDocx("Annual Report", blocks, "French", "May 30, 2026");
ok(docx[0] === 0x50 && docx[1] === 0x4b, "docx is a ZIP (PK header)");
const dfiles = unzipSync(docx);
for (const f of [
  "[Content_Types].xml",
  "_rels/.rels",
  "word/document.xml",
  "word/styles.xml",
  "word/_rels/document.xml.rels",
])
  ok(!!dfiles[f], `docx contains ${f}`);
const docXml = strFromU8(dfiles["word/document.xml"]);
ok(docXml.includes("Executive Summary"), "docx body has heading text");
ok(docXml.includes("&amp;") && docXml.includes("&lt;test&gt;"), "docx escapes special chars");

// ---- DOCX round-trip through the parser ----
const reparsed = parse(docx, "Annual Report.docx");
const texts = reparsed.blocks.map((b) => b.text);
ok(texts.includes("Executive Summary"), "parser recovers heading from generated docx");
ok(
  reparsed.blocks.find((b) => b.text === "Executive Summary")?.type === "heading",
  "parser detects heading style"
);
ok(
  texts.some((t) => t.includes("all regions & markets <test>")),
  "parser decodes escaped chars back"
);

// ---- TXT ----
const txt = parse(new TextEncoder().encode("Para one.\n\nPara two.\n\nPara three."), "notes.txt");
ok(txt.blocks.length === 3, "txt split into 3 paragraphs");

// ---- PPTX ----
const pptx = buildPptx("Annual Report", blocks, "French", "May 30, 2026");
ok(pptx[0] === 0x50 && pptx[1] === 0x4b, "pptx is a ZIP (PK header)");
const pfiles = unzipSync(pptx);
for (const f of [
  "[Content_Types].xml",
  "_rels/.rels",
  "ppt/presentation.xml",
  "ppt/_rels/presentation.xml.rels",
  "ppt/presProps.xml",
  "ppt/theme/theme1.xml",
  "ppt/slideMasters/slideMaster1.xml",
  "ppt/slideMasters/_rels/slideMaster1.xml.rels",
  "ppt/slideLayouts/slideLayout1.xml",
  "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
])
  ok(!!pfiles[f], `pptx contains ${f}`);

const slideNames = Object.keys(pfiles).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
ok(slideNames.length >= 3, `pptx has multiple slides (got ${slideNames.length})`);
// Every slide must have a matching rels file and a content-type override.
const ct = strFromU8(pfiles["[Content_Types].xml"]);
let allWired = true;
for (const n of slideNames) {
  const num = n.match(/slide(\d+)\.xml/)[1];
  if (!pfiles[`ppt/slides/_rels/slide${num}.xml.rels`]) allWired = false;
  if (!ct.includes(`/ppt/slides/slide${num}.xml`)) allWired = false;
}
ok(allWired, "every slide has rels + content-type override");
const pres = strFromU8(pfiles["ppt/presentation.xml"]);
ok((pres.match(/<p:sldId /g) || []).length === slideNames.length, "presentation lists all slides");

console.log(failures === 0 ? "\nALL GENERATOR TESTS PASSED" : `\n${failures} FAILURE(S)`);
process.exit(failures ? 1 : 0);
