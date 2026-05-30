// Build a minimal but valid .pptx from the content model (no external library).
// One title slide + section-aware content slides; text frames use normAutofit
// so PowerPoint shrinks overflowing text to fit.
import { zipSync, strToU8 } from "fflate";
import { esc } from "./xml.js";

const W = 12192000; // 16:9 slide, EMU
const H = 6858000;
const IN = 914400;

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

// ---- packing: split paragraphs into slide-sized bullet groups -------------
function packBullets(bullets) {
  const groups = [];
  let cur = [];
  let used = 0;
  for (const t of bullets) {
    const lines = Math.max(1, Math.ceil(t.length / 90));
    if (cur.length && (used + lines > 13 || cur.length >= 7)) {
      groups.push(cur);
      cur = [];
      used = 0;
    }
    cur.push(t);
    used += lines;
  }
  if (cur.length) groups.push(cur);
  return groups;
}

function groupSlides(title, blocks) {
  const slides = [];
  let curTitle = title;
  let bullets = [];
  let headingRendered = true; // the title slide covers the document title
  const flush = () => {
    if (!bullets.length) return;
    packBullets(bullets).forEach((g, i) =>
      slides.push({ title: i ? `${curTitle} (cont.)` : curTitle, bullets: g })
    );
    bullets = [];
    headingRendered = true;
  };
  for (const b of blocks) {
    if (b.type === "heading") {
      flush();
      if (!headingRendered) slides.push({ title: curTitle, bullets: [] });
      curTitle = b.text;
      headingRendered = false;
    } else {
      bullets.push(b.text);
    }
  }
  flush();
  if (!headingRendered) slides.push({ title: curTitle, bullets: [] });
  return slides;
}

// ---- shape / text helpers --------------------------------------------------
function runs(text, sz, b, color) {
  const rPr = `<a:rPr lang="en-US" sz="${sz}"${b ? ' b="1"' : ""} dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr>`;
  return String(text)
    .split("\n")
    .map((line, i) => `${i ? "<a:br/>" : ""}<a:r>${rPr}<a:t>${esc(line)}</a:t></a:r>`)
    .join("");
}

function aP(text, { sz = 1800, b = false, color = "1F2937", align, spc = 600 } = {}) {
  const algn = align ? ` algn="${align}"` : "";
  const spcAft = spc ? `<a:spcAft><a:spcPts val="${spc}"/></a:spcAft>` : "";
  return `<a:p><a:pPr${algn}>${spcAft}</a:pPr>${runs(text, sz, b, color)}</a:p>`;
}

function textBox(id, name, x, y, cx, cy, paragraphs) {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/>` +
    `<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:normAutofit/></a:bodyPr>` +
    `<a:lstStyle/>${paragraphs}</p:txBody></p:sp>`
  );
}

function bar(id, x, y, cx, cy, color) {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`
  );
}

function slideDoc(shapes) {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<p:sld ${NS}><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    `${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
  );
}

function titleSlide(title, languageName, dateStr) {
  const t = textBox(2, "Title", IN, Math.round(IN * 2.2), W - 2 * IN, Math.round(IN * 1.7), [
    aP(title, { sz: 4000, b: true, color: "2E1065", align: "ctr", spc: 0 }),
  ].join(""));
  const sub = textBox(3, "Subtitle", IN, Math.round(IN * 3.9), W - 2 * IN, IN, [
    aP("Translated Report", { sz: 2000, color: "5B21B6", align: "ctr", spc: 0 }),
  ].join(""));
  const meta = textBox(4, "Meta", IN, Math.round(IN * 5.4), W - 2 * IN, IN, [
    aP(`${languageName}   ·   ${dateStr}`, { sz: 1600, color: "6B7280", align: "ctr", spc: 0 }),
  ].join(""));
  return slideDoc(bar(5, 0, 0, IN / 2, H, "5B21B6") + t + sub + meta);
}

function contentSlide(title, bullets) {
  const top = bar(2, 0, 0, W, 120000, "5B21B6");
  const head = textBox(3, "Heading", Math.round(IN * 0.6), 320000, W - Math.round(IN * 1.2), 900000, [
    aP(title, { sz: 2800, b: true, color: "2E1065", spc: 0 }),
  ].join(""));
  const body = textBox(
    4,
    "Body",
    Math.round(IN * 0.7),
    1500000,
    W - Math.round(IN * 1.4),
    H - 1900000,
    bullets.map((t) => aP("•  " + t, { sz: 1800, color: "1F2937" })).join("") || "<a:p/>"
  );
  return slideDoc(top + head + body);
}

// ---- static package parts --------------------------------------------------
const THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="5B21B6"/></a:accent1><a:accent2><a:srgbClr val="8B5CF6"/></a:accent2><a:accent3><a:srgbClr val="A5B592"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

const SLIDE_MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr/></a:defPPr></p:otherStyle></p:txStyles></p:sldMaster>`;

const SLIDE_LAYOUT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const PRES_PROPS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentationPr ${NS}/>`;

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export function buildPptx(title, blocks, languageName, dateStr) {
  const slides = [titleSlide(title, languageName, dateStr)];
  for (const s of groupSlides(title, blocks)) slides.push(contentSlide(s.title, s.bullets));

  const zip = {};
  // Slides + their rels.
  const slideOverrides = [];
  const presRels = [`<Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`];
  const sldIds = [];
  slides.forEach((xml, i) => {
    const n = i + 1;
    const rid = `rId${n + 1}`; // rId2.. for slides
    zip[`ppt/slides/slide${n}.xml`] = strToU8(xml);
    zip[`ppt/slides/_rels/slide${n}.xml.rels`] = strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`
    );
    slideOverrides.push(
      `<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    );
    presRels.push(`<Relationship Id="${rid}" Type="${REL}/slide" Target="slides/slide${n}.xml"/>`);
    sldIds.push(`<p:sldId id="${255 + n}" r:id="${rid}"/>`);
  });

  const themeRid = `rId${slides.length + 2}`;
  const propsRid = `rId${slides.length + 3}`;
  presRels.push(`<Relationship Id="${themeRid}" Type="${REL}/theme" Target="theme/theme1.xml"/>`);
  presRels.push(`<Relationship Id="${propsRid}" Type="${REL}/presProps" Target="presProps.xml"/>`);

  const presentation = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${sldIds.join("")}</p:sldIdLst><p:sldSz cx="${W}" cy="${H}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${slideOverrides.join("")}</Types>`;

  zip["[Content_Types].xml"] = strToU8(contentTypes);
  zip["_rels/.rels"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`
  );
  zip["ppt/presentation.xml"] = strToU8(presentation);
  zip["ppt/_rels/presentation.xml.rels"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${presRels.join("")}</Relationships>`
  );
  zip["ppt/presProps.xml"] = strToU8(PRES_PROPS);
  zip["ppt/theme/theme1.xml"] = strToU8(THEME);
  zip["ppt/slideMasters/slideMaster1.xml"] = strToU8(SLIDE_MASTER);
  zip["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL}/theme" Target="../theme/theme1.xml"/></Relationships>`
  );
  zip["ppt/slideLayouts/slideLayout1.xml"] = strToU8(SLIDE_LAYOUT);
  zip["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`
  );

  return zipSync(zip);
}
