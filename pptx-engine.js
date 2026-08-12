// pptx-engine.js
// Edits the real Offering_Letter template directly using JSZip (runs in browser or Node).
// This ports the proven Python technique 1:1.

const SHAPE_MAP_SLIDE2 = {
  1178: 'nama_client',
  1183: 'phone',
  1185: 'tanggal',
  1187: 'pax',
  1189: 'activity',
};

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function rupiah(n) {
  return 'Rp. ' + Math.round(n).toLocaleString('id-ID');
}

// Replace the text inside a <a:t> run that sits inside a <p:sp> with a given cNvPr id.
// This mirrors python-pptx's `run.text = value` (only touches the FIRST run's text in that shape).
function replaceShapeFirstRunText(xml, shapeId, newText) {
  const spRegex = new RegExp(
    `(<p:sp>(?:(?!<\\/p:sp>)[\\s\\S])*?<p:cNvPr id="${shapeId}"[\\s\\S]*?<\\/p:sp>)`
  );
  const match = xml.match(spRegex);
  if (!match) return xml;
  let spXml = match[1];
  const tRegex = /(<a:t>)([\s\S]*?)(<\/a:t>)/;
  const newSpXml = spXml.replace(tRegex, (m, open, _old, close) => open + escapeXml(newText) + close);
  return xml.replace(spXml, newSpXml);
}

// Slide 4 shapes don't have stable ids we can rely on the same way; we match by the
// run's existing text prefix instead (same logic as the Python version).
function replaceByTextPrefix(xml, prefix, newFullText) {
  // Find the first <a:t>PREFIX...</a:t> occurrence and replace the whole run text.
  const escapedPrefix = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(<a:t>)(${escapedPrefix}[\\s\\S]*?)(<\\/a:t>)`);
  return xml.replace(re, (m, open, _old, close) => open + escapeXml(newFullText) + close);
}

async function editSlide2(zip, data) {
  const path = 'ppt/slides/slide2.xml';
  let xml = await zip.file(path).async('string');
  xml = replaceShapeFirstRunText(xml, 1178, data.nama_client);
  xml = replaceShapeFirstRunText(xml, 1183, data.phone);
  xml = replaceShapeFirstRunText(xml, 1185, data.tanggal);
  xml = replaceShapeFirstRunText(xml, 1187, data.pax + ' pax');
  xml = replaceShapeFirstRunText(xml, 1189, data.activity);
  zip.file(path, xml);
}

async function editSlide4(zip, data) {
  const path = 'ppt/slides/slide4.xml';
  let xml = await zip.file(path).async('string');
  xml = xml.replace(/(<a:t>)\{\{judul_program\}\}(<\/a:t>)/, (m, o, c) => o + escapeXml(data.judul_program) + c);
  xml = replaceByTextPrefix(xml, 'Harga / Pax', `Harga / Pax\t: ${rupiah(data.harga_pax)} / Pax / Orang`);
  xml = replaceByTextPrefix(xml, 'Total Harga', `Total Harga \t: ${rupiah(data.harga_pax * data.pax)}`);
  if (data.taxNote) {
    xml = replaceShapeFirstRunText(xml, 1227, data.taxNote);
  }
  zip.file(path, xml);
}

// --- buildPptx assembled after all edit functions are defined (see bottom of file) ---

// --- Generic image swap: points a slide's picture relationship to a new image ---
async function swapImage(zip, slideNum, oldRId, newImageBytes, ext) {
  const slidePath = `ppt/slides/slide${slideNum}.xml`;
  const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
  let slideXml = await zip.file(slidePath).async('string');
  let relsXml = await zip.file(relsPath).async('string');

  const newRId = 'rIdSwap' + Math.random().toString(36).slice(2, 8);
  const mediaName = `image_swap_${slideNum}_${Date.now()}_${Math.floor(Math.random()*1e6)}.${ext}`;
  zip.file(`ppt/media/${mediaName}`, newImageBytes);

  const newRel = `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`;
  relsXml = relsXml.replace('</Relationships>', newRel + '</Relationships>');
  slideXml = slideXml.replace(`r:embed="${oldRId}"`, `r:embed="${newRId}"`);

  zip.file(slidePath, slideXml);
  zip.file(relsPath, relsXml);
}

// --- Slide 5: dynamic list rebuild (Include / Meals / Fasilitas) ---
// Supports ANY number of items (not limited to the original template's slot count),
// and auto-fits the font size to the item count so the box height (2699400 EMU) is well used.
const SLIDE5_LIST_HEIGHT = 2699400;
const SLIDE5_LIST_WIDTH = 1666200;

function buildListRunStyle(sz) {
  return `sz="${sz}" b="1" dirty="0"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill><a:latin typeface="Calibri"/><a:ea typeface="Calibri"/><a:cs typeface="Calibri"/><a:sym typeface="Calibri"/>`;
}

function rebuildListParagraph(spXml, items) {
  const paragraphs = spXml.split('</a:p>');
  if (!paragraphs[1]) return spXml;

  const pPrMatch = paragraphs[1].match(/<a:pPr[\s\S]*?<\/a:pPr>/);
  const pPr = pPrMatch ? pPrMatch[0] : '';

  const sz = computeAutoFontSizeForTexts(items, 1, SLIDE5_LIST_HEIGHT, SLIDE5_LIST_WIDTH, 7, 11, 0);
  const style = buildListRunStyle(sz);

  const runs = items.map(text => `<a:r><a:rPr lang="en" ${style}</a:rPr><a:t>${escapeXml(text)}</a:t></a:r>`);
  const br = `<a:br><a:rPr lang="en" ${style}</a:rPr></a:br>`;
  const body = runs.join(br);
  const endParaRPr = `<a:endParaRPr ${style}</a:endParaRPr>`;

  paragraphs[1] = `<a:p>${pPr}${body}${endParaRPr}`;
  return paragraphs.join('</a:p>');
}

async function editSlide5List(zip, shapeId, items) {
  const path = 'ppt/slides/slide5.xml';
  let xml = await zip.file(path).async('string');
  const spRegex = new RegExp(`(<p:sp>(?:(?!<\\/p:sp>)[\\s\\S])*?<p:cNvPr id="${shapeId}"[\\s\\S]*?<\\/p:sp>)`);
  const match = xml.match(spRegex);
  if (!match) { zip.file(path, xml); return; }
  const spXml = match[1];
  const newSpXml = rebuildListParagraph(spXml, items.filter(Boolean));
  xml = xml.replace(spXml, newSpXml);
  zip.file(path, xml);
}

async function editSlide5FacilityList(zip, items) { return editSlide5List(zip, 1240, items); }
async function editSlide5Meals(zip, items) { return editSlide5List(zip, 1241, items); }
async function editSlide5Fasilitas(zip, items) { return editSlide5List(zip, 1242, items); }

async function editSlide5Photos(zip, photos) {
  // photos: array of up to 5 {bytes, ext} — first 4 fill the corner frames,
  // the 5th (if present) fills the centered/front photo.
  const rIds = ['rId3', 'rId4', 'rId5', 'rId6', 'rId7'];
  for (let i = 0; i < Math.min(photos.length, 5); i++) {
    await swapImage(zip, 5, rIds[i], photos[i].bytes, photos[i].ext);
  }
}

const RUNDOWN_COLOR = 'AB8645';
const TIME_COL_W = 1300000;
const INDENT_STEP = 180000;
const SLIDE6_AVAILABLE_HEIGHT = 3200000;   // reduced with safety margin for Google Drive's renderer
const SLIDE7_LEFT_AVAILABLE_HEIGHT = 3500000;
const SLIDE7_RIGHT_AVAILABLE_HEIGHT = 3600000;

// --- Auto-fit v2: wrap-aware. Estimates how many visual lines each text will wrap
// into at a given font size + column width, then iterates toward a font size that
// fills the available height without overflowing. ---
function estimateWrappedLines(text, fontSizePt, colWidthEMU) {
  const colWidthPt = colWidthEMU / 12700;
  const avgCharWidthPt = fontSizePt * 0.62; // conservative estimate — better to overestimate wrap than overflow
  const charsPerLine = Math.max(Math.floor(colWidthPt / avgCharWidthPt), 6);
  return Math.max(1, Math.ceil((text.length || 1) / charsPerLine));
}

function computeAutoFontSizeForTexts(texts, rowCount, availableHeightEMU, colWidthEMU, minPt, maxPt, rowOverheadPt) {
  let size = maxPt;
  for (let iter = 0; iter < 8; iter++) {
    const totalWrapped = texts.reduce((sum, t) => sum + estimateWrappedLines(t, size, colWidthEMU), 0);
    const heightPt = availableHeightEMU / 12700;
    const usablePt = Math.max(heightPt - rowCount * (rowOverheadPt || 0), 20);
    let newSize = usablePt / (Math.max(totalWrapped, 1) * 1.85);
    newSize = Math.max(minPt, Math.min(maxPt, newSize));
    if (Math.abs(newSize - size) < 0.25) { size = newSize; break; }
    size = newSize;
  }
  return Math.round(size * 100);
}

function parseRundownLine(raw) {
  let level = 0, line = raw;
  while (line.startsWith('>')) { level++; line = line.slice(1); }
  line = line.trim();
  let italic = false;
  if (line.length > 1 && line.startsWith('*') && line.endsWith('*')) {
    italic = true;
    line = line.slice(1, -1);
  }
  return { level, italic, text: line };
}

function buildTcPr(extraMarL) {
  return `<a:tcPr marL="${extraMarL || 0}" marR="0" marT="20000" marB="60000"><a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT><a:lnB w="0"><a:noFill/></a:lnB><a:noFill/></a:tcPr>`;
}

function buildRun(text, sz, italic) {
  return `<a:r><a:rPr sz="${sz}" b="1" i="${italic ? 1 : 0}"><a:solidFill><a:srgbClr val="${RUNDOWN_COLOR}"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>${escapeXml(text)}</a:t></a:r>`;
}

function buildDescCell(lines, sz, nested) {
  let paras = '';
  lines.forEach((raw, i) => {
    const { level, italic, text } = parseRundownLine(raw);
    const marL = nested ? INDENT_STEP * level : 0;
    const spcBef = i === 0 ? 0 : 300;
    const pPr = `<a:pPr${marL ? ` marL="${marL}" indent="0"` : ''}><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="${spcBef}"/></a:spcBef></a:pPr>`;
    paras += `<a:p>${pPr}${buildRun(text, sz, italic)}</a:p>`;
  });
  return `<a:tc><a:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${paras}</a:txBody>${buildTcPr()}</a:tc>`;
}

function buildTimeCell(timeText, sz, extraMarL) {
  return `<a:tc><a:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p>${buildRun(timeText, sz, false)}</a:p></a:txBody>${buildTcPr(extraMarL)}</a:tc>`;
}

function buildSeparatorLine(shapeId, x, y, width) {
  const lineWidth = Math.round(width * 0.88);
  const offsetX = x + Math.round((width - lineWidth) / 2);
  return `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="Sep ${shapeId}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${offsetX}" y="${y}"/><a:ext cx="${lineWidth}" cy="3175"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="E4D9BF"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
}

// availableHeightEMU: fixed box height to fill. sizeBounds: {min,max} in points.
function buildRundownTable(shapeId, x, y, width, rows, availableHeightEMU, sizeBounds, nested, timeMarL) {
  const descColWidth = width - TIME_COL_W;
  const allTexts = [];
  rows.forEach(row => {
    row.label.split('\n').map(l => l.trim()).filter(Boolean).forEach(raw => {
      allTexts.push(parseRundownLine(raw).text);
    });
  });
  const sz = computeAutoFontSizeForTexts(allTexts, rows.length, availableHeightEMU, descColWidth, sizeBounds.min, sizeBounds.max, 7);

  const lineHeightPt = (sz / 100) * 1.6;
  const rowOverheadPt = 8;
  let totalHeightEMU = 0;
  const rowHeights = rows.map(row => {
    const lines = row.label.split('\n').map(l => l.trim()).filter(Boolean);
    const visualLineCount = lines.reduce((sum, raw) => {
      const { text } = parseRundownLine(raw);
      return sum + estimateWrappedLines(text, sz / 100, descColWidth);
    }, 0);
    return Math.round((visualLineCount * lineHeightPt + rowOverheadPt) * 12700);
  });

  const trs = rows.map((row, idx) => {
    const timeText = row.time2 ? `${row.time1}-${row.time2}` : row.time1;
    const lines = row.label.split('\n').map(l => l.trim()).filter(Boolean);
    totalHeightEMU += rowHeights[idx];
    return `<a:tr h="${rowHeights[idx]}">${buildTimeCell(timeText, sz, timeMarL)}${buildDescCell(lines, sz, nested)}</a:tr>`;
  }).join('');

  const tableXml = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${shapeId}" name="Table ${shapeId}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${totalHeightEMU}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="0" bandRow="0"/><a:tblGrid><a:gridCol w="${TIME_COL_W}"/><a:gridCol w="${width - TIME_COL_W}"/></a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;

  // separator lines as independent shapes, placed at each row boundary (except after the last row)
  let cumulative = 0;
  let linesXml = '';
  rowHeights.forEach((h, idx) => {
    cumulative += h;
    if (idx < rowHeights.length - 1) {
      linesXml += buildSeparatorLine(shapeId * 100 + idx, x, y + cumulative, width);
    }
  });

  return tableXml + linesXml;
}

function removeShapeByPos(xml, offX, offY) {
  const re = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
  const matches = xml.match(re) || [];
  for (const sp of matches) {
    if (sp.includes(`<a:off x="${offX}" y="${offY}"/>`)) {
      return xml.replace(sp, '');
    }
  }
  return xml;
}

async function editSlide6(zip, leftRows, rightRows) {
  const path = 'ppt/slides/slide6.xml';
  let xml = await zip.file(path).async('string');
  xml = removeShapeByPos(xml, 63190, 1128323);
  xml = removeShapeByPos(xml, 4686290, 1128323);
  const bounds = { min: 7, max: 18 };
  const leftTbl = buildRundownTable(9001, 63190, 1128323, 4544665, leftRows, SLIDE6_AVAILABLE_HEIGHT, bounds, false);
  // Right table's original width (4544665) pushes its right edge past the slide
  // boundary — narrowed here so long text wraps instead of overflowing.
  const rightTbl = buildRundownTable(9002, 4686290, 1128323, 4350000, rightRows, SLIDE6_AVAILABLE_HEIGHT, bounds, false);
  xml = xml.replace('</p:spTree>', leftTbl + rightTbl + '</p:spTree>');
  zip.file(path, xml);
}

async function editSlide7(zip, leftRows, rightRows) {
  const path = 'ppt/slides/slide7.xml';
  let xml = await zip.file(path).async('string');
  xml = removeShapeByPos(xml, 164443, 783454);
  xml = removeShapeByPos(xml, 4421524, 656913);
  const bounds = { min: 7, max: 16 };
  const leftTbl = buildRundownTable(9003, 164443, 783454, 4501662, leftRows, SLIDE7_LEFT_AVAILABLE_HEIGHT, bounds, true);
  // small left padding on the time cell so it doesn't sit right at the column divider
  const rightTbl = buildRundownTable(9004, 4421524, 656913, 4369776, rightRows, SLIDE7_RIGHT_AVAILABLE_HEIGHT, bounds, true, 120000);
  xml = xml.replace('</p:spTree>', leftTbl + rightTbl + '</p:spTree>');
  zip.file(path, xml);
}

async function editSlide8Photos(zip, photos) {
  const rIds = ['rId3','rId4','rId5','rId6','rId7','rId8','rId9','rId10','rId11','rId12','rId13','rId14'];
  for (let i = 0; i < Math.min(photos.length, 12); i++) {
    await swapImage(zip, 8, rIds[i], photos[i].bytes, photos[i].ext);
  }
}

// --- Additional slides (Penginapan / Transportasi / Entertainment, etc.) ---
// Each entry: { categoryLabel, venueName, note, photos: [{bytes, ext}, ...] }
// Title renders as two lines: categoryLabel (big) on top, venueName (smaller) directly below.
const ADD_SLIDE_TITLE_COLOR = 'AB8645';

function buildAddSlideGridPositions(n, cols, startY) {
  cols = cols || 3;
  const margin = 372000, gap = 150000;
  const totalW = 9144000 - 2 * margin;
  const w = Math.floor((totalW - (cols - 1) * gap) / cols);
  const h = Math.floor(w * 0.48);
  const positions = [];
  for (let i = 0; i < n; i++) {
    const r = Math.floor(i / cols), c = i % cols;
    positions.push({ x: margin + c * (w + gap), y: startY + r * (h + gap), w, h });
  }
  return positions;
}

function buildAdditionalSlideXml(categoryLabel, venueName, note, picSpecs) {
  let picsXml = '';
  picSpecs.forEach((p, i) => {
    const shapeId = 9500 + i;
    picsXml += `<p:pic><p:nvPicPr><p:cNvPr id="${shapeId}" name="AddPhoto${shapeId}"/><p:cNvPicPr preferRelativeResize="0"><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${p.rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${p.x}" y="${p.y}"/><a:ext cx="${p.w}" cy="${p.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  });

  const titleXml = `<p:sp><p:nvSpPr><p:cNvPr id="9600" name="AddTitle"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="700000"/><a:ext cx="9144000" cy="430000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en" sz="2200" b="1" u="sng" dirty="0"><a:solidFill><a:srgbClr val="${ADD_SLIDE_TITLE_COLOR}"/></a:solidFill></a:rPr><a:t>${escapeXml(categoryLabel)}</a:t></a:r></a:p></p:txBody></p:sp>`;

  const subtitleXml = venueName ? `<p:sp><p:nvSpPr><p:cNvPr id="9604" name="AddVenueName"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="1160000"/><a:ext cx="9144000" cy="340000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en" sz="1400" b="1" dirty="0"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:rPr><a:t>${escapeXml(venueName)}</a:t></a:r></a:p></p:txBody></p:sp>` : '';

  const noteXml = note ? `<p:sp><p:nvSpPr><p:cNvPr id="9601" name="AddNote"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="372000" y="4400000"/><a:ext cx="8400000" cy="350000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="en" sz="1100" i="1" dirty="0"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:rPr><a:t>${escapeXml(note)}</a:t></a:r></a:p></p:txBody></p:sp>` : '';

  const footerXml = `<p:sp><p:nvSpPr><p:cNvPr id="9602" name="AddYear"/><p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="366100" y="4773588"/><a:ext cx="3539400" cy="282600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="en" sz="1000" b="1"/><a:t>2026</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="9603" name="AddWeb"/><p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="5251900" y="4773588"/><a:ext cx="3539400" cy="282600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="r"/><a:r><a:rPr lang="en" sz="1000" b="1"/><a:t>Navigasi Outdoor Activity.Com</a:t></a:r></a:p></p:txBody></p:sp>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="AddSlideRoot"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${footerXml}${titleXml}${subtitleXml}${picsXml}${noteXml}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

async function addAdditionalSlides(zip, additionalSlides) {
  if (!additionalSlides || additionalSlides.length === 0) return;

  let ct = await zip.file('[Content_Types].xml').async('string');
  let presRels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  let pres = await zip.file('ppt/presentation.xml').async('string');

  const insertAfter = '<p:sldId id="270" r:id="rId11"/>'; // slide8's sldId (fixed in this template)
  let sldIdInsertBlock = '';

  for (let idx = 0; idx < additionalSlides.length; idx++) {
    const entry = additionalSlides[idx];
    const slideNum = 17 + idx; // template has slides 1-16; new ones start at 17
    const slidePath = `ppt/slides/slide${slideNum}.xml`;
    const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;

    // Backward compatible: accept either the new {categoryLabel, venueName} shape
    // or the old combined {title} shape (treated as categoryLabel with no venue line).
    const categoryLabel = entry.categoryLabel !== undefined ? entry.categoryLabel : (entry.title || '');
    const venueName = entry.venueName || '';

    const photos = entry.photos || [];
    const startY = venueName ? 1600000 : 1380000; // extra room when the venue-name line is shown
    const positions = buildAddSlideGridPositions(photos.length, 3, startY);
    const picSpecs = photos.map((photo, i) => ({
      rid: `rId${100 + i}`, x: positions[i].x, y: positions[i].y, w: positions[i].w, h: positions[i].h,
    }));

    const slideXml = buildAdditionalSlideXml(categoryLabel, venueName, entry.note || '', picSpecs);

    let relEntries = photos.map((photo, i) => {
      const mediaName = `add_slide_${slideNum}_photo_${i}_${Date.now()}.${photo.ext}`;
      zip.file(`ppt/media/${mediaName}`, photo.bytes);
      return `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`;
    }).join('');
    relEntries += `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout81.xml"/>`;
    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relEntries}</Relationships>`;

    zip.file(slidePath, slideXml);
    zip.file(relsPath, relsXml);

    const overrideEntry = `<Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
    ct = ct.replace('</Types>', overrideEntry + '</Types>');

    const newRid = `rId${900 + idx}`;
    presRels = presRels.replace('</Relationships>', `<Relationship Id="${newRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/></Relationships>`);

    const newSldId = 9000 + idx;
    sldIdInsertBlock += `<p:sldId id="${newSldId}" r:id="${newRid}"/>`;
  }

  pres = pres.replace(insertAfter, insertAfter + sldIdInsertBlock);

  zip.file('[Content_Types].xml', ct);
  zip.file('ppt/_rels/presentation.xml.rels', presRels);
  zip.file('ppt/presentation.xml', pres);
}

async function buildPptx(templateArrayBuffer, data) {
  const JSZip = (typeof require !== 'undefined') ? require('jszip') : window.JSZip;
  const zip = await JSZip.loadAsync(templateArrayBuffer);
  await editSlide2(zip, data);
  await editSlide4(zip, data);
  if (data.facilityItems) await editSlide5FacilityList(zip, data.facilityItems);
  if (data.mealsItems) await editSlide5Meals(zip, data.mealsItems);
  if (data.fasilitasItems) await editSlide5Fasilitas(zip, data.fasilitasItems);
  if (data.slide5Photos) await editSlide5Photos(zip, data.slide5Photos);
  if (data.slide8Photos) await editSlide8Photos(zip, data.slide8Photos);
  if (data.rundownLeft && data.rundownRight) await editSlide6(zip, data.rundownLeft, data.rundownRight);
  if (data.gamesLeft && data.gamesRight) await editSlide7(zip, data.gamesLeft, data.gamesRight);
  if (data.additionalSlides && data.additionalSlides.length) await addAdditionalSlides(zip, data.additionalSlides);
  const out = await zip.generateAsync({
    type: (typeof window !== 'undefined') ? 'blob' : 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  return out;
}

if (typeof module !== 'undefined') {
  module.exports = {
    buildPptx, editSlide2, editSlide4, editSlide5FacilityList, editSlide5Meals, editSlide5Fasilitas,
    editSlide5Photos, editSlide8Photos, editSlide6, editSlide7, addAdditionalSlides, swapImage,
    replaceShapeFirstRunText, replaceByTextPrefix, rupiah,
  };
}
