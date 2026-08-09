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

// --- Slide 5: facility "Include" list — replaces the 8 known runs (with the
// fragmented "Games M"/"a"/"ster " triple handled the same way as in Python) ---
async function editSlide5FacilityList(zip, items) {
  // items: array of up to 6 strings, in order (Tiket Area, ..., last item)
  const path = 'ppt/slides/slide5.xml';
  let xml = await zip.file(path).async('string');

  const spRegex = /(<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<p:cNvPr id="1240"[\s\S]*?<\/p:sp>)/;
  const match = xml.match(spRegex);
  if (!match) return;
  let spXml = match[1];

  // Grab all <a:t> runs within the second paragraph and overwrite them in order,
  // mirroring the 8-run layout (L1,L2,[L3 in 3 runs],L4,L5,L6) from the template.
  const runTexts = [
    items[0] || '', items[1] || '',
    items[2] || '', '', '',       // L3 collapses into first run, rest blanked
    items[3] || '', items[4] || '', items[5] || '',
  ];
  let i = 0;
  const paragraphs = spXml.split('</a:p>');
  // second paragraph is index 1 (index 0 is the header "Include :")
  if (paragraphs[1]) {
    let para = paragraphs[1];
    para = para.replace(/<a:t>[\s\S]*?<\/a:t>/g, () => `<a:t>${escapeXml(runTexts[i++] ?? '')}</a:t>`);
    paragraphs[1] = para;
  }
  const newSpXml = paragraphs.join('</a:p>');
  xml = xml.replace(spXml, newSpXml);
  zip.file(path, xml);
}

async function editSlide5Photos(zip, photos) {
  // photos: array of up to 4 {bytes, ext}
  const rIds = ['rId3', 'rId4', 'rId5', 'rId6'];
  for (let i = 0; i < Math.min(photos.length, 4); i++) {
    await swapImage(zip, 5, rIds[i], photos[i].bytes, photos[i].ext);
  }
}

// --- Slide 6 & 7: rebuild the two rundown columns as native OOXML tables ---
const RUNDOWN_COLOR = 'AB8645';
const TIME_COL_W = 1150000;
const INDENT_STEP = 180000;

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

function buildTcPr() {
  return `<a:tcPr marL="0" marR="0" marT="30000" marB="160000"><a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT><a:lnB w="0"><a:noFill/></a:lnB><a:noFill/></a:tcPr>`;
}

function buildRun(text, sz, italic) {
  return `<a:r><a:rPr sz="${sz}" b="1" i="${italic ? 1 : 0}"><a:solidFill><a:srgbClr val="${RUNDOWN_COLOR}"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>${escapeXml(text)}</a:t></a:r>`;
}

function buildDescCell(lines, sz, nested) {
  let paras = '';
  lines.forEach((raw, i) => {
    const { level, italic, text } = parseRundownLine(raw);
    const marL = nested ? INDENT_STEP * level : 0;
    const spcBef = i === 0 ? 0 : 600;
    const pPr = `<a:pPr${marL ? ` marL="${marL}" indent="0"` : ''}><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="${spcBef}"/></a:spcBef></a:pPr>`;
    paras += `<a:p>${pPr}${buildRun(text, sz, italic)}</a:p>`;
  });
  return `<a:tc><a:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${paras}</a:txBody>${buildTcPr()}</a:tc>`;
}

function buildTimeCell(timeText, sz) {
  return `<a:tc><a:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p>${buildRun(timeText, sz, false)}</a:p></a:txBody>${buildTcPr()}</a:tc>`;
}

function buildRundownTable(shapeId, x, y, width, rows, sz, nested) {
  const totalHeight = rows.length * 300000;
  const trs = rows.map(row => {
    const timeText = row.time2 ? `${row.time1}-${row.time2}` : row.time1;
    const lines = row.label.split('\n').map(l => l.trim()).filter(Boolean);
    return `<a:tr h="300000">${buildTimeCell(timeText, sz)}${buildDescCell(lines, sz, nested)}</a:tr>`;
  }).join('');

  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${shapeId}" name="Table ${shapeId}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${totalHeight}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="0" bandRow="0"/><a:tblGrid><a:gridCol w="${TIME_COL_W}"/><a:gridCol w="${width - TIME_COL_W}"/></a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
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
  const leftTbl = buildRundownTable(9001, 63190, 1128323, 4544665, leftRows, 1200, false);
  const rightTbl = buildRundownTable(9002, 4686290, 1128323, 4544665, rightRows, 1200, false);
  xml = xml.replace('</p:spTree>', leftTbl + rightTbl + '</p:spTree>');
  zip.file(path, xml);
}

async function editSlide7(zip, leftRows, rightRows) {
  const path = 'ppt/slides/slide7.xml';
  let xml = await zip.file(path).async('string');
  xml = removeShapeByPos(xml, 164443, 783454);
  xml = removeShapeByPos(xml, 4421524, 656913);
  const leftTbl = buildRundownTable(9003, 164443, 783454, 4501662, leftRows, 1000, true);
  const rightTbl = buildRundownTable(9004, 4421524, 656913, 4369776, rightRows, 1000, true);
  xml = xml.replace('</p:spTree>', leftTbl + rightTbl + '</p:spTree>');
  zip.file(path, xml);
}

async function editSlide8Photos(zip, photos) {
  const rIds = ['rId3','rId4','rId5','rId6','rId7','rId8','rId9','rId10','rId11','rId12','rId13','rId14'];
  for (let i = 0; i < Math.min(photos.length, 12); i++) {
    await swapImage(zip, 8, rIds[i], photos[i].bytes, photos[i].ext);
  }
}

async function buildPptx(templateArrayBuffer, data) {
  const JSZip = (typeof require !== 'undefined') ? require('jszip') : window.JSZip;
  const zip = await JSZip.loadAsync(templateArrayBuffer);
  await editSlide2(zip, data);
  await editSlide4(zip, data);
  if (data.facilityItems) await editSlide5FacilityList(zip, data.facilityItems);
  if (data.slide5Photos) await editSlide5Photos(zip, data.slide5Photos);
  if (data.slide8Photos) await editSlide8Photos(zip, data.slide8Photos);
  if (data.rundownLeft && data.rundownRight) await editSlide6(zip, data.rundownLeft, data.rundownRight);
  if (data.gamesLeft && data.gamesRight) await editSlide7(zip, data.gamesLeft, data.gamesRight);
  const out = await zip.generateAsync({
    type: (typeof window !== 'undefined') ? 'blob' : 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  return out;
}

if (typeof module !== 'undefined') {
  module.exports = {
    buildPptx, editSlide2, editSlide4, editSlide5FacilityList, editSlide5Photos, editSlide8Photos,
    editSlide6, editSlide7, swapImage, replaceShapeFirstRunText, replaceByTextPrefix, rupiah,
  };
}
