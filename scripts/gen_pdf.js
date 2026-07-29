const fs = require("fs");
const PDFDocument = require("pdfkit");

const dataPath = process.argv[2];
if (!dataPath) {
  console.error("Usage: node gen_pdf.js <data.json>");
  process.exit(1);
}
const d = JSON.parse(fs.readFileSync(dataPath, "utf-8"));

// PDFKit requires the "#" prefix on hex colors (opposite of pptxgenjs, which forbids it).
const INK = "#1C1F26";
const AMBER_L = "#F0A63D";
const AMBER_D = "#B8690A";
const TXT = "#444B57";
const MUTED = "#7A828E";
const WHITE = "#FFFFFF";
const HEADER_MUTED = "#9BA3B0";
const EMPTY_GREY = "#565C68";
const WARN_BG = "#FFF2DF";
const WARN_BORDER = "#D97706";
const WARN_TEXT = "#8A4B02";
const BRIEF_BG = "#FFF6E8";
const BRIEF_TXT = "#4A4030";

// Helvetica (WinAnsiEncoding) renders umlauts, en-dashes and curly quotes fine,
// but has no glyph for → or ⚠ — those two render as garbage, so replace them.
function sanitize(s) {
  if (!s) return s;
  return s.replace(/→/g, "->").replace(/⚠\s*/g, "");
}

const M = 50; // page margin
const doc = new PDFDocument({ size: "A4", margins: { top: M, bottom: M, left: M, right: M } });
const outFile = d.outFile.replace(/\.pptx$/, ".pdf");
doc.pipe(fs.createWriteStream(outFile));

const pageW = doc.page.width;
const contentW = pageW - 2 * M;

function headerBand() {
  doc.rect(0, 0, pageW, 110).fill(INK);
  doc.fillColor(WHITE).font("Helvetica-Bold").fontSize(24).text("KI-Update", M, 30);
  doc.fillColor(AMBER_L).font("Helvetica-Bold").fontSize(11)
    .text(sanitize(d.dateLabel), M, 30, { width: contentW, align: "right" });
  doc.fillColor(HEADER_MUTED).font("Helvetica").fontSize(11).text(sanitize(d.subtitle), M, 62, { width: contentW });

  const cats = [
    { label: "Praxis-News", val: d.counts.praxis },
    { label: "Tipps & Tricks", val: d.counts.tipps },
    { label: "Regulierung/Politik", val: d.counts.reg },
    { label: "Bedeutende Forschung", val: d.counts.forschung },
  ];
  const kpiW = contentW / 4;
  cats.forEach((c, i) => {
    const x = M + i * kpiW;
    const isEmpty = c.val === null || c.val === undefined;
    doc.fillColor(isEmpty ? EMPTY_GREY : AMBER_L).font("Helvetica-Bold").fontSize(18)
      .text(isEmpty ? "-" : String(c.val), x, 86, { width: kpiW, align: "left" });
  });
  doc.y = 130;
}

function warnBox(text) {
  text = sanitize(text);
  const y = doc.y;
  const h = doc.heightOfString(text, { width: contentW - 30, font: "Helvetica-Bold", fontSize: 11 }) + 20;
  doc.rect(M, y, contentW, h).fillAndStroke(WARN_BG, WARN_BORDER);
  doc.fillColor(WARN_TEXT).font("Helvetica-Bold").fontSize(11)
    .text("Achtung: " + text, M + 15, y + 10, { width: contentW - 30 });
  doc.y = y + h + 15;
}

function briefingBox(blocks) {
  const startY = doc.y;
  let h = 20;
  blocks.forEach((b) => {
    h += 16 + doc.heightOfString(sanitize(b.text), { width: contentW - 30, font: "Helvetica", fontSize: 10 }) + 14;
  });
  doc.rect(M, startY, contentW, h).fill(BRIEF_BG);
  let y = startY + 15;
  blocks.forEach((b) => {
    doc.fillColor(AMBER_D).font("Helvetica-Bold").fontSize(9).text(b.label.toUpperCase(), M + 15, y, { width: contentW - 30 });
    y += 14;
    const text = sanitize(b.text);
    doc.fillColor(BRIEF_TXT).font("Helvetica").fontSize(10).text(text, M + 15, y, { width: contentW - 30 });
    y += doc.heightOfString(text, { width: contentW - 30, font: "Helvetica", fontSize: 10 }) + 14;
  });
  doc.y = startY + h + 15;
}

function ensureSpace(minH) {
  if (doc.y + minH > doc.page.height - M) {
    doc.addPage();
  }
}

function categorySection(label, items, note) {
  if ((!items || !items.length) && !note) return;
  ensureSpace(60);
  doc.moveDown(0.3);
  doc.fillColor(INK).font("Helvetica-Bold").fontSize(15).text(label, M, doc.y);
  doc.moveDown(0.5);

  if (items && items.length) {
    items.forEach((it, i) => {
      ensureSpace(70);
      doc.fillColor(INK).font("Helvetica-Bold").fontSize(12).text(`${i + 1}. ${sanitize(it.title)}`, M, doc.y, { width: contentW });
      doc.moveDown(0.15);
      doc.fillColor(TXT).font("Helvetica").fontSize(10).text(sanitize(it.summary), M, doc.y, { width: contentW, lineGap: 2 });
      doc.moveDown(0.1);
      doc.fillColor(AMBER_D).font("Helvetica-Oblique").fontSize(9)
        .text(it.source, M, doc.y, { width: contentW, link: it.url || undefined, underline: !!it.url });
      doc.moveDown(0.6);
    });
  } else if (note) {
    doc.fillColor(MUTED).font("Helvetica-Oblique").fontSize(10).text(sanitize(note), M, doc.y, { width: contentW });
    doc.moveDown(0.6);
  }
}

function omittedSection(items) {
  if (!items || !items.length) return;
  ensureSpace(60);
  doc.moveDown(0.3);
  doc.fillColor(INK).font("Helvetica-Bold").fontSize(13).text("Weitere gesichtete, aber ausgelassene Punkte", M, doc.y, { width: contentW });
  doc.moveDown(0.4);
  items.forEach((o) => {
    ensureSpace(28);
    doc.fillColor(INK).font("Helvetica").fontSize(9.5)
      .text("- " + sanitize(o.title), M, doc.y, { width: contentW, link: o.url || undefined, underline: !!o.url });
    doc.fillColor(MUTED).font("Helvetica-Oblique").fontSize(9).text(`  (${o.source})`, M + 12, doc.y, { width: contentW });
    doc.moveDown(0.3);
  });
}

headerBand();
if (d.warnbox) warnBox(d.warnbox);
if (d.briefing && d.briefing.length) briefingBox(d.briefing);

categorySection("Praxis-News", d.praxis);
categorySection("Tipps & Tricks", d.tipps, d.tippsNote);
categorySection("Regulierung/Politik", d.reg, d.regNote);
categorySection("Bedeutende Forschung", d.forschung, d.forschungNote);
omittedSection(d.omitted);

doc.end();
console.log("Wrote " + outFile);
