const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const dataPath = process.argv[2];
if (!dataPath) {
  console.error("Usage: node gen_deck.js <data.json>");
  process.exit(1);
}
const d = JSON.parse(fs.readFileSync(dataPath, "utf-8"));

const INK = "1C1F26";
const AMBER_L = "F0A63D";
const AMBER_D = "B8690A";
const BG = "FEFCF9";
const TXT = "565C68";
const MUTED = "9BA3B0";
const DIVIDER = "E7E4DE";
const WARN_BG = "FFF2DF";
const WARN_BORDER = "D97706";
const WARN_TEXT = "8A4B02";
const BRIEF_BG = "FFF6E8";
const BRIEF_TXT = "4A4030";

const FONT_HEAD = "Calibri";
const FONT_BODY = "Calibri";

const W = 13.333;
const H = 7.5;
const MARGIN = 0.7;
const CONTENT_W = W - 2 * MARGIN;

function pres() {
  const p = new pptxgen();
  p.defineLayout({ name: "WIDE", width: W, height: H });
  p.layout = "WIDE";
  return p;
}

function pageHeader(slide, label) {
  slide.addText(label, {
    x: MARGIN, y: 0.35, w: CONTENT_W, h: 0.3,
    fontFace: FONT_BODY, fontSize: 11, color: MUTED, bold: false,
  });
}

function addCover(p, d) {
  const s = p.addSlide();
  s.background = { color: INK };

  // decorative circle motif, top-right, partially off-slide
  s.addShape("ellipse", {
    x: W - 3.2, y: -2.6, w: 6.4, h: 6.4,
    fill: { color: AMBER_L, transparency: 88 },
    line: { type: "none" },
  });
  s.addShape("ellipse", {
    x: -1.6, y: H - 2.0, w: 3.4, h: 3.4,
    fill: { color: AMBER_L, transparency: 92 },
    line: { type: "none" },
  });

  s.addText("KI-Update", {
    x: MARGIN, y: 0.9, w: 8, h: 0.8,
    fontFace: FONT_HEAD, fontSize: 40, bold: true, color: "FFFFFF",
  });
  s.addText(d.dateLabel, {
    x: W - MARGIN - 5, y: 0.98, w: 5, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 15, bold: true, color: AMBER_L,
    align: "right", charSpacing: 2,
  });
  s.addText(d.subtitle, {
    x: MARGIN, y: 1.75, w: 10.5, h: 0.6,
    fontFace: FONT_BODY, fontSize: 16, color: MUTED,
  });

  const cats = [
    { label: "Praxis-News", val: d.counts.praxis },
    { label: "Tipps & Tricks", val: d.counts.tipps },
    { label: "Regulierung/Politik", val: d.counts.reg },
    { label: "Bedeutende Forschung", val: d.counts.forschung },
  ];
  const kpiY = 5.0;
  const kpiW = CONTENT_W / 4;
  cats.forEach((c, i) => {
    const x = MARGIN + i * kpiW;
    const isEmpty = c.val === null || c.val === undefined;
    s.addText(isEmpty ? "—" : String(c.val), {
      x, y: kpiY, w: kpiW - 0.3, h: 0.75,
      fontFace: FONT_HEAD, fontSize: 44, bold: true,
      color: isEmpty ? TXT : AMBER_L,
    });
    s.addText(c.label.toUpperCase(), {
      x, y: kpiY + 0.78, w: kpiW - 0.3, h: 0.4,
      fontFace: FONT_BODY, fontSize: 11, bold: true, color: MUTED,
      charSpacing: 1,
    });
  });
}

function addBriefing(p, d) {
  if (!d.warnbox && !d.briefing) return;
  const s = p.addSlide();
  s.background = { color: BG };
  pageHeader(s, `KI-Update — ${d.dateLabel}`);

  let y = 0.9;
  if (d.warnbox) {
    s.addShape("roundRect", {
      x: MARGIN, y, w: CONTENT_W, h: 0.85,
      rectRadius: 0.06,
      fill: { color: WARN_BG },
      line: { color: WARN_BORDER, width: 1 },
    });
    s.addText("⚠ " + d.warnbox, {
      x: MARGIN + 0.3, y: y + 0.08, w: CONTENT_W - 0.6, h: 0.7,
      fontFace: FONT_BODY, fontSize: 14, bold: true, color: WARN_TEXT,
      valign: "middle",
    });
    y += 1.15;
  }

  if (d.briefing && d.briefing.length) {
    const briefH = H - y - 0.6;
    s.addShape("roundRect", {
      x: MARGIN, y, w: CONTENT_W, h: briefH,
      rectRadius: 0.08,
      fill: { color: BRIEF_BG },
      line: { type: "none" },
    });
    let innerY = y + 0.35;
    const blockH = (briefH - 0.6) / d.briefing.length;
    d.briefing.forEach((b) => {
      s.addText(b.label.toUpperCase(), {
        x: MARGIN + 0.4, y: innerY, w: CONTENT_W - 0.8, h: 0.3,
        fontFace: FONT_BODY, fontSize: 12, bold: true, color: AMBER_D,
        charSpacing: 1,
      });
      s.addText(b.text, {
        x: MARGIN + 0.4, y: innerY + 0.32, w: CONTENT_W - 0.8, h: blockH - 0.4,
        fontFace: FONT_BODY, fontSize: 13, color: BRIEF_TXT, lineSpacingMultiple: 1.3,
      });
      innerY += blockH;
    });
  }
}

function addCategorySlide(p, d, catLabel, items, dateLabel) {
  if (!items || !items.length) return;
  const s = p.addSlide();
  s.background = { color: BG };
  pageHeader(s, `KI-Update — ${dateLabel}`);

  s.addText(catLabel.toUpperCase(), {
    x: MARGIN, y: 0.85, w: CONTENT_W, h: 0.4,
    fontFace: FONT_HEAD, fontSize: 20, bold: true, color: INK, charSpacing: 1,
  });

  const top = 1.5;
  const bottom = 0.5;
  const avail = H - top - bottom;
  const itemH = avail / items.length;

  // Fixed, compact block height per item (title + summary + source),
  // then center that block inside the slot instead of stretching text
  // boxes to fill the slot (stretching detaches the source line from
  // the summary text when there are few items and the slot is large).
  const titleH = 0.4;
  const maxChars = items.reduce((m, it) => Math.max(m, it.summary.length), 0);
  const summaryLines = Math.min(3, Math.max(2, Math.ceil(maxChars / 140)));
  const summaryH = summaryLines * 0.29;
  const sourceH = 0.28;
  const blockH = titleH + summaryH + sourceH;

  items.forEach((it, i) => {
    const slotY = top + i * itemH;
    const y = slotY + Math.max(0, (itemH - blockH) / 2);
    // number badge
    s.addShape("ellipse", {
      x: MARGIN, y: y + 0.03, w: 0.34, h: 0.34,
      fill: { color: AMBER_L }, line: { type: "none" },
    });
    s.addText(String(i + 1), {
      x: MARGIN, y: y + 0.03, w: 0.34, h: 0.34,
      fontFace: FONT_HEAD, fontSize: 13, bold: true, color: INK,
      align: "center", valign: "middle", margin: 0,
    });

    const tx = MARGIN + 0.55;
    const tw = CONTENT_W - 0.55;
    s.addText(it.title, {
      x: tx, y: y, w: tw, h: titleH,
      fontFace: FONT_HEAD, fontSize: 15, bold: true, color: INK,
      valign: "top", margin: 0,
    });
    s.addText(it.summary, {
      x: tx, y: y + titleH, w: tw, h: summaryH,
      fontFace: FONT_BODY, fontSize: 11.5, color: TXT, lineSpacingMultiple: 1.2,
      valign: "top", margin: 0,
    });
    s.addText(it.source, {
      x: tx, y: y + titleH + summaryH, w: tw, h: sourceH,
      fontFace: FONT_BODY, fontSize: 10.5, color: AMBER_D, italic: true,
      valign: "top", margin: 0,
      hyperlink: it.url ? { url: it.url } : undefined,
    });
  });
}

function addRegForschungSlide(p, d, dateLabel) {
  const s = p.addSlide();
  s.background = { color: BG };
  pageHeader(s, `KI-Update — ${dateLabel}`);

  const colW = (CONTENT_W - 0.6) / 2;
  const leftX = MARGIN;
  const rightX = MARGIN + colW + 0.6;

  s.addShape("line", {
    x: MARGIN + colW + 0.3, y: 1.5, w: 0, h: H - 1.5 - 0.5,
    line: { color: DIVIDER, width: 0.75 },
  });

  function renderCol(x, label, items, note) {
    s.addText(label.toUpperCase(), {
      x, y: 0.85, w: colW, h: 0.4,
      fontFace: FONT_HEAD, fontSize: 16, bold: true, color: INK, charSpacing: 1,
    });
    if (items && items.length) {
      let y = 1.5;
      items.forEach((it) => {
        s.addText(it.title, {
          x, y, w: colW, h: 0.6,
          fontFace: FONT_HEAD, fontSize: 15, bold: true, color: INK,
          valign: "top", margin: 0,
        });
        y += 0.62;
        s.addText(it.summary, {
          x, y, w: colW, h: 2.6,
          fontFace: FONT_BODY, fontSize: 12, color: TXT, lineSpacingMultiple: 1.25,
          valign: "top", margin: 0,
        });
        y += 2.4;
        s.addText(it.source, {
          x, y, w: colW, h: 0.3,
          fontFace: FONT_BODY, fontSize: 10.5, color: AMBER_D, italic: true,
          valign: "top", margin: 0,
          hyperlink: it.url ? { url: it.url } : undefined,
        });
      });
    } else if (note) {
      s.addText(note, {
        x, y: 1.6, w: colW, h: 1.2,
        fontFace: FONT_BODY, fontSize: 13, italic: true, color: MUTED,
        valign: "top",
      });
    }
  }

  renderCol(leftX, "Regulierung/Politik", d.reg, d.regNote);
  renderCol(rightX, "Bedeutende Forschung", d.forschung, d.forschungNote);
}

function addOmittedSlide(p, d, dateLabel) {
  if (!d.omitted || !d.omitted.length) return;
  const s = p.addSlide();
  s.background = { color: BG };
  pageHeader(s, `KI-Update — ${dateLabel}`);

  s.addText("Weitere gesichtete, aber ausgelassene Punkte", {
    x: MARGIN, y: 0.85, w: CONTENT_W, h: 0.5,
    fontFace: FONT_HEAD, fontSize: 20, bold: true, color: INK,
  });

  const paras = d.omitted.map((o, i) => ({
    text: o.title + "  ",
    options: {
      bullet: true, color: AMBER_D, fontFace: FONT_BODY, fontSize: 13,
      hyperlink: o.url ? { url: o.url } : undefined,
      breakLine: false,
    },
  }));
  // Build rich paragraphs: title (link) + source in muted, one per line
  const richText = [];
  d.omitted.forEach((o, i) => {
    richText.push({
      text: o.title,
      options: {
        bullet: true, color: INK, fontFace: FONT_BODY, fontSize: 13.5,
        hyperlink: o.url ? { url: o.url } : undefined,
        breakLine: false,
      },
    });
    richText.push({
      text: `  (${o.source})`,
      options: {
        color: MUTED, fontFace: FONT_BODY, fontSize: 12, italic: true,
        breakLine: i === d.omitted.length - 1 ? false : true,
      },
    });
  });

  s.addText(richText, {
    x: MARGIN, y: 1.55, w: CONTENT_W, h: H - 1.55 - 0.5,
    valign: "top", lineSpacingMultiple: 1.5, paraSpaceAfter: 10,
  });
}

const p = pres();
addCover(p, d);
addBriefing(p, d);
addCategorySlide(p, d, "Praxis-News", d.praxis, d.dateLabel);
addCategorySlide(p, d, "Tipps & Tricks", d.tipps, d.dateLabel);
addRegForschungSlide(p, d, d.dateLabel);
addOmittedSlide(p, d, d.dateLabel);

p.writeFile({ fileName: d.outFile }).then(() => {
  console.log("Wrote " + d.outFile);
});
