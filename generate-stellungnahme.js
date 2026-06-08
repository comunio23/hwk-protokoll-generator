'use strict';
const {
  AlignmentType, Document, Footer, Header, ImageRun, Packer,
  PageNumber, Paragraph, Tab, TabStopPosition, TabStopType, TextRun,
  UnderlineType,
} = require('docx');
const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// ── Page geometry (DXA) ────────────────────────────────────────────────────
const PAGE_W  = 11906, PAGE_H = 16838;
const MARGINS = { top: 1843, right: 737, bottom: 1134, left: 1247 };

// ── Logo ───────────────────────────────────────────────────────────────────
const logoPath = path.join(__dirname, 'hwk-logo.jpg');
const logoData = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : null;
// Logo native: 160 × 27 px → display at 2.5× für gute Sichtbarkeit im Header
const LOGO_W = 320, LOGO_H = 54;

// ── Font helpers ───────────────────────────────────────────────────────────
const run = (text, opts = {}) => new TextRun({
  text,
  font:      { name: 'Arial' },
  size:      opts.sz   || 22,       // 11pt = 22 half-points
  bold:      opts.bold || false,
  italics:   opts.ital || false,
  color:     opts.color || undefined,
  underline: opts.ul ? { type: UnderlineType.SINGLE } : undefined,
});

const par = (children, opts = {}) => new Paragraph({
  children,
  alignment: opts.align || AlignmentType.LEFT,
  spacing:   { before: opts.before || 0, after: opts.after || 0 },
  ...(opts.tabStops ? { tabStops: opts.tabStops } : {}),
});

const empty = (before = 0, after = 0) => par([run('')], { before, after });

// ── Header: Logo rechts (oder Text-Platzhalter) ────────────────────────────
const headerChildren = logoData
  ? [new Paragraph({
      children: [new ImageRun({ data: logoData, transformation: { width: LOGO_W, height: LOGO_H } })],
      alignment: AlignmentType.RIGHT,
    })]
  : [par([run('Handwerkskammer Potsdam', { sz: 20, bold: true })], { align: AlignmentType.RIGHT })];

// ── Footer: Adresse + Seitenzahl ───────────────────────────────────────────
const footerPara1 = new Paragraph({
  children: [
    run('Handwerkskammer Potsdam · Charlottenstraße 34–36 · 14467 Potsdam',
        { sz: 18, color: '666666' }),
    new Tab(),
    run('Seite ', { sz: 18, color: '666666' }),
    new TextRun({ children: [PageNumber.CURRENT], font: { name: 'Arial' }, size: 18, color: '666666' }),
    run(' von ', { sz: 18, color: '666666' }),
    new TextRun({ children: [PageNumber.TOTAL_PAGES], font: { name: 'Arial' }, size: 18, color: '666666' }),
  ],
  tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
});

// ── Document body ──────────────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: MARGINS } },

    headers: { default: new Header({ children: headerChildren }) },
    footers: { default: new Footer({ children: [footerPara1] }) },

    children: [
      // Absender-Block
      par([run('Personalrat der Handwerkskammer Potsdam', { sz: 20 })]),
      par([run('Postfach 60 08 51 • 14408 Potsdam',  { sz: 20 })], { after: 600 }),

      // Empfänger-Block
      par([run('Hauptgeschäftsführer')]),
      par([run('Herr Ralph Bührig')], { after: 400 }),

      // Ort & Datum
      par([run('Götz, den 19. Dezember 2025')], { align: AlignmentType.RIGHT, after: 400 }),

      // Bezugszeilen
      par([run('Ihr Schreiben vom:   TT.MM.JJJJ')]),
      par([run('Eingegangen am:     TT.MM.JJJJ')], { after: 600 }),

      // Betreff
      new Paragraph({
        children: [
          run('Stellungnahme des Personalrats zur beabsichtigten unbefristeten ', { bold: true }),
          run('Einstellung von Herrn/Frau Vorname Name', { bold: true }),
        ],
        spacing: { after: 400 },
      }),

      // Anrede
      par([run('Sehr geehrter Herr Bührig,')], { after: 200 }),

      // Einleitungssatz
      par([run('zu den von Ihnen vorgelegten Unterlagen nimmt der Personalrat wie folgt Stellung:')],
          { after: 300 }),

      // Abschnitt 1
      par([run('1.   Unbefristete Einstellung von Herrn/Frau Vorname Name', { bold: true })]),
      par([run('Der Personalrat stimmt der beabsichtigten unbefristeten Einstellung von Herrn/Frau Name '
             + 'als Benennung Tätigkeit ab dem TT.MM.JJJJ zu.')], { after: 300 }),

      // Abschnitt 2
      par([run('2.   Eingruppierung von Herrn/Frau Vorname Name', { bold: true })]),
      par([run('Der Personalrat stimmt der Eingruppierung von Herrn/Frau Name in die EG xxx zu.')],
          { after: 300 }),

      // Grußformel
      par([run('Für Rückfragen stehe ich Ihnen gern zur Verfügung.')]),
      empty(),
      par([run('Mit freundlichen Grüßen')], { after: 1000 }),

      // Unterschrift
      par([run('………………………………………………………………..')]),
      par([run('Niclas Grobheiser')]),
      par([run('Personalratsvorsitzender')]),
    ],
  }],
});

// ── Ausgabe ────────────────────────────────────────────────────────────────
async function main() {
  const dir     = __dirname;
  const docxOut = path.join(dir, 'Stellungnahme-PR.docx');

  fs.writeFileSync(docxOut, await Packer.toBuffer(doc));
  console.log('✓ DOCX erstellt:', docxOut);

  try {
    execSync(`soffice --headless --convert-to pdf "${docxOut}" --outdir "${dir}"`, { timeout: 30000 });
    console.log('✓ PDF  erstellt: Stellungnahme-PR.pdf');
  } catch {
    console.log('  (PDF übersprungen — LibreOffice nicht gefunden oder fehlgeschlagen)');
  }
}

main().catch(e => { console.error(e); process.exit(1); });
