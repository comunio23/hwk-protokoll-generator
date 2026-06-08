'use strict';
const {
  AlignmentType, BorderStyle, Document, HeightRule, Packer,
  Paragraph, ShadingType, Tab, TabStopType,
  Table, TableCell, TableRow, TextRun, WidthType,
} = require('docx');
const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// ── Page geometry (DXA) ────────────────────────────────────────────────────
const PAGE_W  = 11906, PAGE_H = 16838;
const MARGINS = { top: 1417, right: 1417, bottom: 1134, left: 1417 };
const COL_W   = 4336;
const TBL_W   = COL_W * 2;   // 8672

// ── Shared cell border / padding ───────────────────────────────────────────
const B = (s = BorderStyle.SINGLE) =>
  ({ style: s, size: 4, color: '000000' });
const BORDER = { top: B(), bottom: B(), left: B(), right: B() };
const PAD    = { top: 80, bottom: 80, left: 120, right: 120 };
const SHAD   = { type: ShadingType.CLEAR, color: 'auto', fill: 'FFFFFF' };

// ── Helpers ────────────────────────────────────────────────────────────────
const run = (text, bold = false, sz = 22) =>
  new TextRun({ text, font: { name: 'Arial' }, size: sz, bold });

const par = (children, after = 0, before = 0, align = AlignmentType.LEFT, tabStops = []) =>
  new Paragraph({ children, alignment: align, spacing: { before, after },
                  ...(tabStops.length ? { tabStops } : {}) });

function tcell(content, opts = {}) {
  const cs = opts.cs || 1;
  return new TableCell({
    ...(cs > 1 ? { columnSpan: cs } : {}),
    width:   { size: cs === 2 ? TBL_W : COL_W, type: WidthType.DXA },
    borders: BORDER, margins: PAD, shading: SHAD,
    children: [par(
      Array.isArray(content) ? content : [run(content, opts.bold)],
      0, 0,
      opts.align || AlignmentType.LEFT,
    )],
  });
}

function emptyDataRow() {
  const ecell = () => new TableCell({
    width: { size: COL_W, type: WidthType.DXA },
    borders: BORDER, margins: PAD, shading: SHAD,
    children: [par([run('')])],
  });
  return new TableRow({
    height: { value: 600, rule: HeightRule.EXACT },
    children: [ecell(), ecell()],
  });
}

function makeTable(title, dataRows) {
  return new Table({
    width: { size: TBL_W, type: WidthType.DXA },
    columnWidths: [COL_W, COL_W],
    rows: [
      new TableRow({ children: [tcell(title, { cs: 2, bold: true, align: AlignmentType.CENTER })] }),
      new TableRow({ children: [tcell('Name, Vorname'), tcell('Unterschrift')] }),
      ...Array.from({ length: dataRows }, emptyDataRow),
    ],
  });
}

// ── Document ───────────────────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: MARGINS } },
    children: [
      // Kopfbereich
      par([run('Handwerkskammer Potsdam')]),
      par([run('Der Personalrat')], 500),

      // Titel
      par([run('Anwesenheitsliste Personalratssitzung', true, 28)], 400),

      // Metadaten-Zeile mit Tab-Stop
      new Paragraph({
        children: [
          run('am:   ………………………………….', true),
          new Tab(),
          run('in:   ……………………………………..', true),
        ],
        tabStops: [{ type: TabStopType.LEFT, position: 4000 }],
        spacing: { before: 0, after: 500 },
      }),

      // Tabelle 1: Personalratsmitglieder (5 Datenzeilen)
      makeTable('Personalratsmitglieder', 5),

      // minimaler Zwischenabstand zwischen den Tabellen
      par([run('')], 0, 0),

      // Tabelle 2: Sonstige Teilnehmer (9 Datenzeilen)
      makeTable('Sonstige Teilnehmer', 9),

      // Unterschriftsbereich
      par([run('')], 0, 800),   // 800 DXA Abstand davor
      par([run('_________________________________')]),
      par([run('Personalratsvorsitzender')]),
    ],
  }],
});

// ── Ausgabe ────────────────────────────────────────────────────────────────
async function main() {
  const dir     = __dirname;
  const docxOut = path.join(dir, 'Anwesenheitsliste-PR.docx');

  fs.writeFileSync(docxOut, await Packer.toBuffer(doc));
  console.log('✓ DOCX erstellt:', docxOut);

  try {
    execSync(`soffice --headless --convert-to pdf "${docxOut}" --outdir "${dir}"`, { timeout: 30000 });
    console.log('✓ PDF  erstellt: Anwesenheitsliste-PR.pdf');
  } catch {
    console.log('  (PDF übersprungen — LibreOffice nicht gefunden oder fehlgeschlagen)');
  }
}

main().catch(e => { console.error(e); process.exit(1); });
