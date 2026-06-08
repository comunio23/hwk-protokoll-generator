'use strict';
const {
  AlignmentType, Document, HeightRule, Packer,
  Paragraph, Tab, TabStopType, TextRun,
} = require('docx');
const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// ── Page geometry (DXA) ────────────────────────────────────────────────────
const PAGE_W  = 11906, PAGE_H = 16838;
const MARGINS = { top: 1247, right: 1134, bottom: 1021, left: 1418 };

// ── Font helpers (Calibri) ─────────────────────────────────────────────────
const run = (text, bold = false, sz = 22) =>
  new TextRun({ text, font: { name: 'Calibri' }, size: sz, bold });

const par = (children, opts = {}) => new Paragraph({
  children,
  alignment: opts.align || AlignmentType.LEFT,
  spacing:   { before: opts.before || 0, after: opts.after || 0 },
  ...(opts.tabStops ? { tabStops: opts.tabStops } : {}),
  ...(opts.indent   ? { indent: opts.indent }     : {}),
});

const empty = (before = 0, after = 0) => par([run('')], { before, after });

// ── TOP building block ─────────────────────────────────────────────────────
function top(num, title, ...contentLines) {
  const blocks = [
    par([run(`TOP ${num}  ${title}`, true)]),
  ];
  for (const line of contentLines) {
    blocks.push(par([run(line || '')], { after: line === '' ? 180 : 0 }));
  }
  blocks.push(empty(0, 200));  // Abstand nach TOP
  return blocks;
}

// ── Signature row helper (Tab-Stop bei 5000 DXA) ──────────────────────────
const sigLine = (left, right) => new Paragraph({
  children: [run(left), new Tab(), run(right)],
  tabStops: [{ type: TabStopType.LEFT, position: 5000 }],
});

// ── Document ───────────────────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: MARGINS } },
    children: [
      // Kopfbereich
      par([run('Handwerkskammer Potsdam')]),
      par([run('Der Personalrat')], { after: 400 }),

      // Titel
      new Paragraph({
        children: [
          run('Niederschrift über die XX. Sitzung des Personalrats ', true),
          run('der Handwerkskammer Potsdam im Berichtszeitraum', true),
        ],
        spacing: { after: 300 },
      }),

      // Metadaten
      new Paragraph({
        children: [
          run('am:   ………………………………….'),
          new Tab(),
          run('in:   ……………………………………..'),
        ],
        tabStops: [{ type: TabStopType.LEFT, position: 4000 }],
      }),
      par([run('Beginn: ………… Uhr / Ende: …………. Uhr')]),
      par([run('Teilnehmende:')]),
      empty(),
      par([run('Protokollführung:')]),
      empty(0, 200),

      // Einleitungssatz
      par([run('Zur Sitzung lud der Personalratsvorsitzende mit schriftlicher Einladung vom Datum ein.')],
          { after: 400 }),

      // ── TOPs ────────────────────────────────────────────────────────────
      ...top(1,
        'Genehmigung, Änderung oder Ergänzung der Tagesordnung',
        'Die TO wird bestätigt.',
      ),

      ...top(2,
        'Feststellung der Beschlussfähigkeit',
        'Die Beschlussfähigkeit wird festgestellt.',
      ),

      ...top(3,
        'Berichte über die Tätigkeit des Vorstandes sowie zu Gesprächen mit Mitarbeitern, '
        + 'mit Vertretern der Abteilungen und mit der Geschäftsführung',
        '–   ', '',
      ),

      ...top(4,
        'Berichte der Personalratsmitglieder zur Teilnahme an Vorstellungsgesprächen',
        '–   ', '',
      ),

      ...top(5,
        'Beschlussfassung zu mitbestimmungs- oder mitwirkungspflichtigen Angelegenheiten',
        '–   Beschluss:',
        '',
        '–   Beschluss:',
        '',
      ),

      ...top(6,
        'Verschiedenes',
        '–   ', '',
      ),

      // ── Unterschrift ─────────────────────────────────────────────────────
      empty(800),
      sigLine(
        '__________________________________',
        '__________________________________',
      ),
      sigLine(
        'Personalratsvorsitzender',
        'protokollführendes Mitglied des Personalrats',
      ),
    ],
  }],
});

// ── Ausgabe ────────────────────────────────────────────────────────────────
async function main() {
  const dir     = __dirname;
  const docxOut = path.join(dir, 'PR-Sitzung-Vorlage.docx');

  fs.writeFileSync(docxOut, await Packer.toBuffer(doc));
  console.log('✓ DOCX erstellt:', docxOut);

  try {
    execSync(`soffice --headless --convert-to pdf "${docxOut}" --outdir "${dir}"`, { timeout: 30000 });
    console.log('✓ PDF  erstellt: PR-Sitzung-Vorlage.pdf');
  } catch {
    console.log('  (PDF übersprungen — LibreOffice nicht gefunden oder fehlgeschlagen)');
  }
}

main().catch(e => { console.error(e); process.exit(1); });
