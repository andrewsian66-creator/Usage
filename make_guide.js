'use strict';

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak,
  LevelFormat,
} = require('docx');
const fs = require('fs');

// ─── Constants ────────────────────────────────────────────────────────────────

const CONTENT_WIDTH = 10200; // DXA

const FONT = 'Arial';

// Border presets
const bSingle4  = (color = '000000') => ({ style: BorderStyle.SINGLE,  size: 4,  color, space: 0 });
const bSingle8  = (color = '000000') => ({ style: BorderStyle.SINGLE,  size: 8,  color, space: 0 });
const bSingle10 = (color = '000000') => ({ style: BorderStyle.SINGLE,  size: 10, color, space: 8 });
const bSingle12 = (color = '000000') => ({ style: BorderStyle.SINGLE,  size: 12, color, space: 0 });
const bDouble4  = (color = '000000') => ({ style: BorderStyle.DOUBLE,  size: 4,  color, space: 0 });
const bNone     = ()                  => ({ style: BorderStyle.NONE,    size: 0,  color: 'auto', space: 0 });

// All-sides thin border (data tables)
const allThin = () => ({ top: bSingle4(), bottom: bSingle4(), left: bSingle4(), right: bSingle4() });

// Transparent shading (no fill)
const noFill = () => ({ fill: 'FFFFFF', type: ShadingType.CLEAR, color: 'auto' });

// Standard cell margins
const cellMargins = (top = 80, bottom = 80, left = 120, right = 120) =>
  ({ top, bottom, left, right });

// Helper: plain run
const run = (text, opts = {}) =>
  new TextRun({ text, font: FONT, size: 20, color: '000000', ...opts });

// Helper: body paragraph
const bodyPara = (children, opts = {}) =>
  new Paragraph({ spacing: { after: 120 }, children, ...opts });

// Helper: bold paragraph (used for table titles etc.)
const boldPara = (text, alignment = AlignmentType.LEFT) =>
  new Paragraph({
    spacing: { after: 120 },
    alignment,
    children: [run(text, { bold: true })],
  });

// ─── 4.1 Day Book Table ────────────────────────────────────────────────────────
//
// Columns: Date | Customer/Supplier | Invoice No. | £
// Widths:  1530 | 4080              | 2040        | 2550  = 10200
//
// blankTotal: if true, the total row amount cell is empty (practice question).
//             if false, total is auto-calculated from data rows.

function makeDayBookTable(title, rows, { blankTotal = false } = {}) {
  const COL_W = [1530, 4080, 2040, 2550];

  const makeCell = (text, { bold = false, rightAlign = false, width } = {}) =>
    new TableCell({
      borders:  allThin(),
      shading:  noFill(),
      margins:  cellMargins(),
      width:    { size: width, type: WidthType.DXA },
      children: [new Paragraph({
        alignment: rightAlign ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children:  [run(text, { bold })],
      })],
    });

  // Header row
  const headerRow = new TableRow({
    children: [
      makeCell('Date',       { bold: true, width: COL_W[0] }),
      makeCell('Customer/Supplier', { bold: true, width: COL_W[1] }),
      makeCell('Invoice No.', { bold: true, width: COL_W[2] }),
      makeCell('£',     { bold: true, rightAlign: true, width: COL_W[3] }),
    ],
  });

  // Data rows
  const dataRows = rows.map(({ date, name, invoice, amount }) =>
    new TableRow({
      children: [
        makeCell(date    ?? '', { width: COL_W[0] }),
        makeCell(name    ?? '', { width: COL_W[1] }),
        makeCell(invoice ?? '', { width: COL_W[2] }),
        makeCell(amount  ?? '', { rightAlign: true, width: COL_W[3] }),
      ],
    }),
  );

  // Total row
  const total = blankTotal
    ? ''
    : rows.reduce((sum, r) => {
        const n = parseFloat((r.amount ?? '0').replace(/[^\d.]/g, ''));
        return sum + (isNaN(n) ? 0 : n);
      }, 0).toLocaleString('en-GB');

  const totalRow = new TableRow({
    children: [
      makeCell('Total', { bold: true, width: COL_W[0] }),
      makeCell('',      { width: COL_W[1] }),
      makeCell('',      { width: COL_W[2] }),
      makeCell(blankTotal ? '' : `£${total}`, { bold: !blankTotal, rightAlign: true, width: COL_W[3] }),
    ],
  });

  return [
    boldPara(title),
    new Table({
      width:        { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: COL_W,
      rows:         [headerRow, ...dataRows, totalRow],
    }),
    bodyPara([run('')]), // spacing after
  ];
}

// ─── 4.2 T-Account ─────────────────────────────────────────────────────────────
//
// 4 columns: Dr desc | Dr amount | Cr desc | Cr amount
// Widths:    3900     | 1200      | 3900    | 1200  = 10200
//
// Row entry shapes:
//   {}                                        — blank filler
//   { desc, amount }                          — standard entry
//   { isTotal: true, amount }                 — total row (top: thin, bottom: double)
//   { bold: true, desc, amount }              — balance b/d (bold)
//
// The centre vertical line is the right border of col 2 / left border of col 3 (size 12).
// Header row bottom: thick single size 12.
// All outer borders: NONE.

function makeTAccount(title, drEntries, crEntries) {
  const COL_W = [3900, 1200, 3900, 1200];

  // Pad shorter side with blank entries
  const len = Math.max(drEntries.length, crEntries.length);
  const dr  = [...drEntries];
  const cr  = [...crEntries];
  while (dr.length < len) dr.push({});
  while (cr.length < len) cr.push({});

  // ── Cell builders ──────────────────────────────────────────

  // Outer vertical borders are always NONE; centre line handled per-cell.
  // leftBorderStyle, rightBorderStyle: pass bNone() or bSingle12() etc.
  function makeDescCell(entry, { isHeader = false, topBorder = bNone(), bottomBorder = bNone(),
                                  leftBorder = bNone(), rightBorder = bNone(), colWidth } = {}) {
    const isFiller = !entry.desc && !entry.amount && !entry.isTotal;
    const bold     = isHeader || entry.bold || false;
    const text     = isHeader ? (colWidth === COL_W[0] ? 'Dr' : 'Cr') : (entry.desc ?? '');

    return new TableCell({
      borders: { top: topBorder, bottom: bottomBorder, left: leftBorder, right: rightBorder },
      shading: noFill(),
      margins: cellMargins(60, 60, 100, 60),
      width:   { size: colWidth, type: WidthType.DXA },
      children: [new Paragraph({
        alignment: AlignmentType.LEFT,
        children:  [run(text, { bold, size: isHeader ? 22 : 20 })],
      })],
    });
  }

  function makeAmtCell(entry, { isHeader = false, topBorder = bNone(), bottomBorder = bNone(),
                                 leftBorder = bNone(), rightBorder = bNone(), colWidth } = {}) {
    const bold = isHeader || entry.bold || entry.isTotal || false;
    const text = isHeader ? '' : (entry.amount ?? '');

    return new TableCell({
      borders: { top: topBorder, bottom: bottomBorder, left: leftBorder, right: rightBorder },
      shading: noFill(),
      margins: cellMargins(60, 60, 60, 100),
      width:   { size: colWidth, type: WidthType.DXA },
      children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children:  [run(text, { bold })],
      })],
    });
  }

  // ── Header row ─────────────────────────────────────────────
  // Bottom border of header: thick single (size 12) on all 4 cells.
  // Centre vertical line: right of col2, left of col3 — size 12.
  const headerRow = new TableRow({
    children: [
      makeDescCell({}, { isHeader: true, bottomBorder: bSingle12(), colWidth: COL_W[0] }),
      makeAmtCell ({}, { isHeader: true, bottomBorder: bSingle12(), rightBorder: bSingle12(), colWidth: COL_W[1] }),
      makeDescCell({}, { isHeader: true, bottomBorder: bSingle12(), leftBorder:  bSingle12(), colWidth: COL_W[2] }),
      makeAmtCell ({}, { isHeader: true, bottomBorder: bSingle12(), colWidth: COL_W[3] }),
    ],
  });

  // ── Data rows ──────────────────────────────────────────────
  const dataRows = dr.map((drEntry, i) => {
    const crEntry = cr[i];
    const isTotal = drEntry.isTotal || crEntry.isTotal;

    // Total row: thin top border on all cells; double bottom border on all cells.
    const topB    = isTotal ? bSingle4()  : bNone();
    const bottomB = isTotal ? bDouble4()  : bNone();

    return new TableRow({
      children: [
        makeDescCell(drEntry, { topBorder: topB, bottomBorder: bottomB,                                    colWidth: COL_W[0] }),
        makeAmtCell (drEntry, { topBorder: topB, bottomBorder: bottomB, rightBorder: bSingle12(),           colWidth: COL_W[1] }),
        makeDescCell(crEntry, { topBorder: topB, bottomBorder: bottomB, leftBorder:  bSingle12(),           colWidth: COL_W[2] }),
        makeAmtCell (crEntry, { topBorder: topB, bottomBorder: bottomB,                                    colWidth: COL_W[3] }),
      ],
    });
  });

  return [
    boldPara(title, AlignmentType.CENTER),
    new Table({
      width:        { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: COL_W,
      rows:         [headerRow, ...dataRows],
    }),
    bodyPara([run('')]), // spacing after
  ];
}

// ─── Exports ───────────────────────────────────────────────────────────────────

module.exports = {
  makeDayBookTable,
  makeTAccount,
  // shared helpers (available for later content sections)
  run,
  bodyPara,
  boldPara,
  noFill,
  allThin,
  cellMargins,
  bSingle4, bSingle8, bSingle10, bSingle12, bDouble4, bNone,
  CONTENT_WIDTH,
  FONT,
};

// ─── Quick smoke test (node make_guide.js) ─────────────────────────────────────
if (require.main === module) {
  const children = [
    ...makeDayBookTable('Sales Day Book — May 2026', [
      { date: '1 May', name: 'A. Brown',  invoice: 'SDB/101', amount: '400' },
      { date: '3 May', name: 'C. Davies', invoice: 'SDB/102', amount: '180' },
      { date: '7 May', name: 'A. Brown',  invoice: 'SDB/103', amount: '250' },
    ]),
    ...makeTAccount('A. Brown — Sales Ledger', [
      { desc: '1 May  Sales SDB/101', amount: '£400' },
      { desc: '7 May  Sales SDB/103', amount: '250' },
      { isTotal: true, amount: '£650' },
      { desc: 'Bal b/d', amount: '£650', bold: true },
    ], [
      {},
      { desc: '7 May  Balance c/d', amount: '650' },
      { isTotal: true, amount: '£650' },
      {},
    ]),
  ];

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size:   { width: 11906, height: 16838 },
          margin: { top: 1080, right: 850, bottom: 1080, left: 850 },
        },
      },
      children,
    }],
  });

  Packer.toBuffer(doc).then(buf => {
    const out = '/mnt/user-data/outputs/smoke_test.docx';
    fs.mkdirSync('/mnt/user-data/outputs', { recursive: true });
    fs.writeFileSync(out, buf);
    console.log('Smoke test written to', out);
  });
}
