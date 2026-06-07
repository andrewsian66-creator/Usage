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

// ─── 4.3 Trial Balance ─────────────────────────────────────────────────────────
//
// 3 columns: Account | Dr £ | Cr £
// Widths:    6300    | 1950  | 1950  = 10200
// Auto-calculates totals row from provided data.

function makeTrialBalance(title, rows) {
  const COL_W = [6300, 1950, 1950];

  const makeCell = (text, { bold = false, rightAlign = false, width } = {}) =>
    new TableCell({
      borders: allThin(),
      shading: noFill(),
      margins: cellMargins(),
      width:   { size: width, type: WidthType.DXA },
      children: [new Paragraph({
        alignment: rightAlign ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children:  [run(text, { bold })],
      })],
    });

  const headerRow = new TableRow({
    children: [
      makeCell('Account', { bold: true, width: COL_W[0] }),
      makeCell('Dr £',   { bold: true, rightAlign: true, width: COL_W[1] }),
      makeCell('Cr £',   { bold: true, rightAlign: true, width: COL_W[2] }),
    ],
  });

  const dataRows = rows.map(({ account, dr, cr }) =>
    new TableRow({
      children: [
        makeCell(account ?? '', { width: COL_W[0] }),
        makeCell(dr ?? '',      { rightAlign: true, width: COL_W[1] }),
        makeCell(cr ?? '',      { rightAlign: true, width: COL_W[2] }),
      ],
    }),
  );

  const sumSide = key => rows.reduce((s, r) => {
    const n = parseFloat((r[key] ?? '0').replace(/[^\d.]/g, ''));
    return s + (isNaN(n) ? 0 : n);
  }, 0).toLocaleString('en-GB');

  const totalRow = new TableRow({
    children: [
      makeCell('Total', { bold: true, width: COL_W[0] }),
      makeCell(`£${sumSide('dr')}`, { bold: true, rightAlign: true, width: COL_W[1] }),
      makeCell(`£${sumSide('cr')}`, { bold: true, rightAlign: true, width: COL_W[2] }),
    ],
  });

  return [
    boldPara(title),
    new Table({
      width:        { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: COL_W,
      rows:         [headerRow, ...dataRows, totalRow],
    }),
    bodyPara([run('')]),
  ];
}

// ─── 4.4 Cash Book Extract ─────────────────────────────────────────────────────
//
// 3 columns: Date | Customer/Supplier | £
// Widths:    1530 | 6120              | 2550  = 10200
// col2Label: 'Customer' for receipts, 'Supplier' for payments.
// Auto-calculates and shows total row.

function makeCashBookExtract(title, rows, { col2Label = 'Customer' } = {}) {
  const COL_W = [1530, 6120, 2550];

  const makeCell = (text, { bold = false, rightAlign = false, width } = {}) =>
    new TableCell({
      borders: allThin(),
      shading: noFill(),
      margins: cellMargins(),
      width:   { size: width, type: WidthType.DXA },
      children: [new Paragraph({
        alignment: rightAlign ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children:  [run(text, { bold })],
      })],
    });

  const headerRow = new TableRow({
    children: [
      makeCell('Date',      { bold: true, width: COL_W[0] }),
      makeCell(col2Label,   { bold: true, width: COL_W[1] }),
      makeCell('£',    { bold: true, rightAlign: true, width: COL_W[2] }),
    ],
  });

  const dataRows = rows.map(({ date, name, amount }) =>
    new TableRow({
      children: [
        makeCell(date   ?? '', { width: COL_W[0] }),
        makeCell(name   ?? '', { width: COL_W[1] }),
        makeCell(amount ?? '', { rightAlign: true, width: COL_W[2] }),
      ],
    }),
  );

  const total = rows.reduce((s, r) => {
    const n = parseFloat((r.amount ?? '0').replace(/[^\d.]/g, ''));
    return s + (isNaN(n) ? 0 : n);
  }, 0).toLocaleString('en-GB');

  const totalRow = new TableRow({
    children: [
      makeCell('Total', { bold: true, width: COL_W[0] }),
      makeCell('',      { width: COL_W[1] }),
      makeCell(`£${total}`, { bold: true, rightAlign: true, width: COL_W[2] }),
    ],
  });

  return [
    boldPara(title),
    new Table({
      width:        { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: COL_W,
      rows:         [headerRow, ...dataRows, totalRow],
    }),
    bodyPara([run('')]),
  ];
}

// ─── 4.5 Transaction Summary Table ─────────────────────────────────────────────
//
// 3 columns: Transaction | Source | £
// Widths:    7140        | 1530   | 1530  = 10200
// Header row bold; no auto-total.

function makeTransactionSummary(title, rows) {
  const COL_W = [7140, 1530, 1530];

  const makeCell = (text, { bold = false, rightAlign = false, width } = {}) =>
    new TableCell({
      borders: allThin(),
      shading: noFill(),
      margins: cellMargins(),
      width:   { size: width, type: WidthType.DXA },
      children: [new Paragraph({
        alignment: rightAlign ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children:  [run(text, { bold })],
      })],
    });

  const headerRow = new TableRow({
    children: [
      makeCell('Transaction', { bold: true, width: COL_W[0] }),
      makeCell('Source',      { bold: true, width: COL_W[1] }),
      makeCell('£',      { bold: true, rightAlign: true, width: COL_W[2] }),
    ],
  });

  const dataRows = rows.map(({ transaction, source, amount }) =>
    new TableRow({
      children: [
        makeCell(transaction ?? '', { width: COL_W[0] }),
        makeCell(source      ?? '', { width: COL_W[1] }),
        makeCell(amount      ?? '', { rightAlign: true, width: COL_W[2] }),
      ],
    }),
  );

  return [
    boldPara(title),
    new Table({
      width:        { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: COL_W,
      rows:         [headerRow, ...dataRows],
    }),
    bodyPara([run('')]),
  ];
}

// ─── 4.6 Note Box (Callout) ────────────────────────────────────────────────────
//
// Single-cell table, full content width.
// Left border only (size 8); no other borders; no fill.
// Cell margins: top 60, bottom 60, left 200, right 60.
// Contents: bold title paragraph + body paragraph.

function makeNoteBox(titleText, bodyText) {
  const leftOnlyBorders = {
    top:    bNone(),
    bottom: bNone(),
    left:   bSingle8(),
    right:  bNone(),
  };

  const cell = new TableCell({
    borders: leftOnlyBorders,
    shading: noFill(),
    margins: { top: 60, bottom: 60, left: 200, right: 60 },
    width:   { size: CONTENT_WIDTH, type: WidthType.DXA },
    children: [
      new Paragraph({ spacing: { after: 60 },  children: [run(titleText, { bold: true })] }),
      new Paragraph({ spacing: { after: 0 },   children: [run(bodyText)] }),
    ],
  });

  return [
    bodyPara([run('')]), // space before
    new Table({
      width:        { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [CONTENT_WIDTH],
      rows:         [new TableRow({ children: [cell] })],
    }),
    bodyPara([run('')]), // space after
  ];
}

// ─── 4.7 Reconciliation Table ──────────────────────────────────────────────────
//
// 4 columns: Track1 label | Track1 £ | Track2 label | Track2 £
// Widths:    4050         | 1050     | 4050         | 1050  = 10200
// Header row bold; last data row bold (agreed/corrected row).
// Amount columns right-aligned.

function makeReconciliationTable(title, rows) {
  const COL_W = [4050, 1050, 4050, 1050];

  const makeCell = (text, { bold = false, rightAlign = false, width } = {}) =>
    new TableCell({
      borders: allThin(),
      shading: noFill(),
      margins: cellMargins(),
      width:   { size: width, type: WidthType.DXA },
      children: [new Paragraph({
        alignment: rightAlign ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children:  [run(text, { bold })],
      })],
    });

  const headerRow = new TableRow({
    children: [
      makeCell(rows[0]?.header1 ?? 'Track 1', { bold: true, width: COL_W[0] }),
      makeCell('£',                        { bold: true, rightAlign: true, width: COL_W[1] }),
      makeCell(rows[0]?.header2 ?? 'Track 2',  { bold: true, width: COL_W[2] }),
      makeCell('£',                        { bold: true, rightAlign: true, width: COL_W[3] }),
    ],
  });

  const dataRows = rows.map(({ label1, amount1, label2, amount2 }, i) => {
    const isLast = i === rows.length - 1;
    return new TableRow({
      children: [
        makeCell(label1  ?? '', { bold: isLast, width: COL_W[0] }),
        makeCell(amount1 ?? '', { bold: isLast, rightAlign: true, width: COL_W[1] }),
        makeCell(label2  ?? '', { bold: isLast, width: COL_W[2] }),
        makeCell(amount2 ?? '', { bold: isLast, rightAlign: true, width: COL_W[3] }),
      ],
    });
  });

  return [
    boldPara(title),
    new Table({
      width:        { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: COL_W,
      rows:         [headerRow, ...dataRows],
    }),
    bodyPara([run('')]),
  ];
}

// ─── Exports ───────────────────────────────────────────────────────────────────

module.exports = {
  makeDayBookTable,
  makeTAccount,
  makeTrialBalance,
  makeCashBookExtract,
  makeTransactionSummary,
  makeNoteBox,
  makeReconciliationTable,
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

// ─── Heading helpers ──────────────────────────────────────────────────────────

function h1(text) {
  return new Paragraph({
    spacing: { before: 320, after: 180 },
    border:  { bottom: bSingle12() },
    children: [run(text.toUpperCase(), { bold: true, size: 32 })],
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    indent:  { left: 180 },
    border:  { left: { style: BorderStyle.SINGLE, size: 10, color: '000000', space: 8 } },
    children: [run(text, { bold: true, size: 26 })],
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 180, after: 80 },
    children: [run(text, { bold: true, size: 22 })],
  });
}

function pb() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Simple 2-column label/value table used in Part 5 opening position
function makeLabelValueTable(rows) {
  const COL_W = [8100, 2100];
  const makeCell = (text, { bold = false, rightAlign = false, width } = {}) =>
    new TableCell({
      borders: allThin(),
      shading: noFill(),
      margins: cellMargins(),
      width:   { size: width, type: WidthType.DXA },
      children: [new Paragraph({
        alignment: rightAlign ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children:  [run(text, { bold })],
      })],
    });
  return new Table({
    width:        { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: COL_W,
    rows: rows.map(({ label, value, bold = false }) =>
      new TableRow({
        children: [
          makeCell(label ?? '', { bold, width: COL_W[0] }),
          makeCell(value ?? '', { bold, rightAlign: true, width: COL_W[1] }),
        ],
      }),
    ),
  });
}

// Error analysis table for Q5.4 (4 cols)
function makeErrorAnalysisTable(rows) {
  const COL_W = [510, 2550, 3570, 3570];
  const makeCell = (text, { bold = false, width } = {}) =>
    new TableCell({
      borders: allThin(),
      shading: noFill(),
      margins: cellMargins(),
      width:   { size: width, type: WidthType.DXA },
      children: [new Paragraph({ children: [run(text, { bold })] })],
    });
  const header = new TableRow({ children: [
    makeCell('#',                              { bold: true, width: COL_W[0] }),
    makeCell('Error type',                     { bold: true, width: COL_W[1] }),
    makeCell('Track 1 (SLCA) adjustment',      { bold: true, width: COL_W[2] }),
    makeCell('Track 2 (Individual) adjustment',{ bold: true, width: COL_W[3] }),
  ]});
  const dataRows = rows.map(({ num, type, t1, t2 }) =>
    new TableRow({ children: [
      makeCell(num  ?? '', { width: COL_W[0] }),
      makeCell(type ?? '', { width: COL_W[1] }),
      makeCell(t1   ?? '', { width: COL_W[2] }),
      makeCell(t2   ?? '', { width: COL_W[3] }),
    ]}),
  );
  return [
    boldPara('Error analysis:'),
    new Table({ width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: COL_W, rows: [header, ...dataRows] }),
    bodyPara([run('')]),
  ];
}

// ─── Document content ─────────────────────────────────────────────────────────

function buildDocument() {
  const children = [];
  const add = (...items) => items.forEach(i => Array.isArray(i) ? children.push(...i) : children.push(i));

  // ══════════════════════════════════════════════════════════════════
  // TITLE PAGE
  // ══════════════════════════════════════════════════════════════════
  add(
    new Paragraph({ spacing: { before: 1440, after: 240 }, alignment: AlignmentType.CENTER,
      children: [run('Control Accounts', { bold: true, size: 42 })] }),
    new Paragraph({ spacing: { after: 180 }, alignment: AlignmentType.CENTER,
      children: [run('Building Understanding from First Principles', { italics: true, size: 24 })] }),
    new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.CENTER,
      children: [run('A-Level Accounting — Worked Examples and Practice Questions', { size: 20 })] }),
  );

  // ══════════════════════════════════════════════════════════════════
  // PART 1 — CREDIT SALES
  // ══════════════════════════════════════════════════════════════════
  add(
    pb(),
    h1('Part 1 — Credit Sales: From Individual Accounts to the Control Account'),
    bodyPara([run('When a business sells goods on credit, each transaction is recorded in the Sales Day Book. The total is posted to the Sales Account (Credit) in the General Ledger, and individual amounts are posted as debits to each customer’s personal account in the Sales Ledger. The worked example below shows this process.')]),

    h3('Worked Example 1.1 — The Sales Day Book and Individual Customer Accounts'),
    bodyPara([run('Company: Hartley Trading Ltd')]),

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

    ...makeTAccount('C. Davies — Sales Ledger', [
      { desc: '3 May  Sales SDB/102', amount: '£180' },
      { isTotal: true, amount: '£180' },
      { desc: 'Bal b/d', amount: '£180', bold: true },
    ], [
      { desc: '3 May  Balance c/d', amount: '180' },
      { isTotal: true, amount: '£180' },
      {},
    ]),

    ...makeTrialBalance('Summarised Trial Balance — 7 May 2026', [
      { account: 'A. Brown',  dr: '650', cr: '' },
      { account: 'C. Davies', dr: '180', cr: '' },
      { account: 'Sales',     dr: '',    cr: '830' },
    ]),

    h2('Introducing the Sales Ledger Control Account'),
    bodyPara([run('As a business grows, the number of credit customers increases. Maintaining a separate General Ledger account for each customer becomes impractical and clutters the Trial Balance. A single Sales Ledger Control Account (SLCA) summarises all credit customer transactions in total, replacing the many individual accounts.')]),

    h3('Worked Example 1.2 — The SLCA Replaces the Individual Accounts'),

    ...makeTAccount('Sales Ledger Control Account', [
      { desc: '7 May  Credit Sales', amount: '£830' },
      { isTotal: true, amount: '£830' },
      { desc: 'Bal b/d', amount: '£830', bold: true },
    ], [
      { desc: '7 May  Balance c/d', amount: '830' },
      { isTotal: true, amount: '£830' },
      {},
    ]),

    ...makeTrialBalance('Summarised Trial Balance — 7 May 2026', [
      { account: 'Sales Ledger Control Account', dr: '830', cr: '' },
      { account: 'Sales',                        dr: '',    cr: '830' },
    ]),

    h2('The Advantage, the Disadvantage, and the Receivables Ledger'),
    bodyPara([run('The SLCA simplifies the General Ledger and reduces double-entry postings. However, the Trial Balance now shows only the total owed (£830) — it cannot tell us that A. Brown owes £650 and C. Davies owes £180. To recover that detail, businesses maintain a Receivables Ledger — a separate set of individual customer accounts running alongside the General Ledger.')]),
    bodyPara([run('Crucially, these accounts are memorandum records only. They are not part of the double-entry system and do not appear in the Trial Balance. The SLCA is the account of record; the Receivables Ledger provides the supporting analysis.')]),

    h3('Worked Example 1.3 — Cash Receipts: Keeping Both Tracks in Step'),

    ...makeCashBookExtract('Cash Book (Bank Account) — Receipts', [
      { date: '10 May', name: 'A. Brown',  amount: '400' },
      { date: '12 May', name: 'C. Davies', amount: '180' },
    ], { col2Label: 'Customer' }),

    ...makeTAccount('Sales Ledger Control Account — After Receipts', [
      { desc: 'Bal b/f', amount: '£830' },
      { isTotal: true, amount: '£830' },
      { desc: 'Bal b/d', amount: '£250', bold: true },
    ], [
      { desc: '12 May  Bank', amount: '580' },
      { desc: '12 May  Balance c/d', amount: '250' },
      { isTotal: true, amount: '£830' },
      {},
    ]),

    ...makeTAccount('A. Brown — Receivables Ledger', [
      { desc: 'Bal b/f', amount: '£650' },
      { isTotal: true, amount: '£650' },
      { desc: 'Bal b/d', amount: '£250', bold: true },
    ], [
      { desc: '10 May  Bank', amount: '400' },
      { desc: 'Balance c/d', amount: '250' },
      { isTotal: true, amount: '£650' },
      {},
    ]),

    ...makeTAccount('C. Davies — Receivables Ledger', [
      { desc: 'Bal b/f', amount: '£180' },
      { isTotal: true, amount: '£180' },
    ], [
      { desc: '12 May  Bank', amount: '180' },
      { isTotal: true, amount: '£180' },
    ]),

    bodyPara([run('SLCA balance: £250. Receivables Ledger total: £250 + £0 = £250. Agreed. ✓')]),
    bodyPara([run('If the Receivables Ledger had not been updated: SLCA would show £250 while individual accounts would still total £830 — a discrepancy of £580.')]),

    ...makeNoteBox('Key Principle', 'The Receivables Ledger must be updated every time an entry is made in the Sales Ledger Control Account. The SLCA balance and the total of the individual Receivables Ledger balances must always agree.'),

    h3('Practice Questions 1.4–1.6'),
    bodyPara([run('Company: Fernwood Supplies')]),

    ...makeDayBookTable('Sales Day Book — June 2026', [
      { date: '1 Jun', name: 'M. Khan',   invoice: 'SDB/201', amount: '520' },
      { date: '4 Jun', name: 'P. Thomas', invoice: 'SDB/202', amount: '340' },
      { date: '9 Jun', name: 'M. Khan',   invoice: 'SDB/203', amount: '180' },
    ], { blankTotal: true }),

    h3('Question 1.4'),
    bodyPara([run('(a) Total the Sales Day Book.')]),
    bodyPara([run('(b) Post to the individual customer accounts in the Sales Ledger and balance both.')]),
    bodyPara([run('(c) Prepare a summarised Trial Balance as at 9 June.')]),

    h3('Question 1.5'),
    bodyPara([run('(a) Prepare a Sales Ledger Control Account for the same transactions.')]),
    bodyPara([run('(b) Prepare a more summarised Trial Balance.')]),

    ...makeCashBookExtract('Cash Book (Bank Account) — Receipts', [
      { date: '15 Jun', name: 'M. Khan',   amount: '520' },
      { date: '17 Jun', name: 'P. Thomas', amount: '200' },
    ], { col2Label: 'Customer' }),

    h3('Question 1.6'),
    bodyPara([run('(a) Update the SLCA to reflect the receipts. What is the closing balance?')]),
    bodyPara([run('(b) Update both individual accounts in the Receivables Ledger.')]),
    bodyPara([run('(c) Confirm that the SLCA balance agrees with the sum of the individual Receivables Ledger balances.')]),
    bodyPara([run('(d) What would the discrepancy have been if the Receivables Ledger had not been updated?')]),

    h2('Answers — 1.4–1.6'),
    h3('Answer 1.4'),
    bodyPara([run('Day Book total: £1,040')]),

    ...makeTAccount('M. Khan — Sales Ledger', [
      { desc: '1 Jun  Sales SDB/201', amount: '£520' },
      { desc: '9 Jun  Sales SDB/203', amount: '180' },
      { isTotal: true, amount: '£700' },
      { desc: 'Bal b/d', amount: '£700', bold: true },
    ], [
      {},
      { desc: '9 Jun  Balance c/d', amount: '700' },
      { isTotal: true, amount: '£700' },
      {},
    ]),

    ...makeTAccount('P. Thomas — Sales Ledger', [
      { desc: '4 Jun  Sales SDB/202', amount: '£340' },
      { isTotal: true, amount: '£340' },
      { desc: 'Bal b/d', amount: '£340', bold: true },
    ], [
      { desc: '4 Jun  Balance c/d', amount: '340' },
      { isTotal: true, amount: '£340' },
      {},
    ]),

    ...makeTrialBalance('Trial Balance — 9 June 2026', [
      { account: 'M. Khan',   dr: '700',   cr: '' },
      { account: 'P. Thomas', dr: '340',   cr: '' },
      { account: 'Sales',     dr: '',      cr: '1,040' },
    ]),

    h3('Answer 1.5'),

    ...makeTAccount('Sales Ledger Control Account', [
      { desc: '9 Jun  Credit Sales', amount: '£1,040' },
      { isTotal: true, amount: '£1,040' },
      { desc: 'Bal b/d', amount: '£1,040', bold: true },
    ], [
      { desc: '9 Jun  Balance c/d', amount: '1,040' },
      { isTotal: true, amount: '£1,040' },
      {},
    ]),

    ...makeTrialBalance('Trial Balance — 9 June 2026', [
      { account: 'Sales Ledger Control Account', dr: '1,040', cr: '' },
      { account: 'Sales',                        dr: '',      cr: '1,040' },
    ]),

    h3('Answer 1.6'),

    ...makeTAccount('Sales Ledger Control Account', [
      { desc: 'Bal b/f', amount: '£1,040' },
      { isTotal: true, amount: '£1,040' },
      { desc: 'Bal b/d', amount: '£320', bold: true },
    ], [
      { desc: '17 Jun  Bank', amount: '720' },
      { desc: '17 Jun  Balance c/d', amount: '320' },
      { isTotal: true, amount: '£1,040' },
      {},
    ]),

    ...makeTAccount('M. Khan — Receivables Ledger', [
      { desc: 'Bal b/f', amount: '£700' },
      { isTotal: true, amount: '£700' },
      { desc: 'Bal b/d', amount: '£180', bold: true },
    ], [
      { desc: '15 Jun  Bank', amount: '520' },
      { desc: 'Balance c/d', amount: '180' },
      { isTotal: true, amount: '£700' },
      {},
    ]),

    ...makeTAccount('P. Thomas — Receivables Ledger', [
      { desc: 'Bal b/f', amount: '£340' },
      { isTotal: true, amount: '£340' },
      { desc: 'Bal b/d', amount: '£140', bold: true },
    ], [
      { desc: '17 Jun  Bank', amount: '200' },
      { desc: 'Balance c/d', amount: '140' },
      { isTotal: true, amount: '£340' },
      {},
    ]),

    bodyPara([run('SLCA £320. Receivables Ledger: £180 + £140 = £320. Agreed. ✓')]),
    bodyPara([run('Without updating: SLCA £320 vs individual total £1,040. Discrepancy: £720.')]),
  );

  // ══════════════════════════════════════════════════════════════════
  // PART 2 — CREDIT PURCHASES
  // ══════════════════════════════════════════════════════════════════
  add(
    pb(),
    h1('Part 2 — Credit Purchases: The Mirror Image'),
    bodyPara([run('The same twin-track logic applies on the purchases side. Credit purchases flow from the Purchases Day Book to individual supplier accounts in the Purchases Ledger. A Purchases Ledger Control Account (PLCA) in the General Ledger replaces the individual supplier accounts there. The key difference is direction: suppliers are creditors, so the normal balances are reversed.')]),

    h3('Worked Example 2.1 — Purchases Day Book and Individual Supplier Accounts'),
    bodyPara([run('Company: Hartley Trading Ltd')]),

    ...makeDayBookTable('Purchases Day Book — May 2026', [
      { date: '2 May', name: 'T. Mills Ltd',   invoice: 'PDB/501', amount: '600' },
      { date: '5 May', name: 'B. Rogers & Co', invoice: 'PDB/502', amount: '350' },
      { date: '8 May', name: 'T. Mills Ltd',   invoice: 'PDB/503', amount: '150' },
    ]),

    ...makeTAccount('T. Mills Ltd — Purchases Ledger (creditor account — normal balance is Cr)', [
      { desc: 'Balance c/d', amount: '750' },
      { isTotal: true, amount: '£750' },
      {},
    ], [
      { desc: '2 May  Purchases PDB/501', amount: '600' },
      { desc: '8 May  Purchases PDB/503', amount: '150' },
      { isTotal: true, amount: '£750' },
      { desc: 'Bal b/d', amount: '£750', bold: true },
    ]),

    ...makeTAccount('B. Rogers & Co — Purchases Ledger', [
      { desc: 'Balance c/d', amount: '350' },
      { isTotal: true, amount: '£350' },
      {},
    ], [
      { desc: '5 May  Purchases PDB/502', amount: '350' },
      { isTotal: true, amount: '£350' },
      { desc: 'Bal b/d', amount: '£350', bold: true },
    ]),

    ...makeTrialBalance('Summarised Trial Balance — 8 May 2026', [
      { account: 'Purchases',       dr: '1,100', cr: '' },
      { account: 'T. Mills Ltd',    dr: '',      cr: '750' },
      { account: 'B. Rogers & Co',  dr: '',      cr: '350' },
    ]),

    h3('Worked Example 2.2 — The PLCA Replaces the Individual Accounts'),

    ...makeTAccount('Purchases Ledger Control Account', [
      { desc: '8 May  Balance c/d', amount: '1,100' },
      { isTotal: true, amount: '£1,100' },
      {},
    ], [
      { desc: '8 May  Credit Purchases', amount: '1,100' },
      { isTotal: true, amount: '£1,100' },
      { desc: 'Bal b/d', amount: '£1,100', bold: true },
    ]),

    ...makeTrialBalance('Summarised Trial Balance — 8 May 2026', [
      { account: 'Purchases',                        dr: '1,100', cr: '' },
      { account: 'Purchases Ledger Control Account', dr: '',      cr: '1,100' },
    ]),

    h3('Worked Example 2.3 — Bank Payments: Keeping Both Tracks in Step'),

    ...makeCashBookExtract('Cash Book (Bank Account) — Payments', [
      { date: '14 May', name: 'T. Mills Ltd',   amount: '600' },
      { date: '16 May', name: 'B. Rogers & Co', amount: '200' },
    ], { col2Label: 'Supplier' }),

    ...makeTAccount('Purchases Ledger Control Account — After Payments', [
      { desc: '16 May  Bank', amount: '800' },
      { desc: '16 May  Balance c/d', amount: '300' },
      { isTotal: true, amount: '£1,100' },
      {},
    ], [
      { desc: 'Bal b/f', amount: '1,100' },
      {},
      { isTotal: true, amount: '£1,100' },
      { desc: 'Bal b/d', amount: '£300', bold: true },
    ]),

    ...makeTAccount('T. Mills Ltd — Purchases Ledger', [
      { desc: '14 May  Bank', amount: '600' },
      { desc: 'Balance c/d', amount: '150' },
      { isTotal: true, amount: '£750' },
      {},
    ], [
      { desc: 'Bal b/f', amount: '750' },
      {},
      { isTotal: true, amount: '£750' },
      { desc: 'Bal b/d', amount: '£150', bold: true },
    ]),

    ...makeTAccount('B. Rogers & Co — Purchases Ledger', [
      { desc: '16 May  Bank', amount: '200' },
      { desc: 'Balance c/d', amount: '150' },
      { isTotal: true, amount: '£350' },
      {},
    ], [
      { desc: 'Bal b/f', amount: '350' },
      {},
      { isTotal: true, amount: '£350' },
      { desc: 'Bal b/d', amount: '£150', bold: true },
    ]),

    bodyPara([run('PLCA £300. Purchases Ledger: £150 + £150 = £300. Agreed. ✓')]),

    ...makeNoteBox('Key Principle', 'The Purchases Ledger must be updated every time an entry is made in the Purchases Ledger Control Account. The PLCA balance and the total of the individual Purchases Ledger balances must always agree.'),

    h3('Practice Questions 2.4–2.6'),
    bodyPara([run('Company: Fernwood Supplies')]),

    ...makeDayBookTable('Purchases Day Book — June 2026', [
      { date: '2 Jun',  name: 'Apex Components',     invoice: 'PDB/601', amount: '840' },
      { date: '6 Jun',  name: 'Clearwater Supplies',  invoice: 'PDB/602', amount: '290' },
      { date: '10 Jun', name: 'Apex Components',     invoice: 'PDB/603', amount: '370' },
    ], { blankTotal: true }),

    h3('Question 2.4'),
    bodyPara([run('(a) Total the Purchases Day Book.')]),
    bodyPara([run('(b) Post to individual supplier accounts and balance both.')]),
    bodyPara([run('(c) Prepare a summarised Trial Balance as at 10 June.')]),

    h3('Question 2.5'),
    bodyPara([run('(a) Prepare a Purchases Ledger Control Account.')]),
    bodyPara([run('(b) Prepare a more summarised Trial Balance.')]),

    ...makeCashBookExtract('Cash Book (Bank Account) — Payments', [
      { date: '18 Jun', name: 'Apex Components',    amount: '840' },
      { date: '20 Jun', name: 'Clearwater Supplies', amount: '290' },
    ], { col2Label: 'Supplier' }),

    h3('Question 2.6'),
    bodyPara([run('(a) Update the PLCA. What is the closing balance?')]),
    bodyPara([run('(b) Update both individual accounts in the Purchases Ledger.')]),
    bodyPara([run('(c) Confirm PLCA agrees with the sum of the Purchases Ledger balances.')]),
    bodyPara([run('(d) What would the discrepancy have been if the Purchases Ledger had not been updated?')]),

    h2('Answers — 2.4–2.6'),
    h3('Answer 2.4'),
    bodyPara([run('Day Book total: £1,500')]),

    ...makeTAccount('Apex Components — Purchases Ledger', [
      { desc: 'Balance c/d', amount: '1,210' },
      { isTotal: true, amount: '£1,210' },
      {},
    ], [
      { desc: '2 Jun  Purchases PDB/601', amount: '840' },
      { desc: '10 Jun  Purchases PDB/603', amount: '370' },
      { isTotal: true, amount: '£1,210' },
      { desc: 'Bal b/d', amount: '£1,210', bold: true },
    ]),

    ...makeTAccount('Clearwater Supplies — Purchases Ledger', [
      { desc: 'Balance c/d', amount: '290' },
      { isTotal: true, amount: '£290' },
      {},
    ], [
      { desc: '6 Jun  Purchases PDB/602', amount: '290' },
      { isTotal: true, amount: '£290' },
      { desc: 'Bal b/d', amount: '£290', bold: true },
    ]),

    ...makeTrialBalance('Trial Balance — 10 June 2026', [
      { account: 'Purchases',            dr: '1,500', cr: '' },
      { account: 'Apex Components',      dr: '',      cr: '1,210' },
      { account: 'Clearwater Supplies',  dr: '',      cr: '290' },
    ]),

    h3('Answer 2.5'),

    ...makeTAccount('Purchases Ledger Control Account', [
      { desc: '10 Jun  Balance c/d', amount: '1,500' },
      { isTotal: true, amount: '£1,500' },
      {},
    ], [
      { desc: '10 Jun  Credit Purchases', amount: '1,500' },
      { isTotal: true, amount: '£1,500' },
      { desc: 'Bal b/d', amount: '£1,500', bold: true },
    ]),

    ...makeTrialBalance('Trial Balance — 10 June 2026', [
      { account: 'Purchases',                        dr: '1,500', cr: '' },
      { account: 'Purchases Ledger Control Account', dr: '',      cr: '1,500' },
    ]),

    h3('Answer 2.6'),

    ...makeTAccount('Purchases Ledger Control Account', [
      { desc: '20 Jun  Bank', amount: '1,130' },
      { desc: '20 Jun  Balance c/d', amount: '370' },
      { isTotal: true, amount: '£1,500' },
      {},
    ], [
      { desc: 'Bal b/f', amount: '1,500' },
      {},
      { isTotal: true, amount: '£1,500' },
      { desc: 'Bal b/d', amount: '£370', bold: true },
    ]),

    ...makeTAccount('Apex Components — Purchases Ledger', [
      { desc: '18 Jun  Bank', amount: '840' },
      { desc: 'Balance c/d', amount: '370' },
      { isTotal: true, amount: '£1,210' },
      {},
    ], [
      { desc: 'Bal b/f', amount: '1,210' },
      {},
      { isTotal: true, amount: '£1,210' },
      { desc: 'Bal b/d', amount: '£370', bold: true },
    ]),

    ...makeTAccount('Clearwater Supplies — Purchases Ledger', [
      { desc: '20 Jun  Bank', amount: '290' },
      { isTotal: true, amount: '£290' },
      {},
    ], [
      { desc: 'Bal b/f', amount: '290' },
      { isTotal: true, amount: '£290' },
      { desc: '(Nil balance)', amount: '', bold: false },
    ]),

    bodyPara([run('PLCA £370. Purchases Ledger: £370 + £0 = £370. Agreed. ✓')]),
    bodyPara([run('Without updating: PLCA £370 vs individual total £1,500. Discrepancy: £1,130.')]),
  );

  // ══════════════════════════════════════════════════════════════════
  // PART 3 — FURTHER ENTRIES
  // ══════════════════════════════════════════════════════════════════
  add(
    pb(),
    h1('Part 3 — Further Entries in the Control Accounts'),
    bodyPara([run('The examples so far have featured only credit sales/purchases and cash receipts/payments. In practice, several other entry types arise. These are introduced here gradually, starting with the two that most often catch students out.')]),

    h2('Sales Returns'),
    bodyPara([run('When a customer returns goods, a credit note is issued. This reduces the amount owed. In the SLCA: Credit entry. In the customer’s individual Receivables Ledger account: Debit entry.')]),

    h2('Irrecoverable Debts (Bad Debts Written Off)'),
    bodyPara([run('When a debt cannot be collected, it is written off. In the SLCA: Credit entry (asset removed). Debit goes to the Irrecoverable Debts Expense account.')]),

    h3('Worked Example 3.1 — Sales Returns and Irrecoverable Debts'),
    bodyPara([run('Company: Hillcrest Ltd. Opening SLCA balance: £2,400 Dr.')]),

    ...makeTransactionSummary('Transaction Summary', [
      { transaction: 'Credit sales for the quarter',               source: 'Sales Day Book',  amount: '3,600' },
      { transaction: 'Sales returns (credit notes issued)',         source: 'Returns Journal', amount: '280' },
      { transaction: 'Irrecoverable debt written off (J. Briggs)', source: 'General Journal', amount: '150' },
      { transaction: 'Bank receipts from credit customers',         source: 'Cash Book',       amount: '1,800' },
    ]),

    ...makeTAccount('Sales Ledger Control Account', [
      { desc: 'Bal b/f', amount: '£2,400' },
      { desc: 'Credit Sales', amount: '3,600' },
      { isTotal: true, amount: '£6,000' },
      { desc: 'Bal b/d', amount: '£3,770', bold: true },
    ], [
      { desc: 'Sales Returns', amount: '280' },
      { desc: 'Irrecoverable Debts', amount: '150' },
      { desc: 'Bank', amount: '1,800' },
      { desc: 'Balance c/d', amount: '3,770' },
      { isTotal: true, amount: '£6,000' },
      {},
    ]),

    h2('Dishonoured Cheques and Interest Charged'),
    bodyPara([run('Dishonoured cheques: If a customer’s cheque is returned unpaid by the bank, the debt is reinstated. Debit the SLCA (the asset is restored). Credit the Cash Book.')]),
    bodyPara([run('Interest charged: If interest is added to overdue accounts, the amount owed increases. Debit the SLCA.')]),

    h3('Worked Example 3.2 — Dishonoured Cheques and Interest Charged'),
    bodyPara([run('Hillcrest Ltd continues. Opening SLCA balance: £3,770 Dr.')]),

    ...makeTransactionSummary('Transaction Summary', [
      { transaction: 'Credit sales for the quarter',                    source: 'Sales Day Book',  amount: '4,200' },
      { transaction: 'Dishonoured cheque — Customer K reinstated', source: 'Cash Book',       amount: '320' },
      { transaction: 'Interest charged to overdue accounts',            source: 'General Journal', amount: '90' },
      { transaction: 'Bank receipts from credit customers',              source: 'Cash Book',       amount: '3,100' },
      { transaction: 'Sales returns',                                    source: 'Returns Journal', amount: '450' },
    ]),

    ...makeTAccount('Sales Ledger Control Account', [
      { desc: 'Bal b/f', amount: '£3,770' },
      { desc: 'Credit Sales', amount: '4,200' },
      { desc: 'Dishonoured Cheque', amount: '320' },
      { desc: 'Interest Charged', amount: '90' },
      { isTotal: true, amount: '£8,380' },
      { desc: 'Bal b/d', amount: '£4,830', bold: true },
    ], [
      { desc: 'Bank', amount: '3,100' },
      { desc: 'Sales Returns', amount: '450' },
      {},
      { desc: 'Balance c/d', amount: '4,830' },
      { isTotal: true, amount: '£8,380' },
      {},
    ]),

    h2('Contras (Set-off Entries)'),
    bodyPara([run('A contra arises when the same entity is both a credit customer (Sales Ledger) and a credit supplier (Purchases Ledger). The mutual debts are offset: Credit the SLCA; Debit the PLCA. No cash changes hands.')]),

    ...makeNoteBox('Exam Trap — Cash Sales', 'Questions frequently give a total sales figure that includes both cash and credit sales. Only credit sales are posted to the SLCA. Cash sales bypass the control account entirely. Always extract the credit figure before posting.'),

    h3('Worked Example 3.3 — Comprehensive SLCA (All Entry Types)'),
    bodyPara([run('Company: Horizon Wholesale Supplies. Opening SLCA balance: £8,200 Dr.')]),

    ...makeTransactionSummary('Transaction Summary', [
      { transaction: 'Total sales (including £1,500 cash sales — exclude)', source: 'Day Book / Cash Book', amount: '24,500' },
      { transaction: 'Bank receipts from credit customers',                           source: 'Cash Book',           amount: '19,400' },
      { transaction: 'Sales returns',                                                 source: 'Returns Journal',     amount: '680' },
      { transaction: 'Interest charged to overdue accounts',                          source: 'General Journal',     amount: '120' },
      { transaction: 'Dishonoured cheque reinstated',                                 source: 'Cash Book',           amount: '450' },
      { transaction: 'Irrecoverable debt written off',                                source: 'General Journal',     amount: '310' },
      { transaction: 'Contra — set-off against Purchases Ledger balance',        source: 'General Journal',     amount: '400' },
    ]),

    bodyPara([run('Credit sales only: £24,500 − £1,500 = £23,000.')]),

    ...makeTAccount('Sales Ledger Control Account', [
      { desc: 'Bal b/f', amount: '£8,200' },
      { desc: 'Credit Sales', amount: '23,000' },
      { desc: 'Interest Charged', amount: '120' },
      { desc: 'Dishonoured Cheque', amount: '450' },
      { isTotal: true, amount: '£31,770' },
      { desc: 'Bal b/d', amount: '£10,980', bold: true },
    ], [
      { desc: 'Bank', amount: '19,400' },
      { desc: 'Sales Returns', amount: '680' },
      { desc: 'Irrecoverable Debts', amount: '310' },
      { desc: 'Contra / Set-off', amount: '400' },
      { desc: 'Balance c/d', amount: '10,980' },
      { isTotal: true, amount: '£31,770' },
      {},
    ]),

    ...makeNoteBox('Exam Trap — Cash Purchases', 'Same principle applies on the purchases side. Exclude cash purchases from the PLCA. Only credit purchases are posted to the Purchases Ledger Control Account.'),

    h3('Worked Example 3.4 — Comprehensive PLCA'),
    bodyPara([run('Company: Summit Components. Opening PLCA balances: £11,600 Cr (normal); £250 Dr (minority — historical overpayment to one supplier).')]),

    ...makeTransactionSummary('Transaction Summary', [
      { transaction: 'Total purchases (including £3,100 cash — exclude)', source: 'Day Book / Cash Book', amount: '28,400' },
      { transaction: 'Bank payments to credit suppliers',                           source: 'Cash Book',           amount: '19,800' },
      { transaction: 'Purchases returns',                                           source: 'Returns Journal',     amount: '620' },
      { transaction: 'Discounts received from suppliers',                           source: 'Cash Book',           amount: '180' },
      { transaction: 'Interest charged to us by suppliers',                         source: 'General Journal',     amount: '90' },
      { transaction: 'Contra — set-off against Sales Ledger',                  source: 'General Journal',     amount: '670' },
    ]),

    bodyPara([run('Credit purchases only: £28,400 − £3,100 = £25,300. Minority debit balance of £140 remains at period end.')]),

    ...makeTAccount('Purchases Ledger Control Account', [
      { desc: 'Bal b/f (Minority)', amount: '£250' },
      { desc: 'Bank', amount: '19,800' },
      { desc: 'Purchases Returns', amount: '620' },
      { desc: 'Discounts Received', amount: '180' },
      { desc: 'Contra', amount: '670' },
      { desc: 'Bal c/d (Normal)', amount: '15,610' },
      { isTotal: true, amount: '£37,130' },
      { desc: 'Bal b/d (Minority)', amount: '£140', bold: true },
    ], [
      { desc: 'Bal b/f (Normal)', amount: '£11,600' },
      { desc: 'Credit Purchases', amount: '25,300' },
      { desc: 'Interest Charged', amount: '90' },
      { desc: 'Bal c/d (Minority)', amount: '140' },
      {},
      {},
      { isTotal: true, amount: '£37,130' },
      { desc: 'Bal b/d (Normal)', amount: '£15,610', bold: true },
    ]),
  );

  // ══════════════════════════════════════════════════════════════════
  // PART 4 — DISTRACTIONS
  // ══════════════════════════════════════════════════════════════════
  add(
    pb(),
    h1('Part 4 — Things That Can Cause Confusion'),
    bodyPara([run('Two features of control accounts often trip students up mid-question. They are worth examining directly before moving on to errors.')]),

    h2('4.1 Simultaneous Debit and Credit Balances'),
    bodyPara([run('The SLCA normally carries a debit balance (asset). Occasionally, a small number of customers may carry a credit balance — typically because they have overpaid or returned goods after their invoice was fully settled. The SLCA must carry both balances simultaneously. They must never be netted off — both are brought down separately.')]),

    h3('Question 4.1 — SLCA with a Minority Credit Balance'),
    bodyPara([run('Company: Peak Manufacturing. Opening SLCA balances: £14,200 Dr (normal); £310 Cr (minority).')]),

    ...makeTransactionSummary('Transaction Summary', [
      { transaction: 'Total sales (including £2,400 cash sales)', source: 'Day Book',       amount: '31,650' },
      { transaction: 'Bank receipts from credit customers',            source: 'Cash Book',       amount: '24,100' },
      { transaction: 'Sales returns',                                  source: 'Returns Journal', amount: '850' },
      { transaction: 'Dishonoured cheque reinstated',                  source: 'Cash Book',       amount: '520' },
      { transaction: 'Irrecoverable debts written off',                source: 'General Journal', amount: '440' },
      { transaction: 'Interest charged to overdue accounts',           source: 'General Journal', amount: '160' },
      { transaction: 'Contra — set-off against Purchases Ledger', source: 'General Journal', amount: '670' },
    ]),

    bodyPara([run('At 30 November, a minority credit balance of £190 still exists.')]),
    bodyPara([run('Required: Prepare the SLCA for November showing both closing balances.')]),

    h3('Answer — 4.1'),
    bodyPara([run('Credit sales: £31,650 − £2,400 = £29,250.')]),

    ...makeTAccount('Sales Ledger Control Account — November', [
      { desc: 'Nov 1  Bal b/f (Normal)', amount: '£14,200' },
      { desc: 'Credit Sales', amount: '29,250' },
      { desc: 'Interest Charged', amount: '160' },
      { desc: 'Dishonoured Cheque', amount: '520' },
      { desc: 'Bal c/d (Minority)', amount: '190' },
      { isTotal: true, amount: '£44,320' },
      { desc: 'Dec 1  Bal b/d (Normal)', amount: '£17,950', bold: true },
      {},
    ], [
      { desc: 'Nov 1  Bal b/f (Minority)', amount: '£310' },
      { desc: 'Bank', amount: '24,100' },
      { desc: 'Sales Returns', amount: '850' },
      { desc: 'Irrecoverable Debts', amount: '440' },
      { desc: 'Contra', amount: '670' },
      { desc: 'Bal c/d (Normal)', amount: '17,950' },
      { isTotal: true, amount: '£44,320' },
      { desc: 'Dec 1  Bal b/d (Minority)', amount: '£190', bold: true },
    ]),

    bodyPara([run('Verification: Dr side: 14,200 + 29,250 + 160 + 520 + 190 = 44,320 ✓ | Cr side: 310 + 24,100 + 850 + 440 + 670 + 17,950 = 44,320 ✓')]),

    h2('4.2 Simultaneous Balances in the PLCA'),
    bodyPara([run('The PLCA normally carries a credit balance (liability). The equivalent situation is a minority debit balance — a small number of suppliers owe the business money, typically due to overpayment. Both balances are brought down separately.')]),

    h3('Question 4.2 — PLCA with a Minority Debit Balance'),
    bodyPara([run('Company: Summit Trading. Opening PLCA balances: £6,800 Cr (normal); £200 Dr (minority).')]),

    ...makeTransactionSummary('Transaction Summary', [
      { transaction: 'Credit purchases',                  source: 'Purchases Day Book', amount: '9,400' },
      { transaction: 'Bank payments to credit suppliers', source: 'Cash Book',          amount: '7,600' },
      { transaction: 'Purchases returns',                 source: 'Returns Journal',    amount: '380' },
      { transaction: 'Discounts received',                source: 'Cash Book',          amount: '120' },
    ]),

    bodyPara([run('At 31 October, a minority debit balance of £90 remains.')]),
    bodyPara([run('Required: Prepare the PLCA for October showing both closing balances.')]),

    h3('Answer — 4.2'),

    ...makeTAccount('Purchases Ledger Control Account — October', [
      { desc: 'Oct 1  Bal b/f (Minority)', amount: '£200' },
      { desc: 'Bank', amount: '7,600' },
      { desc: 'Purchases Returns', amount: '380' },
      { desc: 'Discounts Received', amount: '120' },
      { desc: 'Bal c/d (Normal)', amount: '7,990' },
      { isTotal: true, amount: '£16,290' },
      { desc: 'Nov 1  Bal b/d (Minority)', amount: '£90', bold: true },
    ], [
      { desc: 'Oct 1  Bal b/f (Normal)', amount: '£6,800' },
      { desc: 'Credit Purchases', amount: '9,400' },
      {},
      {},
      { desc: 'Bal c/d (Minority)', amount: '90' },
      { isTotal: true, amount: '£16,290' },
      { desc: 'Nov 1  Bal b/d (Normal)', amount: '£7,990', bold: true },
    ]),

    bodyPara([run('Verification: Dr: 200 + 7,600 + 380 + 120 + 7,990 = 16,290 ✓ | Cr: 6,800 + 9,400 + 90 = 16,290 ✓')]),
  );

  // ══════════════════════════════════════════════════════════════════
  // PART 5 — ERRORS
  // ══════════════════════════════════════════════════════════════════
  add(
    pb(),
    h1('Part 5 — Understanding Errors'),
    bodyPara([run('When there is a discrepancy between the SLCA balance and the total of the individual Receivables Ledger balances, an error has been made. Different error types corrupt different tracks. Rather than being told which is which, the approach here is to make each type of error deliberately and observe the result first-hand.')]),

    h2('Opening Position — Shared by All Questions in Part 5'),

    makeLabelValueTable([
      { label: 'Sales Ledger Control Account (Track 1)', value: '£2,400 Dr' },
      { label: 'Patel & Co — Receivables Ledger',  value: '£1,400' },
      { label: 'Okafor Ltd — Receivables Ledger',  value: '£1,000' },
      { label: 'Total of individual accounts',           value: '£2,400', bold: true },
    ]),

    bodyPara([run('')]),

    ...makeTransactionSummary('Transactions During March 2026', [
      { transaction: 'Credit sales — Patel & Co (Invoice SDB/301)',  source: '', amount: '360' },
      { transaction: 'Credit sales — Okafor Ltd (Invoice SDB/302)',  source: '', amount: '270' },
      { transaction: 'Sales Day Book total',                               source: '', amount: '630' },
      { transaction: 'Bank receipt — Patel & Co',                    source: '', amount: '800' },
    ]),

    h3('Question 5.1 — Making an Omission Error'),
    bodyPara([run('Step 1: Do it correctly. Complete the SLCA and both Receivables Ledger accounts accurately. Confirm the SLCA balance agrees with the total of the individual accounts.')]),
    bodyPara([run('Step 2: Make the error. Suppose the SLCA is updated correctly (using the day book total of £630), but the posting of Okafor Ltd’s invoice SDB/302 (£270) is omitted from their personal account. What is now the discrepancy? Which track is wrong?')]),

    h3('Answer — 5.1'),
    h3('Step 1 — Correct Position'),

    ...makeTAccount('Sales Ledger Control Account', [
      { desc: 'Bal b/f', amount: '£2,400' },
      { desc: 'Credit Sales', amount: '630' },
      { isTotal: true, amount: '£3,030' },
      { desc: 'Bal b/d', amount: '£2,230', bold: true },
    ], [
      { desc: 'Bank (Patel)', amount: '800' },
      { desc: 'Balance c/d', amount: '2,230' },
      { isTotal: true, amount: '£3,030' },
      {},
    ]),

    ...makeTAccount('Patel & Co — Receivables Ledger', [
      { desc: 'Bal b/f', amount: '£1,400' },
      { desc: 'Sales SDB/301', amount: '360' },
      { isTotal: true, amount: '£1,760' },
      { desc: 'Bal b/d', amount: '£960', bold: true },
    ], [
      { desc: 'Bank', amount: '800' },
      { desc: 'Balance c/d', amount: '960' },
      { isTotal: true, amount: '£1,760' },
      {},
    ]),

    ...makeTAccount('Okafor Ltd — Receivables Ledger', [
      { desc: 'Bal b/f', amount: '£1,000' },
      { desc: 'Sales SDB/302', amount: '270' },
      { isTotal: true, amount: '£1,270' },
      { desc: 'Bal b/d', amount: '£1,270', bold: true },
    ], [
      {},
      { desc: 'Balance c/d', amount: '1,270' },
      { isTotal: true, amount: '£1,270' },
      {},
    ]),

    bodyPara([run('SLCA: £2,230. Individual total: £960 + £1,270 = £2,230. Agreed. ✓')]),

    h3('Step 2 — With the Omission (SDB/302 not posted to Okafor’s account)'),

    ...makeTAccount('Okafor Ltd — Receivables Ledger (Omission)', [
      { desc: 'Bal b/f', amount: '£1,000' },
      { isTotal: true, amount: '£1,000' },
      { desc: 'Bal b/d', amount: '£1,000', bold: true },
    ], [
      { desc: 'Balance c/d', amount: '1,000' },
      { isTotal: true, amount: '£1,000' },
      {},
    ]),

    bodyPara([run('SLCA: £2,230 (correct — day book total was used).')]),
    bodyPara([run('Individual total: £960 + £1,000 = £1,960.')]),
    bodyPara([run('Discrepancy: £2,230 − £1,960 = £270 — the value of the missing invoice exactly.')]),
    bodyPara([run('Conclusion: Track 1 (SLCA) is correct. The error is in Track 2 (Receivables Ledger) only.')]),

    h3('Question 5.2 — Making a Casting Error'),
    bodyPara([run('A casting error occurs when the day book total is calculated incorrectly. Because the SLCA is posted from the day book total, the error enters Track 1. Individual accounts are posted from the actual invoices, so Track 2 remains correct.')]),
    bodyPara([run('Using the same March position and transactions:')]),

    bodyPara([run('Part A — Under-cast: Suppose the day book total is calculated as £530 (£100 too low). The SLCA is updated using £530; individual accounts are posted correctly from the actual invoices. What is the resulting discrepancy?')]),

    h3('Answer — Part A'),
    bodyPara([run('SLCA: £2,400 + £530 − £800 = £2,130.')]),
    bodyPara([run('Individual accounts (correct): £960 + £1,270 = £2,230.')]),
    bodyPara([run('Discrepancy: £2,230 − £2,130 = £100 — exactly the under-cast amount.')]),
    bodyPara([run('Conclusion: Track 2 is correct. Track 1 (SLCA) is understated by £100.')]),

    bodyPara([run('Part B — Over-cast: Suppose the day book total is calculated as £780 (£150 too high). What is the discrepancy?')]),

    h3('Answer — Part B'),
    bodyPara([run('SLCA: £2,400 + £780 − £800 = £2,380.')]),
    bodyPara([run('Individual accounts: £2,230.')]),
    bodyPara([run('Discrepancy: £2,380 − £2,230 = £150 — exactly the over-cast amount.')]),
    bodyPara([run('Conclusion: Track 2 is correct. Track 1 (SLCA) is overstated by £150.')]),

    h3('Question 5.3 — Making a Transposition Error'),
    bodyPara([run('A transposition error occurs when two digits are accidentally swapped. This happens at the point of posting to an individual account, so it corrupts Track 2 only. The SLCA, posted from the day book total, is unaffected.')]),
    bodyPara([run('The SLCA is updated correctly. When posting Okafor Ltd’s invoice of £270, the digits are transposed and £720 is entered in their personal account. Calculate both balances and state the discrepancy.')]),

    h3('Answer — 5.3'),
    bodyPara([run('SLCA: £2,230 (correct, as in Step 1 of Q5.1).')]),
    bodyPara([run('Patel & Co: £960 (correct).')]),
    bodyPara([run('Okafor Ltd: £1,000 + £720 = £1,720 (should be £1,270).')]),
    bodyPara([run('Individual total: £960 + £1,720 = £2,680.')]),
    bodyPara([run('Discrepancy: £2,680 − £2,230 = £450. Note: £720 − £270 = £450 — always the difference between the transposed figures.')]),
    bodyPara([run('Conclusion: Track 1 (SLCA) is correct. Track 2 (individual accounts) is overstated by £450.')]),

    h3('Question 5.4 — Consolidation: Identifying and Correcting Mixed Errors'),
    bodyPara([run('After a further period, the following position is reported:')]),

    makeLabelValueTable([
      { label: 'Sales Ledger Control Account (Track 1)',                    value: '£5,200' },
      { label: 'Total of individual Receivables Ledger balances (Track 2)', value: '£5,080' },
      { label: 'Discrepancy',                                               value: '£120', bold: true },
    ]),

    bodyPara([run('')]),
    bodyPara([run('Investigation reveals the following errors:')]),
    bodyPara([run('1. The Sales Day Book total for the period was over-cast by £60.')]),
    bodyPara([run('2. A credit sale of £100 to D. Hughes was correctly posted to the SLCA via the day book but was entirely omitted from D. Hughes’s personal account in the Receivables Ledger.')]),
    bodyPara([run('3. A bank receipt of £240 from K. Osei was correctly entered in the SLCA but posted as £200 in K. Osei’s personal account (i.e. only £200 was deducted from the balance, not £240).')]),
    bodyPara([run('Required: For each error, state which track is affected and the correction needed. Then calculate the corrected balance for both tracks and confirm they agree.')]),

    h3('Answer — 5.4'),

    ...makeErrorAnalysisTable([
      { num: '1', type: 'Over-cast SDB',               t1: 'Reduce by £60 (SLCA overstated)',                                                          t2: 'No effect' },
      { num: '2', type: 'Omission from personal account', t1: 'No effect',                                                                                   t2: 'Add £100 to D. Hughes’s account (understated)' },
      { num: '3', type: 'Under-posted receipt',          t1: 'No effect',                                                                                    t2: 'K. Osei’s balance is £40 too high — credit account by further £40' },
    ]),

    ...makeReconciliationTable('Reconciliation', [
      { label1: 'Original balance',                            amount1: '5,200', label2: 'Original total',                            amount2: '5,080' },
      { label1: 'Less: Error 1 (over-cast SDB)',               amount1: '−60', label2: 'Add: Error 2 (omission — D. Hughes)', amount2: '+100' },
      { label1: '',                                            amount1: '',       label2: 'Less: Error 3 (under-posted receipt — K. Osei)', amount2: '−40' },
      { label1: 'Corrected SLCA balance',                      amount1: '£5,140', label2: 'Corrected individual accounts total', amount2: '£5,140' },
    ]),

    bodyPara([run('Both tracks agree at £5,140. ✓')]),
    bodyPara([run('Verification: SLCA 5,200 − 60 = 5,140. Individual 5,080 + 100 − 40 = 5,140. ✓')]),
  );

  return children;
}

// ─── Build and write ──────────────────────────────────────────────────────────

if (require.main === module) {
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: FONT, size: 20, color: '000000' } },
      },
    },
    numbering: {
      config: [
        { reference: 'bullets',
          levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
        { reference: 'numbers',
          levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ],
    },
    sections: [{
      properties: {
        page: {
          size:   { width: 11906, height: 16838 },
          margin: { top: 1080, right: 850, bottom: 1080, left: 850 },
        },
      },
      children: buildDocument(),
    }],
  });

  const OUT = '/mnt/user-data/outputs/control_accounts_guide.docx';
  fs.mkdirSync('/mnt/user-data/outputs', { recursive: true });
  Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(OUT, buf);
    console.log('Written:', OUT);
  }).catch(err => { console.error(err); process.exit(1); });
}
