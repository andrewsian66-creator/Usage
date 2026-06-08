// Build script: Control Accounts — Building Understanding from First Principles
// Spec: control_accounts_spec.md
// Output: /mnt/user-data/outputs/control_accounts_guide.docx
//
// Setup phase only — Section 4 helper functions and Section 6 content
// are added in the next pass.

const fs = require('fs');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Header,
  Footer,
  AlignmentType,
  LevelFormat,
  HeadingLevel,
  BorderStyle,
  WidthType,
  ShadingType,
  VerticalAlign,
  PageBreak,
} = require('docx');

// ---------------------------------------------------------------------------
// 1. CONSTANTS (Section 3 — Formatting Rules)
// ---------------------------------------------------------------------------

const FONT = 'Arial';
const BLACK = '000000';

// Page geometry (A4, DXA units)
const PAGE = {
  width: 11906,
  height: 16838,
  margin: { top: 1080, right: 850, bottom: 1080, left: 850 },
};
const CONTENT_WIDTH = 10200; // page width - left margin - right margin

// Border presets
const BORDER_THIN  = { style: BorderStyle.SINGLE, size: 4,  color: BLACK };
const BORDER_THICK = { style: BorderStyle.SINGLE, size: 12, color: BLACK };
const BORDER_DOUBLE_THIN = { style: BorderStyle.DOUBLE, size: 4, color: BLACK };
const BORDER_CALLOUT_LEFT = { style: BorderStyle.SINGLE, size: 8, color: BLACK };
const BORDER_NONE = { style: BorderStyle.NONE, size: 0, color: BLACK };

// Unicode characters used in content
const POUND = '£';
const EMDASH = '—';
const MINUS = '−';
const TICK = '✓';

// ---------------------------------------------------------------------------
// 2. STYLES (Section 3 — Heading styles, body)
// Body 10pt = size 20 half-points
// H1 16pt = 32, H2 13pt = 26, H3 11pt = 22
// ---------------------------------------------------------------------------

const styles = {
  default: {
    document: { run: { font: FONT, size: 20 } },
  },
  paragraphStyles: [
    {
      id: 'Heading1',
      name: 'Heading 1',
      basedOn: 'Normal',
      next: 'Normal',
      quickFormat: true,
      run: { font: FONT, size: 32, bold: true, allCaps: true },
      paragraph: {
        spacing: { before: 320, after: 180 },
        border: { bottom: BORDER_THICK },
        outlineLevel: 0,
      },
    },
    {
      id: 'Heading2',
      name: 'Heading 2',
      basedOn: 'Normal',
      next: 'Normal',
      quickFormat: true,
      run: { font: FONT, size: 26, bold: true },
      paragraph: {
        spacing: { before: 240, after: 120 },
        indent: { left: 180 },
        border: { left: { style: BorderStyle.SINGLE, size: 10, color: BLACK, space: 8 } },
        outlineLevel: 1,
      },
    },
    {
      id: 'Heading3',
      name: 'Heading 3',
      basedOn: 'Normal',
      next: 'Normal',
      quickFormat: true,
      run: { font: FONT, size: 22, bold: true },
      paragraph: {
        spacing: { before: 180, after: 80 },
        outlineLevel: 2,
      },
    },
  ],
};

// ---------------------------------------------------------------------------
// 3. NUMBERING (Section 4.8)
// ---------------------------------------------------------------------------

const numbering = {
  config: [
    {
      reference: 'bullets',
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    },
    {
      reference: 'numbers',
      levels: [{
        level: 0,
        format: LevelFormat.DECIMAL,
        text: '%1.',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    },
  ],
};

// ---------------------------------------------------------------------------
// 4. HELPER FUNCTIONS (Section 4)
// ---------------------------------------------------------------------------

const CELL_MARGINS = { top: 60, bottom: 60, left: 100, right: 100 };

function run(text, opts = {}) {
  return new TextRun({ text: String(text), font: FONT, size: 20, ...opts });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    ...opts,
    children: [run(text, opts.runOpts || {})],
  });
}

function boldTitle(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 120, after: 80 },
    ...opts,
    children: [run(text, { bold: true })],
  });
}

function centredBoldTitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 80 },
    children: [run(text, { bold: true })],
  });
}

// Format an integer (or numeric string) with thousands separators.
function fmt(n) {
  const num = typeof n === 'number' ? n : Number(String(n).replace(/[^0-9.-]/g, ''));
  return Number.isFinite(num) ? num.toLocaleString('en-GB') : String(n);
}

// Simple bordered cell builder for grid tables (Day Book, etc.).
function gridCell({ width, children, bold = false, align = AlignmentType.LEFT, borders = null }) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: CELL_MARGINS,
    borders: borders || {
      top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN,
    },
    children: [new Paragraph({
      alignment: align,
      children: Array.isArray(children) ? children : [run(children, { bold })],
    })],
  });
}

// ---- 4.1 Day Book Table ---------------------------------------------------
// 4 cols: Date | Customer/Supplier | Invoice No. | £
// Widths: 1530 | 4080 | 2040 | 2550 = 10200
// All borders: single black size 4. Header bold. No fill.
// blankTotal: true => empty total row; false => auto-sum.
function dayBookTable({ title, customerLabel = 'Customer', rows, blankTotal = false }) {
  const W = [1530, 4080, 2040, 2550];

  const header = new TableRow({
    tableHeader: true,
    children: [
      gridCell({ width: W[0], children: 'Date',       bold: true }),
      gridCell({ width: W[1], children: customerLabel, bold: true }),
      gridCell({ width: W[2], children: 'Invoice No.', bold: true }),
      gridCell({ width: W[3], children: POUND,        bold: true, align: AlignmentType.RIGHT }),
    ],
  });

  const dataRows = rows.map(r => new TableRow({
    children: [
      gridCell({ width: W[0], children: r.date }),
      gridCell({ width: W[1], children: r.name }),
      gridCell({ width: W[2], children: r.invoice }),
      gridCell({ width: W[3], children: fmt(r.amount), align: AlignmentType.RIGHT }),
    ],
  }));

  const total = blankTotal
    ? ''
    : fmt(rows.reduce((s, r) => s + Number(String(r.amount).replace(/[^0-9.-]/g, '')), 0));

  const totalRow = new TableRow({
    children: [
      gridCell({ width: W[0], children: 'Total', bold: true }),
      gridCell({ width: W[1], children: '' }),
      gridCell({ width: W[2], children: '' }),
      gridCell({ width: W[3], children: total, bold: true, align: AlignmentType.RIGHT }),
    ],
  });

  const table = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: W,
    rows: [header, ...dataRows, totalRow],
  });

  return [boldTitle(title), table, body('')];
}

// ---- 4.2 T-Account --------------------------------------------------------
// 4 cols: Dr desc | Dr amount | Cr desc | Cr amount
// Widths: 3900 | 1200 | 3900 | 1200 = 10200
// Borders: outer NONE. Centre vertical = thick (right of col 2 / left of col 3).
//          Header row bottom = thick. Total row top = thin, bottom = double thin.
//          All other row edges: NONE.
// Row entries (per side):
//   { desc, amount }             standard
//   { isTotal: true, amount }    total row (top thin + bottom double)
//   { bold: true, desc, amount } balance b/d
//   {}                            blank filler
// If drRows.length != crRows.length, pad shorter side with {}.
// Title rendered as bold centred paragraph.
function tAccount({ title, drRows, crRows }) {
  const n = Math.max(drRows.length, crRows.length);
  const dr = [...drRows, ...Array(n - drRows.length).fill({})];
  const cr = [...crRows, ...Array(n - crRows.length).fill({})];

  const W = [3900, 1200, 3900, 1200];

  // Helper to build the four cells of a body row given a Dr entry + Cr entry.
  const buildBodyRow = (d, c) => {
    const isTotal = Boolean(d.isTotal || c.isTotal);
    const top    = isTotal ? BORDER_THIN        : BORDER_NONE;
    const bottom = isTotal ? BORDER_DOUBLE_THIN : BORDER_NONE;

    // Cell content extraction.
    const drDescText   = d.isTotal ? '' : (d.desc   || '');
    const drAmountText = d.amount  || '';
    const crDescText   = c.isTotal ? '' : (c.desc   || '');
    const crAmountText = c.amount  || '';
    const drBold = Boolean(d.bold);
    const crBold = Boolean(c.bold);

    return new TableRow({
      children: [
        // Dr description
        new TableCell({
          width: { size: W[0], type: WidthType.DXA },
          margins: CELL_MARGINS,
          borders: { top, bottom, left: BORDER_NONE, right: BORDER_NONE },
          children: [new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [run(drDescText, { bold: drBold })],
          })],
        }),
        // Dr amount — right border is the thick centre line
        new TableCell({
          width: { size: W[1], type: WidthType.DXA },
          margins: CELL_MARGINS,
          borders: { top, bottom, left: BORDER_NONE, right: BORDER_THICK },
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [run(drAmountText, { bold: drBold || isTotal })],
          })],
        }),
        // Cr description — left border is the thick centre line
        new TableCell({
          width: { size: W[2], type: WidthType.DXA },
          margins: CELL_MARGINS,
          borders: { top, bottom, left: BORDER_THICK, right: BORDER_NONE },
          children: [new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [run(crDescText, { bold: crBold })],
          })],
        }),
        // Cr amount
        new TableCell({
          width: { size: W[3], type: WidthType.DXA },
          margins: CELL_MARGINS,
          borders: { top, bottom, left: BORDER_NONE, right: BORDER_NONE },
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [run(crAmountText, { bold: crBold || isTotal })],
          })],
        }),
      ],
    });
  };

  // Header row: "Dr" spans cols 1-2, "Cr" spans cols 3-4. Thick bottom border.
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      new TableCell({
        columnSpan: 2,
        width: { size: W[0] + W[1], type: WidthType.DXA },
        margins: CELL_MARGINS,
        borders: {
          top: BORDER_NONE, bottom: BORDER_THICK,
          left: BORDER_NONE, right: BORDER_THICK,
        },
        children: [new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [run('Dr', { bold: true })],
        })],
      }),
      new TableCell({
        columnSpan: 2,
        width: { size: W[2] + W[3], type: WidthType.DXA },
        margins: CELL_MARGINS,
        borders: {
          top: BORDER_NONE, bottom: BORDER_THICK,
          left: BORDER_THICK, right: BORDER_NONE,
        },
        children: [new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [run('Cr', { bold: true })],
        })],
      }),
    ],
  });

  const bodyRows = dr.map((d, i) => buildBodyRow(d, cr[i]));

  const table = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: W,
    rows: [headerRow, ...bodyRows],
  });

  return [centredBoldTitle(title), table, body('')];
}

// ---- 4.3 Trial Balance ----------------------------------------------------
// 3 cols: Account | Dr £ | Cr £
// Widths: 6300 | 1950 | 1950 = 10200. Thin borders. Header bold.
// Auto-calculates totals row from rows. Amount cols right-aligned.
// rows: [{ account, dr?, cr? }, ...]  — pass either dr or cr per row.
function trialBalance({ title, rows }) {
  const W = [6300, 1950, 1950];

  const header = new TableRow({
    tableHeader: true,
    children: [
      gridCell({ width: W[0], children: 'Account', bold: true }),
      gridCell({ width: W[1], children: 'Dr £',    bold: true, align: AlignmentType.RIGHT }),
      gridCell({ width: W[2], children: 'Cr £',    bold: true, align: AlignmentType.RIGHT }),
    ],
  });

  const num = v => Number(String(v ?? '').replace(/[^0-9.-]/g, '')) || 0;

  const dataRows = rows.map(r => new TableRow({
    children: [
      gridCell({ width: W[0], children: r.account }),
      gridCell({ width: W[1], children: r.dr != null && r.dr !== '' ? fmt(r.dr) : '', align: AlignmentType.RIGHT }),
      gridCell({ width: W[2], children: r.cr != null && r.cr !== '' ? fmt(r.cr) : '', align: AlignmentType.RIGHT }),
    ],
  }));

  const drTotal = rows.reduce((s, r) => s + num(r.dr), 0);
  const crTotal = rows.reduce((s, r) => s + num(r.cr), 0);

  const totalRow = new TableRow({
    children: [
      gridCell({ width: W[0], children: 'Total',    bold: true }),
      gridCell({ width: W[1], children: fmt(drTotal), bold: true, align: AlignmentType.RIGHT }),
      gridCell({ width: W[2], children: fmt(crTotal), bold: true, align: AlignmentType.RIGHT }),
    ],
  });

  const table = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: W,
    rows: [header, ...dataRows, totalRow],
  });

  return [boldTitle(title), table, body('')];
}

// ---- 4.4 Cash Book Extract ------------------------------------------------
// 3 cols: Date | Customer/Supplier | £
// Widths: 1530 | 6120 | 2550 = 10200. Thin borders. Header bold.
// mode: 'receipts' => column 2 header "Customer"; 'payments' => "Supplier".
// Auto-totals.
function cashBookExtract({ title, mode = 'receipts', rows }) {
  const W = [1530, 6120, 2550];
  const partyLabel = mode === 'payments' ? 'Supplier' : 'Customer';

  const header = new TableRow({
    tableHeader: true,
    children: [
      gridCell({ width: W[0], children: 'Date',     bold: true }),
      gridCell({ width: W[1], children: partyLabel, bold: true }),
      gridCell({ width: W[2], children: POUND,      bold: true, align: AlignmentType.RIGHT }),
    ],
  });

  const dataRows = rows.map(r => new TableRow({
    children: [
      gridCell({ width: W[0], children: r.date }),
      gridCell({ width: W[1], children: r.name }),
      gridCell({ width: W[2], children: fmt(r.amount), align: AlignmentType.RIGHT }),
    ],
  }));

  const total = rows.reduce((s, r) => s + Number(String(r.amount).replace(/[^0-9.-]/g, '')), 0);

  const totalRow = new TableRow({
    children: [
      gridCell({ width: W[0], children: 'Total',  bold: true }),
      gridCell({ width: W[1], children: '' }),
      gridCell({ width: W[2], children: fmt(total), bold: true, align: AlignmentType.RIGHT }),
    ],
  });

  const table = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: W,
    rows: [header, ...dataRows, totalRow],
  });

  return [boldTitle(title), table, body('')];
}

// ---- 4.5 Transaction Summary Table ----------------------------------------
// 3 cols: Transaction | Source | £
// Widths: 7140 | 1530 | 1530 = 10200. Thin borders. Header bold. No total row.
function transactionSummary({ title, rows }) {
  const W = [7140, 1530, 1530];

  const header = new TableRow({
    tableHeader: true,
    children: [
      gridCell({ width: W[0], children: 'Transaction', bold: true }),
      gridCell({ width: W[1], children: 'Source',      bold: true }),
      gridCell({ width: W[2], children: POUND,         bold: true, align: AlignmentType.RIGHT }),
    ],
  });

  const dataRows = rows.map(r => new TableRow({
    children: [
      gridCell({ width: W[0], children: r.transaction }),
      gridCell({ width: W[1], children: r.source }),
      gridCell({ width: W[2], children: fmt(r.amount), align: AlignmentType.RIGHT }),
    ],
  }));

  const table = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: W,
    rows: [header, ...dataRows],
  });

  const out = [table, body('')];
  if (title) out.unshift(boldTitle(title));
  return out;
}

// ---- 4.6 Note Box (Callout) -----------------------------------------------
// Single-cell, full content width.
// Left border only: single size 8. No fill, no other borders.
// Internal margins: top 60, bottom 60, left 200, right 60.
// Contents: bold title paragraph + body paragraph.
// Surrounded by blank spacing paragraphs.
function noteBox({ title, content }) {
  const noteBorders = {
    top:    BORDER_NONE,
    bottom: BORDER_NONE,
    left:   BORDER_CALLOUT_LEFT,
    right:  BORDER_NONE,
  };

  const cell = new TableCell({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 200, right: 60 },
    borders: noteBorders,
    children: [
      new Paragraph({
        spacing: { after: 60 },
        children: [run(title, { bold: true })],
      }),
      new Paragraph({
        children: [run(content)],
      }),
    ],
  });

  const table = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    borders: {
      top:               BORDER_NONE,
      bottom:            BORDER_NONE,
      left:              BORDER_NONE,
      right:             BORDER_NONE,
      insideHorizontal:  BORDER_NONE,
      insideVertical:    BORDER_NONE,
    },
    rows: [new TableRow({ children: [cell] })],
  });

  return [body(''), table, body('')];
}

// ---- 4.7 Reconciliation Table ---------------------------------------------
// 4 cols: Track 1 label | Track 1 £ | Track 2 label | Track 2 £
// Widths: 4050 | 1050 | 4050 | 1050 = 10200. Thin borders. Header bold.
// Last data row: bold (the "agreed/corrected" row). Amount cols right-aligned.
// headers = [track1Header, track2Header]
// rows: [{ l1, l1Amount, l2, l2Amount }, ...]
function reconciliationTable({ title, headers, rows }) {
  const W = [4050, 1050, 4050, 1050];

  const header = new TableRow({
    tableHeader: true,
    children: [
      gridCell({ width: W[0], children: headers[0], bold: true }),
      gridCell({ width: W[1], children: POUND,      bold: true, align: AlignmentType.RIGHT }),
      gridCell({ width: W[2], children: headers[1], bold: true }),
      gridCell({ width: W[3], children: POUND,      bold: true, align: AlignmentType.RIGHT }),
    ],
  });

  const lastIdx = rows.length - 1;
  const dataRows = rows.map((r, i) => {
    const isLast = i === lastIdx;
    return new TableRow({
      children: [
        gridCell({ width: W[0], children: r.l1 || '',       bold: isLast }),
        gridCell({ width: W[1], children: r.l1Amount || '', bold: isLast, align: AlignmentType.RIGHT }),
        gridCell({ width: W[2], children: r.l2 || '',       bold: isLast }),
        gridCell({ width: W[3], children: r.l2Amount || '', bold: isLast, align: AlignmentType.RIGHT }),
      ],
    });
  });

  const table = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: W,
    rows: [header, ...dataRows],
  });

  const out = [table, body('')];
  if (title) out.unshift(boldTitle(title));
  return out;
}

// ---------------------------------------------------------------------------
// 5. EXTRA SMALL HELPERS FOR SECTION 6 CONTENT
// ---------------------------------------------------------------------------

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: FONT, size: 32, bold: true, allCaps: true })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: FONT, size: 26, bold: true })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: FONT, size: 22, bold: true })],
  });
}
function pageBreakPara() {
  return new Paragraph({ children: [new PageBreak()] });
}
// Inline-formatted paragraph: parts = array of strings OR {text, bold?, italics?}
function para(parts, opts = {}) {
  const arr = Array.isArray(parts) ? parts : [parts];
  return new Paragraph({
    spacing: { after: 120 },
    ...opts,
    children: arr.map(p =>
      typeof p === 'string'
        ? run(p)
        : run(p.text, { bold: p.bold, italics: p.italics })
    ),
  });
}
// Sub-item paragraph for (a), (b), (c) — plain text per spec §4.8
function sub(text) {
  return new Paragraph({
    spacing: { after: 80 },
    indent: { left: 360 },
    children: [run(text)],
  });
}
// Generic small table used for the Part-5 opening positions etc.
// headers: array of strings (header row). rows: array of arrays of strings.
// widths: array of DXA widths matching column count.
// boldLastRow: makes the last row bold.
function smallTable({ title, headers, rows, widths, boldLastRow = false, amountColIndex = null }) {
  const headerCells = headers.map((h, i) => gridCell({
    width: widths[i],
    children: h,
    bold: true,
    align: amountColIndex !== null && i >= amountColIndex ? AlignmentType.RIGHT : AlignmentType.LEFT,
  }));
  const headerRow = new TableRow({ tableHeader: true, children: headerCells });

  const lastIdx = rows.length - 1;
  const dataRows = rows.map((r, i) => new TableRow({
    children: r.map((cell, ci) => gridCell({
      width: widths[ci],
      children: cell,
      bold: boldLastRow && i === lastIdx,
      align: amountColIndex !== null && ci >= amountColIndex ? AlignmentType.RIGHT : AlignmentType.LEFT,
    })),
  }));

  const table = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: [headerRow, ...dataRows],
  });

  const out = [table, body('')];
  if (title) out.unshift(boldTitle(title));
  return out;
}

// ---------------------------------------------------------------------------
// 6. CONTENT (Section 6 of the spec)
// ---------------------------------------------------------------------------

function titlePage() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 3600, after: 240 },
      children: [run('Control Accounts', { bold: true, size: 42 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 320 },
      children: [run('Building Understanding from First Principles', { italics: true, size: 24 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [run('A-Level Accounting — Worked Examples and Practice Questions')],
    }),
  ];
}

// ----- PART 1 — CREDIT SALES ----------------------------------------------
function part1() {
  return [
    pageBreakPara(),
    h1('Part 1 — Credit Sales: From Individual Accounts to the Control Account'),

    para('When a business sells goods on credit, each transaction is recorded in the Sales Day Book. The total is posted to the Sales Account (Credit) in the General Ledger, and individual amounts are posted as debits to each customer’s personal account in the Sales Ledger. The worked example below shows this process.'),

    // ---- Worked Example 1.1 -----------------------------------------------
    h3('Worked Example 1.1 — The Sales Day Book and Individual Customer Accounts'),
    para([{ text: 'Company: ', bold: true }, 'Hartley Trading Ltd']),

    ...dayBookTable({
      title: 'Sales Day Book — May 2026',
      rows: [
        { date: '1 May', name: 'A. Brown',  invoice: 'SDB/101', amount: 400 },
        { date: '3 May', name: 'C. Davies', invoice: 'SDB/102', amount: 180 },
        { date: '7 May', name: 'A. Brown',  invoice: 'SDB/103', amount: 250 },
      ],
    }),

    ...tAccount({
      title: 'A. Brown — Sales Ledger',
      drRows: [
        { desc: '1 May  Sales SDB/101', amount: '£400' },
        { desc: '7 May  Sales SDB/103', amount: '250' },
        { isTotal: true, amount: '£650' },
        { bold: true, desc: 'Bal b/d', amount: '£650' },
      ],
      crRows: [
        {},
        { desc: '7 May  Balance c/d', amount: '650' },
        { isTotal: true, amount: '£650' },
        {},
      ],
    }),

    ...tAccount({
      title: 'C. Davies — Sales Ledger',
      drRows: [
        { desc: '3 May  Sales SDB/102', amount: '£180' },
        { isTotal: true, amount: '£180' },
        { bold: true, desc: 'Bal b/d', amount: '£180' },
      ],
      crRows: [
        { desc: '3 May  Balance c/d', amount: '180' },
        { isTotal: true, amount: '£180' },
        {},
      ],
    }),

    ...trialBalance({
      title: 'Summarised Trial Balance — 7 May 2026',
      rows: [
        { account: 'A. Brown',  dr: 650, cr: '' },
        { account: 'C. Davies', dr: 180, cr: '' },
        { account: 'Sales',     dr: '',  cr: 830 },
      ],
    }),

    // ---- Bridge ------------------------------------------------------------
    h2('Introducing the Sales Ledger Control Account'),
    para('As a business grows, the number of credit customers increases. Maintaining a separate General Ledger account for each customer becomes impractical and clutters the Trial Balance. A single Sales Ledger Control Account (SLCA) summarises all credit customer transactions in total, replacing the many individual accounts.'),

    // ---- Worked Example 1.2 -----------------------------------------------
    h3('Worked Example 1.2 — The SLCA Replaces the Individual Accounts'),

    ...tAccount({
      title: 'Sales Ledger Control Account',
      drRows: [
        { desc: '7 May  Credit Sales', amount: '£830' },
        { isTotal: true, amount: '£830' },
        { bold: true, desc: 'Bal b/d', amount: '£830' },
      ],
      crRows: [
        { desc: '7 May  Balance c/d', amount: '830' },
        { isTotal: true, amount: '£830' },
        {},
      ],
    }),

    ...trialBalance({
      title: 'Summarised Trial Balance — 7 May 2026',
      rows: [
        { account: 'Sales Ledger Control Account', dr: 830, cr: '' },
        { account: 'Sales',                        dr: '',  cr: 830 },
      ],
    }),

    // ---- Bridge ------------------------------------------------------------
    h2('The Advantage, the Disadvantage, and the Receivables Ledger'),
    para('The SLCA simplifies the General Ledger and reduces double-entry postings. However, the Trial Balance now shows only the total owed (£830) — it cannot tell us that A. Brown owes £650 and C. Davies owes £180. To recover that detail, businesses maintain a Receivables Ledger — a separate set of individual customer accounts running alongside the General Ledger.'),
    para('Crucially, these accounts are memorandum records only. They are not part of the double-entry system and do not appear in the Trial Balance. The SLCA is the account of record; the Receivables Ledger provides the supporting analysis.'),

    // ---- Worked Example 1.3 -----------------------------------------------
    h3('Worked Example 1.3 — Cash Receipts: Keeping Both Tracks in Step'),

    ...cashBookExtract({
      title: 'Cash Book (Bank Account) — Receipts',
      mode: 'receipts',
      rows: [
        { date: '10 May', name: 'A. Brown',  amount: 400 },
        { date: '12 May', name: 'C. Davies', amount: 180 },
      ],
    }),

    ...tAccount({
      title: 'Sales Ledger Control Account (after receipts)',
      drRows: [
        { desc: 'Bal b/f', amount: '£830' },
        {},
        { isTotal: true, amount: '£830' },
        { bold: true, desc: 'Bal b/d', amount: '£250' },
      ],
      crRows: [
        { desc: '12 May  Bank',         amount: '580' },
        { desc: '12 May  Balance c/d',  amount: '250' },
        { isTotal: true, amount: '£830' },
        {},
      ],
    }),

    ...tAccount({
      title: 'A. Brown — Receivables Ledger',
      drRows: [
        { desc: 'Bal b/f', amount: '£650' },
        {},
        { isTotal: true, amount: '£650' },
        { bold: true, desc: 'Bal b/d', amount: '£250' },
      ],
      crRows: [
        { desc: '10 May  Bank',        amount: '400' },
        { desc: 'Balance c/d',         amount: '250' },
        { isTotal: true, amount: '£650' },
        {},
      ],
    }),

    ...tAccount({
      title: 'C. Davies — Receivables Ledger',
      drRows: [
        { desc: 'Bal b/f', amount: '£180' },
        { isTotal: true, amount: '£180' },
      ],
      crRows: [
        { desc: '12 May  Bank', amount: '180' },
        { isTotal: true, amount: '£180' },
      ],
    }),
    para([{ text: '(Nil balance)', italics: true }]),

    para('SLCA balance: £250. Receivables Ledger total: £250 + £0 = £250. Agreed. ' + TICK),
    para('If the Receivables Ledger had not been updated: SLCA would show £250 while individual accounts would still total £830 — a discrepancy of £580.'),

    ...noteBox({
      title: 'Key Principle',
      content: 'The Receivables Ledger must be updated every time an entry is made in the Sales Ledger Control Account. The SLCA balance and the total of the individual Receivables Ledger balances must always agree.',
    }),

    // ---- Practice Questions 1.4–1.6 ---------------------------------------
    h3('Practice Questions 1.4–1.6'),
    para([{ text: 'Company: ', bold: true }, 'Fernwood Supplies']),

    ...dayBookTable({
      title: 'Sales Day Book — June 2026',
      rows: [
        { date: '1 Jun', name: 'M. Khan',   invoice: 'SDB/201', amount: 520 },
        { date: '4 Jun', name: 'P. Thomas', invoice: 'SDB/202', amount: 340 },
        { date: '9 Jun', name: 'M. Khan',   invoice: 'SDB/203', amount: 180 },
      ],
      blankTotal: true,
    }),

    para([{ text: 'Question 1.4', bold: true }]),
    sub('(a) Total the Sales Day Book.'),
    sub('(b) Post to the individual customer accounts in the Sales Ledger and balance both.'),
    sub('(c) Prepare a summarised Trial Balance as at 9 June.'),

    para([{ text: 'Question 1.5', bold: true }]),
    sub('(a) Prepare a Sales Ledger Control Account for the same transactions.'),
    sub('(b) Prepare a more summarised Trial Balance.'),

    ...cashBookExtract({
      title: 'Cash Book (Bank Account) — Receipts',
      mode: 'receipts',
      rows: [
        { date: '15 Jun', name: 'M. Khan',   amount: 520 },
        { date: '17 Jun', name: 'P. Thomas', amount: 200 },
      ],
    }),

    para([{ text: 'Question 1.6', bold: true }]),
    sub('(a) Update the SLCA to reflect the receipts. What is the closing balance?'),
    sub('(b) Update both individual accounts in the Receivables Ledger.'),
    sub('(c) Confirm that the SLCA balance agrees with the sum of the individual Receivables Ledger balances.'),
    sub('(d) What would the discrepancy have been if the Receivables Ledger had not been updated?'),

    // ---- Answers 1.4–1.6 ---------------------------------------------------
    h3('Answers — 1.4–1.6'),

    para([{ text: 'Answer 1.4:', bold: true }]),
    para('Day Book total: £1,040'),

    ...tAccount({
      title: 'M. Khan — Sales Ledger',
      drRows: [
        { desc: '1 Jun  Sales SDB/201', amount: '£520' },
        { desc: '9 Jun  Sales SDB/203', amount: '180' },
        { isTotal: true, amount: '£700' },
        { bold: true, desc: 'Bal b/d', amount: '£700' },
      ],
      crRows: [
        {},
        { desc: '9 Jun  Balance c/d', amount: '700' },
        { isTotal: true, amount: '£700' },
        {},
      ],
    }),

    ...tAccount({
      title: 'P. Thomas — Sales Ledger',
      drRows: [
        { desc: '4 Jun  Sales SDB/202', amount: '£340' },
        { isTotal: true, amount: '£340' },
        { bold: true, desc: 'Bal b/d', amount: '£340' },
      ],
      crRows: [
        { desc: '4 Jun  Balance c/d', amount: '340' },
        { isTotal: true, amount: '£340' },
        {},
      ],
    }),

    ...trialBalance({
      title: 'Trial Balance — 9 June 2026',
      rows: [
        { account: 'M. Khan',   dr: 700,  cr: '' },
        { account: 'P. Thomas', dr: 340,  cr: '' },
        { account: 'Sales',     dr: '',   cr: 1040 },
      ],
    }),

    para([{ text: 'Answer 1.5:', bold: true }]),

    ...tAccount({
      title: 'Sales Ledger Control Account',
      drRows: [
        { desc: '9 Jun  Credit Sales', amount: '£1,040' },
        { isTotal: true, amount: '£1,040' },
        { bold: true, desc: 'Bal b/d', amount: '£1,040' },
      ],
      crRows: [
        { desc: '9 Jun  Balance c/d', amount: '1,040' },
        { isTotal: true, amount: '£1,040' },
        {},
      ],
    }),

    ...trialBalance({
      title: 'Trial Balance — 9 June 2026',
      rows: [
        { account: 'Sales Ledger Control Account', dr: 1040, cr: '' },
        { account: 'Sales',                        dr: '',   cr: 1040 },
      ],
    }),

    para([{ text: 'Answer 1.6:', bold: true }]),

    ...tAccount({
      title: 'Sales Ledger Control Account',
      drRows: [
        { desc: 'Bal b/f', amount: '£1,040' },
        {},
        { isTotal: true, amount: '£1,040' },
        { bold: true, desc: 'Bal b/d', amount: '£320' },
      ],
      crRows: [
        { desc: '17 Jun  Bank',        amount: '720' },
        { desc: '17 Jun  Balance c/d', amount: '320' },
        { isTotal: true, amount: '£1,040' },
        {},
      ],
    }),

    ...tAccount({
      title: 'M. Khan — Receivables Ledger',
      drRows: [
        { desc: 'Bal b/f', amount: '£700' },
        {},
        { isTotal: true, amount: '£700' },
        { bold: true, desc: 'Bal b/d', amount: '£180' },
      ],
      crRows: [
        { desc: '15 Jun  Bank', amount: '520' },
        { desc: 'Balance c/d',  amount: '180' },
        { isTotal: true, amount: '£700' },
        {},
      ],
    }),

    ...tAccount({
      title: 'P. Thomas — Receivables Ledger',
      drRows: [
        { desc: 'Bal b/f', amount: '£340' },
        {},
        { isTotal: true, amount: '£340' },
        { bold: true, desc: 'Bal b/d', amount: '£140' },
      ],
      crRows: [
        { desc: '17 Jun  Bank', amount: '200' },
        { desc: 'Balance c/d',  amount: '140' },
        { isTotal: true, amount: '£340' },
        {},
      ],
    }),

    para('SLCA £320. Receivables Ledger: £180 + £140 = £320. Agreed. ' + TICK),
    para('Without updating: SLCA £320 vs individual total £1,040. Discrepancy: £720.'),
  ];
}

// ----- PART 2 — CREDIT PURCHASES ------------------------------------------
function part2() {
  return [
    pageBreakPara(),
    h1('Part 2 — Credit Purchases: The Mirror Image'),

    para('The same twin-track logic applies on the purchases side. Credit purchases flow from the Purchases Day Book to individual supplier accounts in the Purchases Ledger. A Purchases Ledger Control Account (PLCA) in the General Ledger replaces the individual supplier accounts there. The key difference is direction: suppliers are creditors, so the normal balances are reversed.'),

    // ---- Worked Example 2.1 -----------------------------------------------
    h3('Worked Example 2.1 — Purchases Day Book and Individual Supplier Accounts'),
    para([{ text: 'Company: ', bold: true }, 'Hartley Trading Ltd']),

    ...dayBookTable({
      title: 'Purchases Day Book — May 2026',
      customerLabel: 'Supplier',
      rows: [
        { date: '2 May', name: 'T. Mills Ltd',   invoice: 'PDB/501', amount: 600 },
        { date: '5 May', name: 'B. Rogers & Co', invoice: 'PDB/502', amount: 350 },
        { date: '8 May', name: 'T. Mills Ltd',   invoice: 'PDB/503', amount: 150 },
      ],
    }),

    ...tAccount({
      title: 'T. Mills Ltd — Purchases Ledger (creditor — normal balance is Cr)',
      drRows: [
        { desc: 'Balance c/d', amount: '750' },
        {},
        { isTotal: true, amount: '£750' },
        {},
      ],
      crRows: [
        { desc: '2 May  Purchases PDB/501', amount: '600' },
        { desc: '8 May  Purchases PDB/503', amount: '150' },
        { isTotal: true, amount: '£750' },
        { bold: true, desc: 'Bal b/d', amount: '£750' },
      ],
    }),

    ...tAccount({
      title: 'B. Rogers & Co — Purchases Ledger',
      drRows: [
        { desc: 'Balance c/d', amount: '350' },
        { isTotal: true, amount: '£350' },
        {},
      ],
      crRows: [
        { desc: '5 May  Purchases PDB/502', amount: '350' },
        { isTotal: true, amount: '£350' },
        { bold: true, desc: 'Bal b/d', amount: '£350' },
      ],
    }),

    ...trialBalance({
      title: 'Summarised Trial Balance — 8 May 2026',
      rows: [
        { account: 'Purchases',       dr: 1100, cr: '' },
        { account: 'T. Mills Ltd',    dr: '',   cr: 750 },
        { account: 'B. Rogers & Co',  dr: '',   cr: 350 },
      ],
    }),

    // ---- Worked Example 2.2 -----------------------------------------------
    h3('Worked Example 2.2 — The PLCA Replaces the Individual Accounts'),

    ...tAccount({
      title: 'Purchases Ledger Control Account',
      drRows: [
        { desc: '8 May  Balance c/d', amount: '1,100' },
        { isTotal: true, amount: '£1,100' },
        {},
      ],
      crRows: [
        { desc: '8 May  Credit Purchases', amount: '1,100' },
        { isTotal: true, amount: '£1,100' },
        { bold: true, desc: 'Bal b/d', amount: '£1,100' },
      ],
    }),

    ...trialBalance({
      title: 'Summarised Trial Balance — 8 May 2026',
      rows: [
        { account: 'Purchases',                          dr: 1100, cr: '' },
        { account: 'Purchases Ledger Control Account',   dr: '',   cr: 1100 },
      ],
    }),

    // ---- Worked Example 2.3 -----------------------------------------------
    h3('Worked Example 2.3 — Bank Payments: Keeping Both Tracks in Step'),

    ...cashBookExtract({
      title: 'Cash Book (Bank Account) — Payments',
      mode: 'payments',
      rows: [
        { date: '14 May', name: 'T. Mills Ltd',   amount: 600 },
        { date: '16 May', name: 'B. Rogers & Co', amount: 200 },
      ],
    }),

    ...tAccount({
      title: 'Purchases Ledger Control Account (after payments)',
      drRows: [
        { desc: '16 May  Bank',         amount: '800' },
        { desc: '16 May  Balance c/d',  amount: '300' },
        { isTotal: true, amount: '£1,100' },
        {},
      ],
      crRows: [
        { desc: 'Bal b/f', amount: '1,100' },
        {},
        { isTotal: true, amount: '£1,100' },
        { bold: true, desc: 'Bal b/d', amount: '£300' },
      ],
    }),

    ...tAccount({
      title: 'T. Mills Ltd — Purchases Ledger',
      drRows: [
        { desc: '14 May  Bank', amount: '600' },
        { desc: 'Balance c/d',  amount: '150' },
        { isTotal: true, amount: '£750' },
        {},
      ],
      crRows: [
        { desc: 'Bal b/f', amount: '750' },
        {},
        { isTotal: true, amount: '£750' },
        { bold: true, desc: 'Bal b/d', amount: '£150' },
      ],
    }),

    ...tAccount({
      title: 'B. Rogers & Co — Purchases Ledger',
      drRows: [
        { desc: '16 May  Bank', amount: '200' },
        { desc: 'Balance c/d',  amount: '150' },
        { isTotal: true, amount: '£350' },
        {},
      ],
      crRows: [
        { desc: 'Bal b/f', amount: '350' },
        {},
        { isTotal: true, amount: '£350' },
        { bold: true, desc: 'Bal b/d', amount: '£150' },
      ],
    }),

    para('PLCA £300. Purchases Ledger: £150 + £150 = £300. Agreed. ' + TICK),

    ...noteBox({
      title: 'Key Principle',
      content: 'The Purchases Ledger must be updated every time an entry is made in the Purchases Ledger Control Account. The PLCA balance and the total of the individual Purchases Ledger balances must always agree.',
    }),

    // ---- Practice Questions 2.4–2.6 ---------------------------------------
    h3('Practice Questions 2.4–2.6'),
    para([{ text: 'Company: ', bold: true }, 'Fernwood Supplies']),

    ...dayBookTable({
      title: 'Purchases Day Book — June 2026',
      customerLabel: 'Supplier',
      rows: [
        { date: '2 Jun',  name: 'Apex Components',     invoice: 'PDB/601', amount: 840 },
        { date: '6 Jun',  name: 'Clearwater Supplies', invoice: 'PDB/602', amount: 290 },
        { date: '10 Jun', name: 'Apex Components',     invoice: 'PDB/603', amount: 370 },
      ],
      blankTotal: true,
    }),

    para([{ text: 'Question 2.4', bold: true }]),
    sub('(a) Total the Purchases Day Book.'),
    sub('(b) Post to individual supplier accounts and balance both.'),
    sub('(c) Prepare a summarised Trial Balance as at 10 June.'),

    para([{ text: 'Question 2.5', bold: true }]),
    sub('(a) Prepare a Purchases Ledger Control Account.'),
    sub('(b) Prepare a more summarised Trial Balance.'),

    ...cashBookExtract({
      title: 'Cash Book (Bank Account) — Payments',
      mode: 'payments',
      rows: [
        { date: '18 Jun', name: 'Apex Components',     amount: 840 },
        { date: '20 Jun', name: 'Clearwater Supplies', amount: 290 },
      ],
    }),

    para([{ text: 'Question 2.6', bold: true }]),
    sub('(a) Update the PLCA. What is the closing balance?'),
    sub('(b) Update both individual accounts in the Purchases Ledger.'),
    sub('(c) Confirm PLCA agrees with the sum of the Purchases Ledger balances.'),
    sub('(d) What would the discrepancy have been if the Purchases Ledger had not been updated?'),

    // ---- Answers 2.4–2.6 --------------------------------------------------
    h3('Answers — 2.4–2.6'),

    para([{ text: 'Answer 2.4:', bold: true }]),
    para('Day Book total: £1,500'),

    ...tAccount({
      title: 'Apex Components — Purchases Ledger',
      drRows: [
        { desc: 'Balance c/d', amount: '1,210' },
        {},
        { isTotal: true, amount: '£1,210' },
        {},
      ],
      crRows: [
        { desc: '2 Jun  Purchases PDB/601',  amount: '840' },
        { desc: '10 Jun  Purchases PDB/603', amount: '370' },
        { isTotal: true, amount: '£1,210' },
        { bold: true, desc: 'Bal b/d', amount: '£1,210' },
      ],
    }),

    ...tAccount({
      title: 'Clearwater Supplies — Purchases Ledger',
      drRows: [
        { desc: 'Balance c/d', amount: '290' },
        { isTotal: true, amount: '£290' },
        {},
      ],
      crRows: [
        { desc: '6 Jun  Purchases PDB/602', amount: '290' },
        { isTotal: true, amount: '£290' },
        { bold: true, desc: 'Bal b/d', amount: '£290' },
      ],
    }),

    ...trialBalance({
      title: 'Trial Balance — 10 June 2026',
      rows: [
        { account: 'Purchases',            dr: 1500, cr: '' },
        { account: 'Apex Components',      dr: '',   cr: 1210 },
        { account: 'Clearwater Supplies',  dr: '',   cr: 290 },
      ],
    }),

    para([{ text: 'Answer 2.5:', bold: true }]),

    ...tAccount({
      title: 'Purchases Ledger Control Account',
      drRows: [
        { desc: '10 Jun  Balance c/d', amount: '1,500' },
        { isTotal: true, amount: '£1,500' },
        {},
      ],
      crRows: [
        { desc: '10 Jun  Credit Purchases', amount: '1,500' },
        { isTotal: true, amount: '£1,500' },
        { bold: true, desc: 'Bal b/d', amount: '£1,500' },
      ],
    }),

    ...trialBalance({
      title: 'Trial Balance — 10 June 2026',
      rows: [
        { account: 'Purchases',                         dr: 1500, cr: '' },
        { account: 'Purchases Ledger Control Account',  dr: '',   cr: 1500 },
      ],
    }),

    para([{ text: 'Answer 2.6:', bold: true }]),

    ...tAccount({
      title: 'Purchases Ledger Control Account',
      drRows: [
        { desc: '20 Jun  Bank',         amount: '1,130' },
        { desc: '20 Jun  Balance c/d',  amount: '370' },
        { isTotal: true, amount: '£1,500' },
        {},
      ],
      crRows: [
        { desc: 'Bal b/f', amount: '1,500' },
        {},
        { isTotal: true, amount: '£1,500' },
        { bold: true, desc: 'Bal b/d', amount: '£370' },
      ],
    }),

    ...tAccount({
      title: 'Apex Components — Purchases Ledger',
      drRows: [
        { desc: '18 Jun  Bank', amount: '840' },
        { desc: 'Balance c/d',  amount: '370' },
        { isTotal: true, amount: '£1,210' },
        {},
      ],
      crRows: [
        { desc: 'Bal b/f', amount: '1,210' },
        {},
        { isTotal: true, amount: '£1,210' },
        { bold: true, desc: 'Bal b/d', amount: '£370' },
      ],
    }),

    ...tAccount({
      title: 'Clearwater Supplies — Purchases Ledger',
      drRows: [
        { desc: '20 Jun  Bank', amount: '290' },
        { isTotal: true, amount: '£290' },
      ],
      crRows: [
        { desc: 'Bal b/f', amount: '290' },
        { isTotal: true, amount: '£290' },
      ],
    }),
    para([{ text: '(Nil balance)', italics: true }]),

    para('PLCA £370. Purchases Ledger: £370 + £0 = £370. Agreed. ' + TICK),
    para('Without updating: PLCA £370 vs individual total £1,500. Discrepancy: £1,130.'),
  ];
}

// ----- PART 3 — FURTHER ENTRIES -------------------------------------------
function part3() {
  return [
    pageBreakPara(),
    h1('Part 3 — Further Entries in the Control Accounts'),

    para('The examples so far have featured only credit sales/purchases and cash receipts/payments. In practice, several other entry types arise. These are introduced here gradually, starting with the two that most often catch students out.'),

    h3('Sales Returns'),
    para('When a customer returns goods, a credit note is issued. This reduces the amount owed. In the SLCA: Credit entry. In the customer’s individual Receivables Ledger account: Debit entry.'),

    h3('Irrecoverable Debts (Bad Debts Written Off)'),
    para('When a debt cannot be collected, it is written off. In the SLCA: Credit entry (asset removed). Debit goes to the Irrecoverable Debts Expense account.'),

    // ---- Worked Example 3.1 -----------------------------------------------
    h3('Worked Example 3.1 — Sales Returns and Irrecoverable Debts'),
    para([
      { text: 'Company: ', bold: true }, 'Hillcrest Ltd. ',
      { text: 'Opening SLCA balance: £2,400 Dr.', bold: true },
    ]),

    ...transactionSummary({
      rows: [
        { transaction: 'Credit sales for the quarter',                  source: 'Sales Day Book',   amount: 3600 },
        { transaction: 'Sales returns (credit notes issued)',           source: 'Returns Journal',  amount: 280 },
        { transaction: 'Irrecoverable debt written off (J. Briggs)',    source: 'General Journal',  amount: 150 },
        { transaction: 'Bank receipts from credit customers',           source: 'Cash Book',        amount: 1800 },
      ],
    }),

    ...tAccount({
      title: 'Sales Ledger Control Account',
      drRows: [
        { desc: 'Bal b/f',       amount: '£2,400' },
        { desc: 'Credit Sales',  amount: '3,600' },
        {},
        {},
        { isTotal: true, amount: '£6,000' },
        { bold: true, desc: 'Bal b/d', amount: '£3,770' },
      ],
      crRows: [
        { desc: 'Sales Returns',      amount: '280' },
        { desc: 'Irrecoverable Debts', amount: '150' },
        { desc: 'Bank',               amount: '1,800' },
        { desc: 'Balance c/d',        amount: '3,770' },
        { isTotal: true, amount: '£6,000' },
        {},
      ],
    }),

    // ---- Dishonoured Cheques and Interest Charged --------------------------
    h3('Dishonoured Cheques and Interest Charged'),
    para([{ text: 'Dishonoured cheques: ', bold: true }, 'If a customer’s cheque is returned unpaid by the bank, the debt is reinstated. Debit the SLCA (the asset is restored). Credit the Cash Book.']),
    para([{ text: 'Interest charged: ', bold: true }, 'If interest is added to overdue accounts, the amount owed increases. Debit the SLCA.']),

    // ---- Worked Example 3.2 -----------------------------------------------
    h3('Worked Example 3.2 — Dishonoured Cheques and Interest Charged'),
    para([
      'Hillcrest Ltd continues. ',
      { text: 'Opening SLCA balance: £3,770 Dr.', bold: true },
    ]),

    ...transactionSummary({
      rows: [
        { transaction: 'Credit sales for the quarter',                  source: 'Sales Day Book',  amount: 4200 },
        { transaction: 'Dishonoured cheque — Customer K reinstated',    source: 'Cash Book',       amount: 320 },
        { transaction: 'Interest charged to overdue accounts',          source: 'General Journal', amount: 90 },
        { transaction: 'Bank receipts from credit customers',           source: 'Cash Book',       amount: 3100 },
        { transaction: 'Sales returns',                                 source: 'Returns Journal', amount: 450 },
      ],
    }),

    ...tAccount({
      title: 'Sales Ledger Control Account',
      drRows: [
        { desc: 'Bal b/f',            amount: '£3,770' },
        { desc: 'Credit Sales',       amount: '4,200' },
        { desc: 'Dishonoured Cheque', amount: '320' },
        { desc: 'Interest Charged',   amount: '90' },
        { isTotal: true, amount: '£8,380' },
        { bold: true, desc: 'Bal b/d', amount: '£4,830' },
      ],
      crRows: [
        { desc: 'Bank',          amount: '3,100' },
        { desc: 'Sales Returns', amount: '450' },
        {},
        { desc: 'Balance c/d',   amount: '4,830' },
        { isTotal: true, amount: '£8,380' },
        {},
      ],
    }),

    // ---- Contras -----------------------------------------------------------
    h3('Contras (Set-off Entries)'),
    para('A contra arises when the same entity is both a credit customer (Sales Ledger) and a credit supplier (Purchases Ledger). The mutual debts are offset: Credit the SLCA; Debit the PLCA. No cash changes hands.'),

    ...noteBox({
      title: 'Exam Trap — Cash Sales',
      content: 'Questions frequently give a total sales figure that includes both cash and credit sales. Only credit sales are posted to the SLCA. Cash sales bypass the control account entirely. Always extract the credit figure before posting.',
    }),

    // ---- Worked Example 3.3 -----------------------------------------------
    h3('Worked Example 3.3 — Comprehensive SLCA (All Entry Types)'),
    para([
      { text: 'Company: ', bold: true }, 'Horizon Wholesale Supplies. ',
      { text: 'Opening SLCA balance: £8,200 Dr.', bold: true },
    ]),

    ...transactionSummary({
      rows: [
        { transaction: 'Total sales (including £1,500 cash sales — exclude)', source: 'Day Book / Cash Book', amount: 24500 },
        { transaction: 'Bank receipts from credit customers',                  source: 'Cash Book',            amount: 19400 },
        { transaction: 'Sales returns',                                        source: 'Returns Journal',      amount: 680 },
        { transaction: 'Interest charged to overdue accounts',                 source: 'General Journal',      amount: 120 },
        { transaction: 'Dishonoured cheque reinstated',                        source: 'Cash Book',            amount: 450 },
        { transaction: 'Irrecoverable debt written off',                       source: 'General Journal',      amount: 310 },
        { transaction: 'Contra — set-off against Purchases Ledger balance',    source: 'General Journal',      amount: 400 },
      ],
    }),

    para([{ text: 'Credit sales only: £24,500 ' + MINUS + ' £1,500 = £23,000.', italics: true }]),

    ...tAccount({
      title: 'Sales Ledger Control Account',
      drRows: [
        { desc: 'Bal b/f',            amount: '£8,200' },
        { desc: 'Credit Sales',       amount: '23,000' },
        { desc: 'Interest Charged',   amount: '120' },
        { desc: 'Dishonoured Cheque', amount: '450' },
        {},
        { isTotal: true, amount: '£31,770' },
        { bold: true, desc: 'Bal b/d', amount: '£10,980' },
      ],
      crRows: [
        { desc: 'Bank',                amount: '19,400' },
        { desc: 'Sales Returns',       amount: '680' },
        { desc: 'Irrecoverable Debts', amount: '310' },
        { desc: 'Contra / Set-off',    amount: '400' },
        { desc: 'Balance c/d',         amount: '10,980' },
        { isTotal: true, amount: '£31,770' },
        {},
      ],
    }),

    // ---- Worked Example 3.4 -----------------------------------------------
    h3('Worked Example 3.4 — Comprehensive PLCA'),

    ...noteBox({
      title: 'Exam Trap — Cash Purchases',
      content: 'Same principle applies. Exclude cash purchases from the PLCA.',
    }),

    para([{ text: 'Company: ', bold: true }, 'Summit Components.']),
    para([{ text: 'Opening PLCA balances: £11,600 Cr (normal); £250 Dr (minority — historical overpayment to one supplier).', bold: true }]),

    ...transactionSummary({
      rows: [
        { transaction: 'Total purchases (including £3,100 cash — exclude)',  source: 'Day Book / Cash Book', amount: 28400 },
        { transaction: 'Bank payments to credit suppliers',                   source: 'Cash Book',            amount: 19800 },
        { transaction: 'Purchases returns',                                   source: 'Returns Journal',      amount: 620 },
        { transaction: 'Discounts received from suppliers',                   source: 'Cash Book',            amount: 180 },
        { transaction: 'Interest charged to us by suppliers',                 source: 'General Journal',      amount: 90 },
        { transaction: 'Contra — set-off against Sales Ledger',               source: 'General Journal',      amount: 670 },
      ],
    }),

    para([{ text: 'Credit purchases only: £28,400 ' + MINUS + ' £3,100 = £25,300. Minority debit balance of £140 remains at period end.', italics: true }]),

    ...tAccount({
      title: 'Purchases Ledger Control Account',
      drRows: [
        { desc: 'Bal b/f (Minority)',  amount: '£250' },
        { desc: 'Bank',                amount: '19,800' },
        { desc: 'Purchases Returns',   amount: '620' },
        { desc: 'Discounts Received',  amount: '180' },
        { desc: 'Contra',              amount: '670' },
        { desc: 'Bal c/d (Normal)',    amount: '15,610' },
        { isTotal: true, amount: '£37,130' },
        { bold: true, desc: 'Bal b/d (Minority)', amount: '£140' },
      ],
      crRows: [
        { desc: 'Bal b/f (Normal)',    amount: '£11,600' },
        { desc: 'Credit Purchases',    amount: '25,300' },
        { desc: 'Interest Charged',    amount: '90' },
        { desc: 'Bal c/d (Minority)',  amount: '140' },
        {},
        {},
        { isTotal: true, amount: '£37,130' },
        { bold: true, desc: 'Bal b/d (Normal)', amount: '£15,610' },
      ],
    }),
  ];
}

// ----- PART 4 — DISTRACTIONS ----------------------------------------------
function part4() {
  return [
    pageBreakPara(),
    h1('Part 4 — Things That Can Cause Confusion'),

    para('Two features of control accounts often trip students up mid-question. They are worth examining directly before moving on to errors.'),

    h3('4.1 Simultaneous Debit and Credit Balances'),
    para('The SLCA normally carries a debit balance (asset). Occasionally, a small number of customers may carry a credit balance — typically because they have overpaid or returned goods after their invoice was fully settled. The SLCA must carry both balances simultaneously. They must never be netted off — both are brought down separately.'),

    h3('Question 4.1 — SLCA with a Minority Credit Balance'),
    para([{ text: 'Company: ', bold: true }, 'Peak Manufacturing.']),
    para([{ text: 'Opening SLCA balances: £14,200 Dr (normal); £310 Cr (minority).', bold: true }]),

    ...transactionSummary({
      rows: [
        { transaction: 'Total sales (including £2,400 cash sales)',           source: 'Day Book',         amount: 31650 },
        { transaction: 'Bank receipts from credit customers',                  source: 'Cash Book',        amount: 24100 },
        { transaction: 'Sales returns',                                        source: 'Returns Journal',  amount: 850 },
        { transaction: 'Dishonoured cheque reinstated',                        source: 'Cash Book',        amount: 520 },
        { transaction: 'Irrecoverable debts written off',                      source: 'General Journal',  amount: 440 },
        { transaction: 'Interest charged to overdue accounts',                 source: 'General Journal',  amount: 160 },
        { transaction: 'Contra — set-off against Purchases Ledger',            source: 'General Journal',  amount: 670 },
      ],
    }),

    para('At 30 November, a minority credit balance of £190 still exists.'),
    para([{ text: 'Required: ', bold: true }, 'Prepare the SLCA for November showing both closing balances.']),

    h3('Answer — 4.1'),
    para('Credit sales: £31,650 ' + MINUS + ' £2,400 = £29,250.'),

    ...tAccount({
      title: 'Sales Ledger Control Account',
      drRows: [
        { desc: 'Nov 1  Bal b/f (Normal)', amount: '£14,200' },
        { desc: 'Credit Sales',            amount: '29,250' },
        { desc: 'Interest Charged',        amount: '160' },
        { desc: 'Dishonoured Cheque',      amount: '520' },
        { desc: 'Bal c/d (Minority)',      amount: '190' },
        { isTotal: true, amount: '£44,320' },
        { bold: true, desc: 'Dec 1  Bal b/d (Normal)', amount: '£17,950' },
        {},
      ],
      crRows: [
        { desc: 'Nov 1  Bal b/f (Minority)', amount: '£310' },
        { desc: 'Bank',                      amount: '24,100' },
        { desc: 'Sales Returns',             amount: '850' },
        { desc: 'Irrecoverable Debts',       amount: '440' },
        { desc: 'Contra',                    amount: '670' },
        { desc: 'Bal c/d (Normal)',          amount: '17,950' },
        { isTotal: true, amount: '£44,320' },
        { bold: true, desc: 'Dec 1  Bal b/d (Minority)', amount: '£190' },
      ],
    }),

    h3('4.2 Simultaneous Balances in the PLCA'),
    para('The PLCA normally carries a credit balance (liability). The equivalent situation is a minority debit balance — a small number of suppliers owe the business money, typically due to overpayment. Both balances are brought down separately.'),

    h3('Question 4.2 — PLCA with a Minority Debit Balance'),
    para([{ text: 'Company: ', bold: true }, 'Summit Trading.']),
    para([{ text: 'Opening PLCA balances: £6,800 Cr (normal); £200 Dr (minority).', bold: true }]),

    ...transactionSummary({
      rows: [
        { transaction: 'Credit purchases',                  source: 'Purchases Day Book', amount: 9400 },
        { transaction: 'Bank payments to credit suppliers', source: 'Cash Book',          amount: 7600 },
        { transaction: 'Purchases returns',                 source: 'Returns Journal',    amount: 380 },
        { transaction: 'Discounts received',                source: 'Cash Book',          amount: 120 },
      ],
    }),

    para('At 31 October, a minority debit balance of £90 remains.'),
    para([{ text: 'Required: ', bold: true }, 'Prepare the PLCA for October showing both closing balances.']),

    h3('Answer — 4.2'),

    ...tAccount({
      title: 'Purchases Ledger Control Account',
      drRows: [
        { desc: 'Oct 1  Bal b/f (Minority)', amount: '£200' },
        { desc: 'Bank',                      amount: '7,600' },
        { desc: 'Purchases Returns',         amount: '380' },
        { desc: 'Discounts Received',        amount: '120' },
        { desc: 'Bal c/d (Normal)',          amount: '7,990' },
        { isTotal: true, amount: '£16,290' },
        { bold: true, desc: 'Nov 1  Bal b/d (Minority)', amount: '£90' },
      ],
      crRows: [
        { desc: 'Oct 1  Bal b/f (Normal)', amount: '£6,800' },
        { desc: 'Credit Purchases',        amount: '9,400' },
        {},
        {},
        { desc: 'Bal c/d (Minority)',      amount: '90' },
        { isTotal: true, amount: '£16,290' },
        { bold: true, desc: 'Nov 1  Bal b/d (Normal)', amount: '£7,990' },
      ],
    }),
  ];
}

// ----- PART 5 — ERRORS ----------------------------------------------------
function part5() {
  return [
    pageBreakPara(),
    h1('Part 5 — Understanding Errors'),

    para('When there is a discrepancy between the SLCA balance and the total of the individual Receivables Ledger balances, an error has been made. Different error types corrupt different tracks. Rather than being told which is which, the approach here is to make each type of error deliberately and observe the result first-hand.'),

    para([{ text: 'Opening position — shared by all questions in Part 5:', bold: true }]),

    ...smallTable({
      headers: ['', '£'],
      widths: [7650, 2550],
      amountColIndex: 1,
      boldLastRow: true,
      rows: [
        ['Sales Ledger Control Account (Track 1)', '2,400 Dr'],
        ['Patel & Co — Receivables Ledger',        '1,400'],
        ['Okafor Ltd — Receivables Ledger',        '1,000'],
        ['Total of individual accounts',           '2,400'],
      ],
    }),

    para([{ text: 'Transactions during March 2026:', bold: true }]),

    ...smallTable({
      headers: ['Transaction', '£'],
      widths: [7650, 2550],
      amountColIndex: 1,
      rows: [
        ['Credit sales — Patel & Co (Invoice SDB/301)', '360'],
        ['Credit sales — Okafor Ltd (Invoice SDB/302)', '270'],
        ['Sales Day Book total',                        '630'],
        ['Bank receipt — Patel & Co',                   '800'],
      ],
    }),

    // ---- Q5.1 --------------------------------------------------------------
    h3('Question 5.1 — Making an Omission Error'),
    para([{ text: 'Step 1: Do it correctly. ', bold: true }, 'Complete the SLCA and both Receivables Ledger accounts accurately. Confirm the SLCA balance agrees with the total of the individual accounts.']),
    para([{ text: 'Step 2: Make the error. ', bold: true }, 'Suppose the SLCA is updated correctly (using the day book total of £630), but the posting of Okafor Ltd’s invoice SDB/302 (£270) is omitted from their personal account. What is now the discrepancy? Which track is wrong?']),

    h3('Answer — 5.1'),
    para([{ text: 'Step 1 — Correct position:', bold: true }]),

    ...tAccount({
      title: 'Sales Ledger Control Account',
      drRows: [
        { desc: 'Bal b/f',      amount: '£2,400' },
        { desc: 'Credit Sales', amount: '630' },
        { isTotal: true, amount: '£3,030' },
        { bold: true, desc: 'Bal b/d', amount: '£2,230' },
      ],
      crRows: [
        { desc: 'Bank (Patel)', amount: '800' },
        { desc: 'Balance c/d',  amount: '2,230' },
        { isTotal: true, amount: '£3,030' },
        {},
      ],
    }),

    ...tAccount({
      title: 'Patel & Co — Receivables Ledger',
      drRows: [
        { desc: 'Bal b/f',         amount: '£1,400' },
        { desc: 'Sales SDB/301',   amount: '360' },
        { isTotal: true, amount: '£1,760' },
        { bold: true, desc: 'Bal b/d', amount: '£960' },
      ],
      crRows: [
        { desc: 'Bank',        amount: '800' },
        { desc: 'Balance c/d', amount: '960' },
        { isTotal: true, amount: '£1,760' },
        {},
      ],
    }),

    ...tAccount({
      title: 'Okafor Ltd — Receivables Ledger',
      drRows: [
        { desc: 'Bal b/f',        amount: '£1,000' },
        { desc: 'Sales SDB/302',  amount: '270' },
        { isTotal: true, amount: '£1,270' },
        { bold: true, desc: 'Bal b/d', amount: '£1,270' },
      ],
      crRows: [
        {},
        { desc: 'Balance c/d', amount: '1,270' },
        { isTotal: true, amount: '£1,270' },
        {},
      ],
    }),

    para('SLCA: £2,230. Individual total: £960 + £1,270 = £2,230. Agreed. ' + TICK),

    para([{ text: 'Step 2 — With the omission (SDB/302 not posted to Okafor’s account):', bold: true }]),

    ...tAccount({
      title: 'Okafor Ltd — Receivables Ledger (Omission)',
      drRows: [
        { desc: 'Bal b/f', amount: '£1,000' },
        { isTotal: true, amount: '£1,000' },
        { bold: true, desc: 'Bal b/d', amount: '£1,000' },
      ],
      crRows: [
        { desc: 'Balance c/d', amount: '1,000' },
        { isTotal: true, amount: '£1,000' },
        {},
      ],
    }),

    para('SLCA: £2,230 (correct — day book total was used).'),
    para('Individual total: £960 + £1,000 = £1,960.'),
    para([{ text: 'Discrepancy: £2,230 ' + MINUS + ' £1,960 = £270 — the value of the missing invoice exactly.', bold: true }]),
    para('Conclusion: Track 1 (SLCA) is correct. The error is in Track 2 (Receivables Ledger) only.'),

    // ---- Q5.2 --------------------------------------------------------------
    h3('Question 5.2 — Making a Casting Error'),
    para([{ text: 'Narrative: ', bold: true }, 'A casting error occurs when the day book total is calculated incorrectly. Because the SLCA is posted from the day book total, the error enters Track 1. Individual accounts are posted from the actual invoices, so Track 2 remains correct.']),
    para('Using the same March position and transactions:'),

    para([{ text: 'Part A — Under-cast', bold: true }]),
    para('Suppose the day book total is calculated as £530 (£100 too low). The SLCA is updated using £530; individual accounts are posted correctly from the actual invoices. What is the resulting discrepancy?'),

    para([{ text: 'Answer — Part A:', bold: true }]),
    para(['SLCA: £2,400 + £530 ' + MINUS + ' £800 = ', { text: '£2,130.', bold: true }]),
    para(['Individual accounts (correct): £960 + £1,270 = ', { text: '£2,230.', bold: true }]),
    para(['Discrepancy: £2,230 ' + MINUS + ' £2,130 = ', { text: '£100', bold: true }, ' — exactly the under-cast amount.']),
    para('Conclusion: Track 2 is correct. Track 1 (SLCA) is understated by £100.'),

    para([{ text: 'Part B — Over-cast', bold: true }]),
    para('Suppose the day book total is calculated as £780 (£150 too high). What is the discrepancy?'),

    para([{ text: 'Answer — Part B:', bold: true }]),
    para(['SLCA: £2,400 + £780 ' + MINUS + ' £800 = ', { text: '£2,380.', bold: true }]),
    para(['Individual accounts: ', { text: '£2,230.', bold: true }]),
    para(['Discrepancy: £2,380 ' + MINUS + ' £2,230 = ', { text: '£150', bold: true }, ' — exactly the over-cast amount.']),
    para('Conclusion: Track 2 is correct. Track 1 (SLCA) is overstated by £150.'),

    // ---- Q5.3 --------------------------------------------------------------
    h3('Question 5.3 — Making a Transposition Error'),
    para([{ text: 'Narrative: ', bold: true }, 'A transposition error occurs when two digits are accidentally swapped. This happens at the point of posting to an individual account, so it corrupts Track 2 only. The SLCA, posted from the day book total, is unaffected.']),
    para('The SLCA is updated correctly. When posting Okafor Ltd’s invoice of £270, the digits are transposed and £720 is entered in their personal account. Calculate both balances and state the discrepancy.'),

    para([{ text: 'Answer — 5.3:', bold: true }]),
    para('SLCA: £2,230 (correct, as in Step 1 of Q5.1).'),
    para('Patel & Co: £960 (correct).'),
    para(['Okafor Ltd: £1,000 + £720 = ', { text: '£1,720', bold: true }, ' (should be £1,270).']),
    para(['Individual total: £960 + £1,720 = ', { text: '£2,680.', bold: true }]),
    para(['Discrepancy: £2,680 ' + MINUS + ' £2,230 = ', { text: '£450.', bold: true }, ' Note: £720 ' + MINUS + ' £270 = £450 — always the difference between the transposed figures.']),
    para('Conclusion: Track 1 (SLCA) is correct. Track 2 (individual accounts) is overstated by £450.'),

    // ---- Q5.4 --------------------------------------------------------------
    h3('Question 5.4 — Consolidation: Identifying and Correcting Mixed Errors'),
    para('After a further period, the following position is reported:'),

    ...smallTable({
      headers: ['', '£'],
      widths: [7650, 2550],
      amountColIndex: 1,
      boldLastRow: true,
      rows: [
        ['Sales Ledger Control Account (Track 1)',                       '5,200'],
        ['Total of individual Receivables Ledger balances (Track 2)',    '5,080'],
        ['Discrepancy',                                                  '120'],
      ],
    }),

    para('Investigation reveals the following errors:'),
    new Paragraph({
      numbering: { reference: 'numbers', level: 0 },
      children: [run('The Sales Day Book total for the period was over-cast by £60.')],
    }),
    new Paragraph({
      numbering: { reference: 'numbers', level: 0 },
      children: [run('A credit sale of £100 to D. Hughes was correctly posted to the SLCA via the day book but was entirely omitted from D. Hughes’s personal account in the Receivables Ledger.')],
    }),
    new Paragraph({
      numbering: { reference: 'numbers', level: 0 },
      children: [run('A bank receipt of £240 from K. Osei was correctly entered in the SLCA but posted as £200 in K. Osei’s personal account (i.e. only £200 was deducted from the balance, not £240).')],
    }),

    para([{ text: 'Required: ', bold: true }, 'For each error, state which track is affected and the correction needed. Then calculate the corrected balance for both tracks and confirm they agree.']),

    h3('Answer — 5.4'),

    para([{ text: 'Error analysis:', bold: true }]),
    ...smallTable({
      headers: ['#', 'Error type', 'Track 1 (SLCA) adjustment', 'Track 2 (Individual accounts) adjustment'],
      widths: [600, 2900, 3350, 3350],
      rows: [
        ['1', 'Over-cast SDB',                  'Reduce by £60 (SLCA overstated)', 'No effect'],
        ['2', 'Omission from personal account', 'No effect',                       'Add £100 to D. Hughes’s account (understated)'],
        ['3', 'Under-posted receipt',           'No effect',                       'K. Osei’s balance is £40 too high (only £200 deducted instead of £240) — credit account by further £40'],
      ],
    }),

    para([{ text: 'Reconciliation:', bold: true }]),
    ...reconciliationTable({
      headers: ['Track 1 — SLCA', 'Track 2 — Individual Accounts'],
      rows: [
        { l1: 'Original balance',              l1Amount: '5,200',  l2: 'Original total',                                 l2Amount: '5,080' },
        { l1: 'Less: Error 1 (over-cast SDB)', l1Amount: MINUS + '60', l2: 'Add: Error 2 (omission — D. Hughes)',            l2Amount: '+100' },
        { l1: '',                              l1Amount: '',       l2: 'Less: Error 3 (under-posted receipt — K. Osei)', l2Amount: MINUS + '40' },
        { l1: 'Corrected SLCA balance',        l1Amount: '£5,140', l2: 'Corrected individual accounts total',            l2Amount: '£5,140' },
      ],
    }),

    para('Both tracks agree at £5,140. ' + TICK),
  ];
}

const children = [
  ...titlePage(),
  ...part1(),
  ...part2(),
  ...part3(),
  ...part4(),
  ...part5(),
];

// ---------------------------------------------------------------------------
// 6. ASSEMBLE & WRITE
// ---------------------------------------------------------------------------

const doc = new Document({
  styles,
  numbering,
  sections: [{
    properties: { page: { size: { width: PAGE.width, height: PAGE.height }, margin: PAGE.margin } },
    children,
  }],
});

const OUT = '/mnt/user-data/outputs/control_accounts_guide.docx';
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('Wrote', OUT, '(', buf.length, 'bytes )');
});
