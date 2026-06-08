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
// 5. SMOKE TEST — verify all seven helpers render correctly.
//    (Replaced with real Section 6 content in a later pass.)
// ---------------------------------------------------------------------------

const children = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 240 },
    children: [run('Helper smoke test: Day Book + T-Account', { bold: true, size: 28 })],
  }),

  // §4.1 smoke test: Sales Day Book from Worked Example 1.1
  ...dayBookTable({
    title: 'Sales Day Book — May 2026',
    rows: [
      { date: '1 May', name: 'A. Brown',  invoice: 'SDB/101', amount: 400 },
      { date: '3 May', name: 'C. Davies', invoice: 'SDB/102', amount: 180 },
      { date: '7 May', name: 'A. Brown',  invoice: 'SDB/103', amount: 250 },
    ],
  }),

  // §4.1 smoke test: blank-total variant (for practice questions)
  ...dayBookTable({
    title: 'Sales Day Book — June 2026 (blank total)',
    rows: [
      { date: '1 Jun', name: 'M. Khan',   invoice: 'SDB/201', amount: 520 },
      { date: '4 Jun', name: 'P. Thomas', invoice: 'SDB/202', amount: 340 },
      { date: '9 Jun', name: 'M. Khan',   invoice: 'SDB/203', amount: 180 },
    ],
    blankTotal: true,
  }),

  // §4.2 smoke test: A. Brown's Sales Ledger account (Worked Example 1.1)
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

  // §4.2 smoke test: padding — Dr longer than Cr (unbalanced lengths)
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
    ],
  }),

  // §4.3 smoke test: Trial Balance (auto-totalled)
  ...trialBalance({
    title: 'Summarised Trial Balance — 7 May 2026',
    rows: [
      { account: 'A. Brown',  dr: 650, cr: '' },
      { account: 'C. Davies', dr: 180, cr: '' },
      { account: 'Sales',     dr: '',  cr: 830 },
    ],
  }),

  // §4.4 smoke test: Cash Book Extract — receipts (default) and payments
  ...cashBookExtract({
    title: 'Cash Book (Bank Account) — Receipts',
    mode: 'receipts',
    rows: [
      { date: '10 May', name: 'A. Brown',  amount: 400 },
      { date: '12 May', name: 'C. Davies', amount: 180 },
    ],
  }),
  ...cashBookExtract({
    title: 'Cash Book (Bank Account) — Payments',
    mode: 'payments',
    rows: [
      { date: '14 May', name: 'T. Mills Ltd',   amount: 600 },
      { date: '16 May', name: 'B. Rogers & Co', amount: 200 },
    ],
  }),

  // §4.5 smoke test: Transaction Summary
  ...transactionSummary({
    title: 'Transaction summary',
    rows: [
      { transaction: 'Credit sales for the quarter',        source: 'Sales Day Book',  amount: 3600 },
      { transaction: 'Sales returns (credit notes issued)', source: 'Returns Journal', amount: 280 },
      { transaction: 'Irrecoverable debt written off',      source: 'General Journal', amount: 150 },
      { transaction: 'Bank receipts from credit customers', source: 'Cash Book',       amount: 1800 },
    ],
  }),

  // §4.6 smoke test: Note Box
  ...noteBox({
    title: 'Key Principle',
    content: 'The Receivables Ledger must be updated every time an entry is made in the Sales Ledger Control Account. The SLCA balance and the total of the individual Receivables Ledger balances must always agree.',
  }),

  // §4.7 smoke test: Reconciliation Table (last row bold)
  ...reconciliationTable({
    title: 'Error reconciliation',
    headers: ['Track 1 — SLCA', 'Track 2 — Individual Accounts'],
    rows: [
      { l1: 'Original balance',                 l1Amount: '5,200', l2: 'Original total',                              l2Amount: '5,080' },
      { l1: 'Less: Error 1 (over-cast SDB)',    l1Amount: '−60',   l2: 'Add: Error 2 (omission — D. Hughes)',          l2Amount: '+100' },
      { l1: '',                                 l1Amount: '',      l2: 'Less: Error 3 (under-posted receipt — K. Osei)', l2Amount: '−40' },
      { l1: 'Corrected SLCA balance',           l1Amount: '£5,140', l2: 'Corrected individual accounts total',         l2Amount: '£5,140' },
    ],
  }),
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
