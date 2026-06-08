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

// ---------------------------------------------------------------------------
// 5. SMOKE TEST — verify the two helpers render correctly.
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
