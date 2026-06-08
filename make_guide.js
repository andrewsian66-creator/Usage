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
// 4. HELPER FUNCTIONS (Section 4) — TO BE FILLED IN NEXT PASS
// ---------------------------------------------------------------------------
// Planned helpers:
//   - dayBookTable({ title, rows, blankTotal })          // §4.1
//   - tAccount({ title, drRows, crRows })                // §4.2
//   - trialBalance({ title, rows })                       // §4.3
//   - cashBookExtract({ title, mode, rows })             // §4.4
//   - transactionSummary({ rows })                        // §4.5
//   - noteBox({ title, body })                            // §4.6
//   - reconciliationTable({ headers, rows })              // §4.7
//   - body(text, opts), bold(text), pageBreak()           // misc

// ---------------------------------------------------------------------------
// 5. CONTENT (Section 6) — TO BE FILLED IN NEXT PASS
// ---------------------------------------------------------------------------

const children = [
  // Placeholder until content is built
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Control Accounts (placeholder shell)', bold: true })],
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
