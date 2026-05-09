import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  HeadingLevel, PageBreak, Header, Footer, PageNumber, UnderlineType
} from '/home/claude/.npm-global/lib/node_modules/docx/dist/index.mjs';
import fs from 'fs';

// ─── Color palette ───────────────────────────────────────────────────────────
const NAVY   = "1B3A5C";   // deep navy — headings, accents
const STEEL  = "2E6DA4";   // medium blue — section headers
const LTBLUE = "D6E4F0";   // light blue — section header background
const CREAM  = "F7F9FC";   // near-white — input cell bg
const LGRAY  = "E8ECF0";   // light gray — label cell bg
const MGRAY  = "9AA5B1";   // mid gray — placeholder text color
const BLACK  = "1A1A1A";   // near-black body text
const WHITE  = "FFFFFF";
const NONE   = "FFFFFF";   // for "no border" (white border effectively)

// ─── Helpers ─────────────────────────────────────────────────────────────────
const pt = (n) => n * 20;      // points → half-points for w:sz
const inch = (n) => n * 1440;  // inches → DXA

// Standard borders
const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: "C5CDD6" };
const noBorder   = { style: BorderStyle.NONE, size: 0, color: WHITE };
const allThin    = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noAllBorder= { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// Page content width: 8.5" - 1.25" left - 1" right = 6.25" = 9000 DXA
const PAGE_W = 9000;

function label(text, opts = {}) {
  return new TextRun({
    text,
    font: "Calibri",
    size: opts.size || 18,
    bold: opts.bold || false,
    color: opts.color || BLACK,
    italics: opts.italics || false,
  });
}

function labelPara(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: 0, after: 0 },
    children: [label(text, opts)],
  });
}

// Section header row spanning full width
function sectionHeader(text) {
  return new Table({
    width: { size: PAGE_W, type: WidthType.DXA },
    columnWidths: [PAGE_W],
    borders: { ...noAllBorder, bottom: thinBorder },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: PAGE_W, type: WidthType.DXA },
            shading: { fill: NAVY, type: ShadingType.CLEAR },
            borders: noAllBorder,
            margins: { top: 80, bottom: 80, left: 180, right: 180 },
            children: [new Paragraph({
              spacing: { before: 0, after: 0 },
              children: [new TextRun({ text, font: "Calibri", size: 22, bold: true, color: WHITE, allCaps: true })]
            })]
          })
        ]
      })
    ]
  });
}

// Empty paragraph spacer
function spacer(before = 80, after = 80) {
  return new Paragraph({ spacing: { before, after }, children: [] });
}

// ─── SDT Content Control XML snippets (injected via XML hack via docx internal) ──
// docx-js doesn't natively support SDT, so we use a trick:
// We'll build the doc with special placeholder text and then post-process XML.
// Actually — we'll use the xmlData escape hatch in docx-js for raw XML paragraphs.

// Build a raw XML paragraph containing an SDT text field
let _sdtId = 1000;
function sdtField(placeholderLabel, widthDxa, tag, multiline = false) {
  const id = ++_sdtId;
  const tagClean = (tag || placeholderLabel).replace(/[^a-zA-Z0-9_]/g, '_').slice(0, 40);
  const ph = placeholderLabel.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  return `<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:sdtPr>
    <w:alias w:val="${ph}"/>
    <w:tag w:val="${tagClean}"/>
    <w:id w:val="${id}"/>
    <w:showingPlcHdr/>
    <w:text/>
  </w:sdtPr>
  <w:sdtContent>
    <w:r>
      <w:rPr>
        <w:rStyle w:val="PlaceholderText"/>
        <w:color w:val="${MGRAY}"/>
        <w:sz w:val="18"/>
        <w:szCs w:val="18"/>
      </w:rPr>
      <w:t>${ph}</w:t>
    </w:r>
  </w:sdtContent>
</w:sdt>`;
}

// Build a raw XML for a checkbox SDT
function sdtCheckbox(label) {
  const id = ++_sdtId;
  const tagClean = label.replace(/[^a-zA-Z0-9_]/g, '_').slice(0, 40);
  const lbl = label.replace(/&/g,'&amp;');
  return `<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:sdtPr>
    <w:alias w:val="${lbl}"/>
    <w:tag w:val="${tagClean}"/>
    <w:id w:val="${id}"/>
    <w14:checkbox>
      <w14:checked w14:val="0"/>
      <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
      <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
    </w14:checkbox>
  </w:sdtPr>
  <w:sdtContent>
    <w:r>
      <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
      <w:t>&#x2610;</w:t>
    </w:r>
  </w:sdtContent>
</w:sdt>`;
}

// A table cell that contains an SDT field
function inputCell(placeholder, width, tag, opts = {}) {
  const xmlData = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:spacing w:before="60" w:after="60"/>
    <w:ind w:left="60"/>
  </w:pPr>
  ${sdtField(placeholder, width, tag)}
</w:p>`;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: allThin,
    shading: { fill: CREAM, type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 100, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ xmlData })]
  });
}

// A label cell
function labelCell(text, width, opts = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: allThin,
    shading: { fill: opts.fill || LGRAY, type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 120, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [new TextRun({
        text,
        font: "Calibri",
        size: opts.size || 17,
        bold: opts.bold || false,
        color: NAVY,
      })]
    })]
  });
}

// A cell that just displays text (not a label, not an input)
function textCell(text, width, opts = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: allThin,
    shading: { fill: opts.fill || WHITE, type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 120, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: opts.span,
    children: [new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [new TextRun({ text, font: "Calibri", size: opts.size || 17, color: BLACK, bold: opts.bold })]
    })]
  });
}

// Checkbox cell: label + checkbox
function checkCell(checkLabel, displayLabel, width) {
  const xmlData = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>
  ${sdtCheckbox(checkLabel)}
  <w:r><w:rPr><w:sz w:val="18"/><w:color w:val="${BLACK}"/></w:rPr>
    <w:t xml:space="preserve"> ${displayLabel}</w:t>
  </w:r>
</w:p>`;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: allThin,
    shading: { fill: CREAM, type: ShadingType.CLEAR },
    margins: { top: 40, bottom: 40, left: 120, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ xmlData })]
  });
}

// Yes/No pair in one cell
function yesNoCell(yesLabel, noLabel, width) {
  const xmlData = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>
  ${sdtCheckbox(yesLabel)}
  <w:r><w:rPr><w:sz w:val="18"/><w:color w:val="${BLACK}"/></w:rPr>
    <w:t xml:space="preserve"> Yes&#x2003;</w:t>
  </w:r>
  ${sdtCheckbox(noLabel)}
  <w:r><w:rPr><w:sz w:val="18"/><w:color w:val="${BLACK}"/></w:rPr>
    <w:t xml:space="preserve"> No</w:t>
  </w:r>
</w:p>`;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: allThin,
    shading: { fill: CREAM, type: ShadingType.CLEAR },
    margins: { top: 40, bottom: 40, left: 120, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ xmlData })]
  });
}

// Row height helper
function rowHeight(val) { return { value: val, rule: "exact" }; }

// Standard row height for input rows
const ROW_H = { value: 460, rule: "atLeast" };
const LBL_H = { value: 360, rule: "atLeast" };

// ─── Helper: full-width input row (label on left, input on right) ─────────────
function fullRow(lbl, placeholder, tag, lblW = 2200) {
  const inputW = PAGE_W - lblW;
  return new TableRow({
    height: ROW_H,
    children: [
      labelCell(lbl, lblW),
      inputCell(placeholder, inputW, tag),
    ]
  });
}

// ─── TABLE BUILDER HELPER ────────────────────────────────────────────────────
function makeTable(colWidths, rows) {
  return new Table({
    width: { size: PAGE_W, type: WidthType.DXA },
    columnWidths: colWidths,
    borders: noAllBorder,
    rows,
  });
}

// ============================================================
// BUILD DOCUMENT SECTIONS
// ============================================================

// ── HEADER ───────────────────────────────────────────────────────────────────
function buildHeader() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: "Integrated Sycamore Group LLC", font: "Calibri", size: 52, bold: true, color: NAVY })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: "8315 Sierra Hill Ct  •  Houston TX 77083  •  Tel: 832-335-4098  /  713-835-2122", font: "Calibri", size: 18, color: STEEL })]
    }),
    // Horizontal rule via bottom-border paragraph
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 80 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 6 } },
      children: []
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 160 },
      children: [new TextRun({ text: "EMPLOYMENT APPLICATION", font: "Calibri", size: 36, bold: true, color: NAVY, allCaps: true })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({
        text: "Integrated Sycamore Group LLC provides equal employment opportunities without regard to race, color, sex, national origin, age, non-disqualifying disability, veteran status or any other protected status.",
        font: "Calibri", size: 16, color: STEEL, italics: true
      })]
    }),
  ];
}

// ── POSITION INFO TABLE ───────────────────────────────────────────────────────
function buildPositionSection() {
  const L = 2000, R = PAGE_W - L;
  // How did you hear - checkboxes row
  const hearXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
  ${sdtCheckbox("Employee Referral")}
  <w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> Employee Referral&#x2003;</w:t></w:r>
  ${sdtCheckbox("Newspaper")}
  <w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> Newspaper&#x2003;</w:t></w:r>
  ${sdtCheckbox("Friend")}
  <w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> Friend&#x2003;</w:t></w:r>
  ${sdtCheckbox("Walk In")}
  <w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> Walk In&#x2003;</w:t></w:r>
  ${sdtCheckbox("Other")}
  <w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> Other</w:t></w:r>
</w:p>`;

  return [
    sectionHeader("Position & Application"),
    makeTable([2000, PAGE_W-2000], [
      new TableRow({ height: ROW_H, children: [labelCell("Position Applied For:", 2000), inputCell("Enter position title", PAGE_W-2000, "position")] }),
      new TableRow({ height: ROW_H, children: [labelCell("Today's Date:", 2000), inputCell("MM/DD/YYYY", PAGE_W-2000, "app_date")] }),
    ]),
    makeTable([2000, PAGE_W-2000], [
      new TableRow({ height: ROW_H, children: [
        labelCell("How did you hear\nabout this opening?", 2000),
        new TableCell({
          width: { size: PAGE_W-2000, type: WidthType.DXA },
          borders: allThin,
          shading: { fill: CREAM, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 80 },
          children: [new Paragraph({ xmlData: hearXml })]
        })
      ]}),
    ]),
    makeTable([2000, PAGE_W-2000], [
      new TableRow({ height: ROW_H, children: [labelCell("If Other, explain:", 2000), inputCell("Explain here...", PAGE_W-2000, "how_heard_other")] }),
    ]),
    spacer(160),
  ];
}

// ── PERSONAL INFORMATION ──────────────────────────────────────────────────────
function buildPersonalSection() {
  const nameW = [2200, 2200, 1600, 1500, 1500]; // Last, First, Middle, SSN spacer
  // Row: Last | First | Middle | SSN
  const r1cols = [2300, 2000, 1500, 3200];
  const r2cols = [4500, 2000, 2500];
  return [
    sectionHeader("Personal Information"),
    makeTable(r1cols, [
      new TableRow({ height: ROW_H, children: [
        labelCell("Last Name:", r1cols[0]),
        labelCell("First Name:", r1cols[1]),
        labelCell("Middle Name:", r1cols[2]),
        labelCell("Social Security No.:", r1cols[3]),
      ]}),
      new TableRow({ height: ROW_H, children: [
        inputCell("Last name", r1cols[0], "last_name"),
        inputCell("First name", r1cols[1], "first_name"),
        inputCell("Middle", r1cols[2], "middle_name"),
        inputCell("SSN (XXX-XX-XXXX)", r1cols[3], "ssn"),
      ]}),
    ]),
    makeTable([PAGE_W], [
      new TableRow({ height: LBL_H, children: [labelCell("Home Address:", PAGE_W)] }),
    ]),
    makeTable([3600, 1200, 2000, 1200, 1000], [
      new TableRow({ height: ROW_H, children: [
        inputCell("Street address", 3600, "addr_street"),
        inputCell("Apt #", 1200, "addr_apt"),
        inputCell("City", 2000, "addr_city"),
        inputCell("State", 1200, "addr_state"),
        inputCell("Zip", 1000, "addr_zip"),
      ]}),
    ]),
    makeTable([PAGE_W], [
      new TableRow({ height: LBL_H, children: [labelCell("Previous Address (if lived at above less than 12 months):", PAGE_W)] }),
      new TableRow({ height: ROW_H, children: [inputCell("Previous address", PAGE_W, "prev_address")] }),
    ]),
    makeTable([2400, 2400, 4200], [
      new TableRow({ height: ROW_H, children: [
        labelCell("Home Phone No.:", 2400),
        labelCell("Alternate Phone No.:", 2400),
        labelCell("Emergency Contact (name & relationship):", 4200),
      ]}),
      new TableRow({ height: ROW_H, children: [
        inputCell("(XXX) XXX-XXXX", 2400, "phone_home"),
        inputCell("(XXX) XXX-XXXX", 2400, "phone_alt"),
        inputCell("Name and relationship", 4200, "emergency_contact"),
      ]}),
    ]),
    makeTable([6000, 3000], [
      new TableRow({ height: ROW_H, children: [
        labelCell("Emergency Contact Phone No.:", 6000),
        inputCell("(XXX) XXX-XXXX", 3000, "emergency_phone"),
      ]}),
    ]),
    spacer(120),
    // Yes/No questions
    makeTable([6000, 3000], [
      new TableRow({ height: ROW_H, children: [
        textCell("Are you 18 years of age or older?", 6000),
        yesNoCell("18_yes", "18_no", 3000),
      ]}),
      new TableRow({ height: ROW_H, children: [
        textCell("If under 18, can you submit a work permit after employment?", 6000),
        yesNoCell("work_permit_yes", "work_permit_no", 3000),
      ]}),
      new TableRow({ height: ROW_H, children: [
        textCell("Can you submit verification of your legal right to work in the United States?", 6000),
        yesNoCell("work_auth_yes", "work_auth_no", 3000),
      ]}),
      new TableRow({ height: ROW_H, children: [
        textCell("Do you have adequate means of transportation to get to work on time?", 6000),
        yesNoCell("transport_yes", "transport_no", 3000),
      ]}),
      new TableRow({ height: ROW_H, children: [
        textCell("Are you able to perform the essential functions of the position you are applying for?", 6000),
        yesNoCell("essential_yes", "essential_no", 3000),
      ]}),
    ]),
    makeTable([PAGE_W], [
      new TableRow({ height: LBL_H, children: [labelCell("If no to essential functions, explain:", PAGE_W)] }),
      new TableRow({ height: ROW_H, children: [inputCell("Explanation...", PAGE_W, "essential_explain")] }),
    ]),
    spacer(120),
    // Conviction questions
    makeTable([6000, 3000], [
      new TableRow({ height: ROW_H, children: [
        textCell("Have you ever been convicted of a felony?", 6000),
        yesNoCell("felony_yes", "felony_no", 3000),
      ]}),
      new TableRow({ height: ROW_H, children: [
        textCell("Have you ever been convicted of a misdemeanor involving theft?", 6000),
        yesNoCell("theft_yes", "theft_no", 3000),
      ]}),
      new TableRow({ height: ROW_H, children: [
        textCell("Have you ever been convicted of a misdemeanor involving abuse, neglect, or mistreatment?", 6000),
        yesNoCell("abuse_conv_yes", "abuse_conv_no", 3000),
      ]}),
    ]),
    new Paragraph({ spacing: { before: 60, after: 40 }, children: [new TextRun({ text: "A conviction will not necessarily disqualify you. If you answered Yes to any of the above, list the convictions below:", font: "Calibri", size: 17, italics: true, color: STEEL })] }),
    makeTable([3000, 1800, 2000, 2200], [
      new TableRow({ height: LBL_H, children: [labelCell("Offense", 3000), labelCell("Date", 1800), labelCell("Place", 2000), labelCell("Disposition", 2200)] }),
      new TableRow({ height: ROW_H, children: [inputCell("Offense description", 3000, "offense1"), inputCell("Date", 1800, "off1_date"), inputCell("Place", 2000, "off1_place"), inputCell("Disposition", 2200, "off1_disp")] }),
      new TableRow({ height: ROW_H, children: [inputCell("Offense description", 3000, "offense2"), inputCell("Date", 1800, "off2_date"), inputCell("Place", 2000, "off2_place"), inputCell("Disposition", 2200, "off2_disp")] }),
    ]),
    spacer(160),
  ];
}

// ── EDUCATION ────────────────────────────────────────────────────────────────
function buildEducationSection() {
  const cols = [3800, 1800, 2000, 1400];
  function eduRow(label, tag) {
    return new TableRow({ height: ROW_H, children: [
      inputCell(label, cols[0], `${tag}_name`),
      inputCell("Years", cols[1], `${tag}_years`),
      inputCell("Degree", cols[2], `${tag}_degree`),
      inputCell("Year", cols[3], `${tag}_grad`),
    ]});
  }
  return [
    sectionHeader("Education"),
    makeTable(cols, [
      new TableRow({ height: LBL_H, children: [
        labelCell("Name & Address of School Attended", cols[0]),
        labelCell("Year(s) of Attendance", cols[1]),
        labelCell("Degree Obtained", cols[2]),
        labelCell("Year Graduated", cols[3]),
      ]}),
      eduRow("Last High School — name & address", "hs"),
      eduRow("Last College / University or Nursing School", "college"),
      eduRow("Graduate, Technical or Vocational School", "grad"),
      eduRow("Courses currently enrolled in", "current"),
    ]),
    spacer(160),
  ];
}

// ── LICENSES & CERTIFICATIONS ────────────────────────────────────────────────
function buildLicensesSection() {
  const cols4 = [2700, 1500, 1800, 1500, 1500];
  function licRow(typeTag) {
    return new TableRow({ height: ROW_H, children: [
      inputCell("License/certification type", cols4[0], `${typeTag}_type`),
      inputCell("State", cols4[1], `${typeTag}_state`),
      inputCell("Date Issued", cols4[2], `${typeTag}_issued`),
      inputCell("Expires On", cols4[3], `${typeTag}_expires`),
      inputCell("Number", cols4[4], `${typeTag}_number`),
    ]});
  }
  // CPR + ACLS row
  const cprXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
  <w:r><w:rPr><w:sz w:val="18"/><w:b/></w:rPr><w:t xml:space="preserve">CPR: </w:t></w:r>
  ${sdtCheckbox("CPR_yes")}
  <w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> Yes&#x2003;</w:t></w:r>
  ${sdtCheckbox("CPR_no")}
  <w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> No&#x2003;Expires: </w:t></w:r>
  ${sdtField("MM/YYYY", 1000, "cpr_expires")}
</w:p>`;
  const aclsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
  <w:r><w:rPr><w:sz w:val="18"/><w:b/></w:rPr><w:t xml:space="preserve">ACLS: </w:t></w:r>
  ${sdtCheckbox("ACLS_yes")}
  <w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> Yes&#x2003;</w:t></w:r>
  ${sdtCheckbox("ACLS_no")}
  <w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> No&#x2003;Expires: </w:t></w:r>
  ${sdtField("MM/YYYY", 1000, "acls_expires")}
</w:p>`;

  return [
    sectionHeader("Professional Licenses & Certifications"),
    makeTable(cols4, [
      new TableRow({ height: LBL_H, children: [
        labelCell("License / Certification Type", cols4[0]),
        labelCell("State Issued", cols4[1]),
        labelCell("Date Issued", cols4[2]),
        labelCell("Expires On", cols4[3]),
        labelCell("License No.", cols4[4]),
      ]}),
      licRow("lic1"),
      licRow("lic2"),
    ]),
    spacer(60),
    makeTable([PAGE_W/2, PAGE_W/2], [
      new TableRow({ height: ROW_H, children: [
        new TableCell({ width: { size: PAGE_W/2, type: WidthType.DXA }, borders: allThin, shading: { fill: CREAM, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 80 }, children: [new Paragraph({ xmlData: cprXml })] }),
        new TableCell({ width: { size: PAGE_W/2, type: WidthType.DXA }, borders: allThin, shading: { fill: CREAM, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 80 }, children: [new Paragraph({ xmlData: aclsXml })] }),
      ]}),
    ]),
    makeTable([PAGE_W/2, PAGE_W/2], [
      new TableRow({ height: ROW_H, children: [labelCell("Other Certifications:", PAGE_W/2), labelCell("Other Special Skills:", PAGE_W/2)] }),
      new TableRow({ height: ROW_H, children: [inputCell("List other certifications", PAGE_W/2, "other_certs"), inputCell("List special skills", PAGE_W/2, "other_skills")] }),
    ]),
    makeTable([3600, 5400], [
      new TableRow({ height: ROW_H, children: [
        textCell("Do you have malpractice insurance?", 3600),
        yesNoCell("malpractice_yes", "malpractice_no", 5400),
      ]}),
      new TableRow({ height: ROW_H, children: [labelCell("If Yes — Policy Name:", 3600), inputCell("Policy name", 5400, "malpractice_policy_name")] }),
      new TableRow({ height: ROW_H, children: [labelCell("If Yes — Policy No.:", 3600), inputCell("Policy number", 5400, "malpractice_policy_no")] }),
    ]),
    spacer(160),
  ];
}

// ── AVAILABILITY ─────────────────────────────────────────────────────────────
function buildAvailabilitySection() {
  const dayW = Math.floor(PAGE_W / 8);
  const shiftW = Math.floor(PAGE_W / 6);

  // Days checkboxes
  const days = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"];
  const dayXml = days.map(d =>
    `${sdtCheckbox(d)}<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> ${d}&#x2003;</w:t></w:r>`
  ).join('\n');
  const daysParaXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
  ${dayXml}
</w:p>`;

  const holidayXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
  ${sdtCheckbox("Holidays")}
  <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> Holidays</w:t></w:r>
</w:p>`;

  const shifts = [["7am–3pm","shift_7_3"],["3pm–11pm","shift_3_11"],["11pm–7am","shift_11_7"],["7am–7pm","shift_7a_7p"],["7pm–7am","shift_7p_7a"],["Other","shift_other"]];
  const shiftXml = shifts.map(([lbl,tag]) =>
    `${sdtCheckbox(tag)}<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve"> ${lbl}&#x2003;</w:t></w:r>`
  ).join('\n');
  const shiftParaXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
  ${shiftXml}
</w:p>`;

  return [
    sectionHeader("Availability"),
    makeTable([2200, PAGE_W-2200], [
      new TableRow({ height: ROW_H, children: [labelCell("Date Available to Start:", 2200), inputCell("MM/DD/YYYY", PAGE_W-2200, "start_date")] }),
    ]),
    makeTable([PAGE_W], [
      new TableRow({ height: LBL_H, children: [labelCell("Days Available for Assignment (check all that apply):", PAGE_W)] }),
      new TableRow({ height: ROW_H, children: [
        new TableCell({
          width: { size: PAGE_W, type: WidthType.DXA },
          borders: allThin,
          shading: { fill: CREAM, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 80 },
          children: [new Paragraph({ xmlData: daysParaXml }), new Paragraph({ xmlData: holidayXml })]
        })
      ]}),
    ]),
    makeTable([PAGE_W], [
      new TableRow({ height: LBL_H, children: [labelCell("Shifts Available (check all that apply):", PAGE_W)] }),
      new TableRow({ height: ROW_H, children: [
        new TableCell({
          width: { size: PAGE_W, type: WidthType.DXA },
          borders: allThin,
          shading: { fill: CREAM, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 80 },
          children: [new Paragraph({ xmlData: shiftParaXml })]
        })
      ]}),
    ]),
    spacer(160),
  ];
}

// ── EMPLOYMENT HISTORY ────────────────────────────────────────────────────────
function buildEmploymentSection() {
  function empBlock(n) {
    const tag = `emp${n}`;
    return [
      new TableRow({ height: LBL_H, children: [
        new TableCell({
          width: { size: PAGE_W, type: WidthType.DXA },
          columnSpan: 6,
          borders: { top: noBorder, left: noBorder, right: noBorder, bottom: { style: BorderStyle.SINGLE, size: 6, color: STEEL } },
          shading: { fill: LTBLUE, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 80 },
          children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: `Employer ${n}`, font: "Calibri", size: 19, bold: true, color: NAVY })] })]
        })
      ]}),
      new TableRow({ height: ROW_H, children: [
        labelCell("Employer:", 2000),
        inputCell("Company name", 3200, `${tag}_name`),
        labelCell("Position:", 1400),
        inputCell("Job title", 2400, `${tag}_position`),
      ]}),
      new TableRow({ height: ROW_H, children: [
        labelCell("Address:", 2000),
        inputCell("Street, City, State, Zip", 3200, `${tag}_address`),
        labelCell("Phone No.:", 1400),
        inputCell("(XXX) XXX-XXXX", 2400, `${tag}_phone`),
      ]}),
      new TableRow({ height: ROW_H, children: [
        labelCell("Employed From:", 2000),
        inputCell("MM/YYYY", 1600, `${tag}_from`),
        labelCell("To:", 800),
        inputCell("MM/YYYY or Present", 1600, `${tag}_to`),
        labelCell("Starting $:", 1500),
        inputCell("Amount", 1500, `${tag}_start_wage`),
      ]}),
      new TableRow({ height: ROW_H, children: [
        labelCell("Supervisor Name & Title:", 2000),
        inputCell("Name and title", 3200, `${tag}_supervisor`),
        labelCell("Final $:", 1400),
        inputCell("Amount", 2400, `${tag}_final_wage`),
      ]}),
      new TableRow({ height: { value: 600, rule: "atLeast" }, children: [
        labelCell("Duties:", 2000),
        new TableCell({ width: { size: 7000, type: WidthType.DXA }, columnSpan: 3, borders: allThin, shading: { fill: CREAM, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 80 }, children: [
          new Paragraph({ xmlData: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:before="60" w:after="60"/><w:ind w:left="60"/></w:pPr>${sdtField("Briefly describe your duties", 7000, `${tag}_duties`)}</w:p>` })
        ]}),
      ]}),
      new TableRow({ height: ROW_H, children: [
        labelCell("Reason for Leaving:", 2000),
        new TableCell({ width: { size: 7000, type: WidthType.DXA }, columnSpan: 3, borders: allThin, shading: { fill: CREAM, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 80 }, children: [
          new Paragraph({ xmlData: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:before="60" w:after="60"/><w:ind w:left="60"/></w:pPr>${sdtField("Reason for leaving", 7000, `${tag}_reason`)}</w:p>` })
        ]}),
      ]}),
    ];
  }

  return [
    sectionHeader("Employment History"),
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "List your most recent employer first. Account for any gaps. Include at least 3 employers or 5 years of history. Use additional paper if necessary.", font: "Calibri", size: 17, italics: true, color: STEEL })] }),
    makeTable([2000, 2800, 1400, 1600, 1300, 1000], [
      ...empBlock(1),
      new TableRow({ height: { value: 200, rule: "exact" }, children: [new TableCell({ width: { size: PAGE_W, type: WidthType.DXA }, columnSpan: 6, borders: noAllBorder, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [] })] })] }),
      ...empBlock(2),
      new TableRow({ height: { value: 200, rule: "exact" }, children: [new TableCell({ width: { size: PAGE_W, type: WidthType.DXA }, columnSpan: 6, borders: noAllBorder, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [] })] })] }),
      ...empBlock(3),
    ]),
    spacer(160),
  ];
}

// ── REFERENCES ────────────────────────────────────────────────────────────────
function buildReferencesSection() {
  const cols = [3600, 2400, 3000];
  function refRow(n) {
    return new TableRow({ height: ROW_H, children: [
      inputCell(`Reference ${n} — Name & Title`, cols[0], `ref${n}_name`),
      inputCell("Relationship", cols[1], `ref${n}_rel`),
      inputCell("(XXX) XXX-XXXX", cols[2], `ref${n}_phone`),
    ]});
  }
  return [
    sectionHeader("References"),
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "List persons not related to you whom you have known for at least two years and who we may contact for character references.", font: "Calibri", size: 17, italics: true, color: STEEL })] }),
    makeTable(cols, [
      new TableRow({ height: LBL_H, children: [labelCell("Name & Title", cols[0]), labelCell("Relationship", cols[1]), labelCell("Phone Number", cols[2])] }),
      refRow(1), refRow(2), refRow(3),
    ]),
    spacer(160),
  ];
}

// ── CERTIFICATION / SIGNATURE ─────────────────────────────────────────────────
function buildSignatureSection() {
  const disclosureText = [
    "I understand that employment at Integrated Sycamore Group LLC is at-will and that either party may terminate the employment relationship at any time, for any reason, with or without notice.",
    "I understand that Integrated Sycamore Group LLC reserves the right to require employees to submit to blood tests or urinalysis for alcohol or drug screening. Refusal may result in termination.",
    "I authorize Integrated Sycamore Group LLC to contact any references and former employers listed herein. I release all parties from any liability arising from such inquiry.",
    "I certify that all statements herein are true. I understand that any false statement, unsatisfactory references, failed drug screening, failed background check, or inability to perform essential job functions may result in discipline up to and including discharge.",
  ];

  return [
    sectionHeader("Certification & Signature"),
    new Paragraph({ spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "PLEASE READ CAREFULLY BEFORE SIGNING:", font: "Calibri", size: 18, bold: true, color: NAVY })] }),
    ...disclosureText.map((t, i) => new Paragraph({
      spacing: { before: 60, after: 60 },
      indent: { left: 200 },
      children: [new TextRun({ text: `${i+1}. ${t}`, font: "Calibri", size: 17, color: BLACK })]
    })),
    spacer(120),
    makeTable([5400, 3600], [
      new TableRow({ height: ROW_H, children: [labelCell("Applicant Signature:", 5400), inputCell("Sign here", 3600, "signature")] }),
      new TableRow({ height: ROW_H, children: [labelCell("Date:", 5400), inputCell("MM/DD/YYYY", 3600, "sig_date")] }),
    ]),
    spacer(200),
  ];
}

// ── LETTER OF REFERENCE ───────────────────────────────────────────────────────
function buildRefLetterSection(num) {
  return [
    sectionHeader(`Letter of Reference — #${num}`),
    new Paragraph({ spacing: { before: 80, after: 60 }, children: [new TextRun({ text: "Please have your referee complete this letter and return it with your application.", font: "Calibri", size: 17, italics: true, color: STEEL })] }),
    makeTable([2200, PAGE_W-2200], [
      new TableRow({ height: ROW_H, children: [labelCell("Referee Name:", 2200), inputCell("Full name", PAGE_W-2200, `ref${num}_referee_name`)] }),
      new TableRow({ height: ROW_H, children: [labelCell("Complete Address:", 2200), inputCell("Street, City, State, Zip", PAGE_W-2200, `ref${num}_referee_addr`)] }),
      new TableRow({ height: ROW_H, children: [labelCell("Phone Number:", 2200), inputCell("(XXX) XXX-XXXX", PAGE_W-2200, `ref${num}_referee_phone`)] }),
    ]),
    spacer(80),
    makeTable([PAGE_W], [
      new TableRow({ height: { value: 500, rule: "atLeast" }, children: [
        new TableCell({ width: { size: PAGE_W, type: WidthType.DXA }, borders: allThin, shading: { fill: CREAM, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 200, right: 200 }, children: [
          new Paragraph({ spacing: { before: 60, after: 60 }, children: [
            new TextRun({ text: "My name is ", font: "Calibri", size: 18 }),
          ]}),
          new Paragraph({ xmlData: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">I am writing to recommend </w:t></w:r>${sdtField("Applicant name", PAGE_W-400, `ref${num}_applicant`)}<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">.</w:t></w:r></w:p>` }),
          new Paragraph({ xmlData: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">I have known her/him for </w:t></w:r>${sdtField("Number", 800, `ref${num}_years`)}<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> years. She/He is not related to me by blood.</w:t></w:r></w:p>` }),
          new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "She/He is capable of providing a safe and healthy environment for the clients or individuals being served.", font: "Calibri", size: 18 })] }),
          new Paragraph({ xmlData: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">For further information I can be reached at </w:t></w:r>${sdtField("Phone / email", 2000, `ref${num}_contact`)}<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">.</w:t></w:r></w:p>` }),
        ]})
      ]}),
    ]),
    makeTable([4500, 4500], [
      new TableRow({ height: ROW_H, children: [labelCell("Referee Signature:", 4500), inputCell("Signature", 4500, `ref${num}_sig`)] }),
      new TableRow({ height: ROW_H, children: [labelCell("Print Name:", 4500), inputCell("Printed name", 4500, `ref${num}_print`)] }),
      new TableRow({ height: ROW_H, children: [labelCell("Date:", 4500), inputCell("MM/DD/YYYY", 4500, `ref${num}_date`)] }),
    ]),
    spacer(200),
  ];
}

// ── EVALUATION ───────────────────────────────────────────────────────────────
const evalQuestions = [
  ["Explain the philosophy of Home and Community Based Services (HCS).", "hcs_philosophy"],
  ["Based on your understanding, list your job responsibilities.", "job_responsibilities"],
  ["How would you respond to a medical emergency situation? Explain.", "medical_emergency"],
  ["What would you do if your shift ends and no one shows up to take over from you?", "shift_end"],
  ["What do you understand by 'Consumer Confidentiality' (HIPAA Law) and what would you do to protect it?", "hipaa"],
  ["What would you do if a consumer fails to eat because he/she does not like what was served?", "consumer_eating"],
  ["What do you understand by Abuse, Neglect and Exploitation?", "abuse_neglect"],
  ["What are you obliged to do if you suspect that a consumer has been abused?", "abuse_report"],
  ["How would you assist in promoting your consumer's self-esteem?", "self_esteem"],
];

function buildEvalSection() {
  function evalRow(question, tag) {
    return [
      makeTable([PAGE_W], [
        new TableRow({ height: LBL_H, children: [labelCell(question, PAGE_W)] }),
        new TableRow({ height: { value: 900, rule: "atLeast" }, children: [
          new TableCell({ width: { size: PAGE_W, type: WidthType.DXA }, borders: allThin, shading: { fill: CREAM, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 80 }, children: [
            new Paragraph({ xmlData: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:spacing w:before="60" w:after="60"/><w:ind w:left="60"/></w:pPr>${sdtField("Enter your response here...", PAGE_W, tag)}</w:p>` })
          ]})
        ]}),
      ]),
      spacer(80),
    ];
  }

  return [
    sectionHeader("Evaluation — To Be Completed by Applicant"),
    new Paragraph({ spacing: { before: 80, after: 120 }, children: [new TextRun({ text: `Employee's Name: `, font: "Calibri", size: 18, bold: true, color: NAVY })] }),
    makeTable([2200, PAGE_W - 2200], [
      new TableRow({ height: ROW_H, children: [labelCell("Employee Name:", 2200), inputCell("Last, First, M.I.", PAGE_W-2200, "eval_emp_name")] }),
    ]),
    spacer(80),
    ...evalQuestions.flatMap(([q, t]) => evalRow(q, t)),
    makeTable([5400, 3600], [
      new TableRow({ height: ROW_H, children: [labelCell("Applicant Signature:", 5400), inputCell("Signature", 3600, "eval_sig")] }),
      new TableRow({ height: ROW_H, children: [labelCell("Date:", 5400), inputCell("MM/DD/YYYY", 3600, "eval_date")] }),
    ]),
    makeTable([5400, 3600], [
      new TableRow({ height: ROW_H, children: [labelCell("Reviewed By:", 5400), inputCell("Reviewer name", 3600, "reviewed_by")] }),
      new TableRow({ height: ROW_H, children: [labelCell("Date Reviewed:", 5400), inputCell("MM/DD/YYYY", 3600, "reviewed_date")] }),
      new TableRow({ height: { value: 700, rule: "atLeast" }, children: [labelCell("Remarks:", 5400), inputCell("Remarks...", 3600, "remarks")] }),
    ]),
    spacer(200),
  ];
}

// ── ANE ACKNOWLEDGEMENT ───────────────────────────────────────────────────────
function buildANESection() {
  return [
    sectionHeader("Abuse, Neglect & Exploitation — Acknowledgement"),
    new Paragraph({ spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "TEXAS LEGISLATURE BILL 2170", font: "Calibri", size: 18, bold: true, color: NAVY })] }),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Integrated Sycamore Group LLC strives to provide the best possible services to its residents in a comfortable atmosphere.", font: "Calibri", size: 17 })] }),
    new Paragraph({ spacing: { before: 60, after: 80 }, children: [new TextRun({ text: "If you suspect that any resident in our program is being abused, neglected, or exploited, please call and report immediately to:", font: "Calibri", size: 17 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Texas Department of Protective and Regulatory Services (TDPRS)", font: "Calibri", size: 20, bold: true, color: NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Within ONE HOUR  —  1-800-647-7418", font: "Calibri", size: 22, bold: true, color: STEEL })] }),
    new Paragraph({ spacing: { before: 60, after: 80 }, children: [new TextRun({ text: "This acknowledges that I have been informed of how to report allegations of abuse, neglect, or exploitation to the appropriate regulatory agency, and have been provided with the DFPS toll-free telephone number: 1-800-647-7418.", font: "Calibri", size: 17 })] }),
    makeTable([3000, 3000, 3000], [
      new TableRow({ height: ROW_H, children: [
        labelCell("Consumer / Staff Signature:", 3000),
        labelCell("Legally Authorized Representative:", 3000),
        labelCell("Company Representative:", 3000),
      ]}),
      new TableRow({ height: ROW_H, children: [
        inputCell("Signature", 3000, "ane_consumer_sig"),
        inputCell("Signature", 3000, "ane_lar_sig"),
        inputCell("Signature", 3000, "ane_company_sig"),
      ]}),
      new TableRow({ height: ROW_H, children: [
        inputCell("MM/DD/YYYY", 3000, "ane_consumer_date"),
        inputCell("MM/DD/YYYY", 3000, "ane_lar_date"),
        inputCell("MM/DD/YYYY", 3000, "ane_company_date"),
      ]}),
    ]),
    spacer(200),
  ];
}

// ── COMPLAINT PROCEDURE ───────────────────────────────────────────────────────
function buildComplaintSection() {
  const steps = [
    "Present a written and/or verbal complaint to an administrator.",
    "Complaints by consumers must be reviewed by the Case Manager. Complaints by staff or administrators must be presented to the Program Director.",
    "Complaints will be reviewed by the Director within 10 days of receiving date (may include a meeting with persons involved).",
    "Director must initiate a solution or resolution within 15 days of receiving date.",
    "Director will meet with the person making the complaint or respond in writing within 20 days of receiving date.",
  ];
  const contacts = [
    ["Program Manager", "832-335-4098"],
    ["Nurse", "713-882-0161"],
    ["Program Director", "832-335-4098"],
    ["DADS (if unresolved)", "1-800-458-9858"],
  ];

  return [
    sectionHeader("Complaint Procedure"),
    new Paragraph({ spacing: { before: 100, after: 60 }, children: [new TextRun({ text: "To All Consumers / Staff / Legally Authorized Representatives:", font: "Calibri", size: 18, bold: true, color: NAVY })] }),
    ...steps.map((s, i) => new Paragraph({
      spacing: { before: 60, after: 60 },
      indent: { left: 200, hanging: 200 },
      children: [new TextRun({ text: `${i + 1}.  ${s}`, font: "Calibri", size: 17, color: BLACK })]
    })),
    new Paragraph({ spacing: { before: 80, after: 60 }, children: [new TextRun({ text: "NOTE: If you feel your complaint was not handled professionally or satisfactorily, you are encouraged to report to the DADS office at 1-800-458-9858.", font: "Calibri", size: 17, italics: true, color: STEEL })] }),
    spacer(80),
    makeTable([3600, 5400], [
      new TableRow({ height: LBL_H, children: [
        labelCell("Contact", 3600, { bold: true }),
        labelCell("Phone Number", 5400, { bold: true }),
      ]}),
      ...contacts.map(([name, phone]) => new TableRow({ height: ROW_H, children: [
        textCell(name, 3600),
        textCell(phone, 5400),
      ]})),
    ]),
    spacer(120),
    makeTable([4500, 4500], [
      new TableRow({ height: ROW_H, children: [labelCell("Residential Manager / Provider Rep. Signature:", 4500), inputCell("Signature", 4500, "complaint_mgr_sig")] }),
      new TableRow({ height: ROW_H, children: [labelCell("Program Director Signature:", 4500), inputCell("Signature", 4500, "complaint_dir_sig")] }),
    ]),
    spacer(200),
  ];
}

// ── FOOTER ────────────────────────────────────────────────────────────────────
const docFooter = new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: "Integrated Sycamore Group LLC  •  8315 Sierra Hill Ct, Houston TX 77083  •  Page ", font: "Calibri", size: 16, color: STEEL }),
      new TextRun({ children: [PageNumber.CURRENT], font: "Calibri", size: 16, color: STEEL }),
      new TextRun({ text: " of ", font: "Calibri", size: 16, color: STEEL }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Calibri", size: 16, color: STEEL }),
    ]
  })]
});

// ─── ASSEMBLE ─────────────────────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1260 }
      }
    },
    footers: { default: docFooter },
    children: [
      ...buildHeader(),
      ...buildPositionSection(),
      ...buildPersonalSection(),
      ...buildEducationSection(),
      ...buildLicensesSection(),
      ...buildAvailabilitySection(),
      ...buildEmploymentSection(),
      ...buildReferencesSection(),
      ...buildSignatureSection(),
      new Paragraph({ children: [new PageBreak()] }),
      ...buildRefLetterSection(1),
      ...buildRefLetterSection(2),
      ...buildRefLetterSection(3),
      new Paragraph({ children: [new PageBreak()] }),
      ...buildEvalSection(),
      new Paragraph({ children: [new PageBreak()] }),
      ...buildANESection(),
      ...buildComplaintSection(),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/home/claude/Employment_APPLICATION_Fillable_v2.docx", buf);
  console.log("Done");
});
