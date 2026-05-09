# Fillable Word Form Automation

Converted a static 16-page employment application into a fully interactive fillable Word document — using Python and Node.js to inject Word Content Controls (SDT fields and checkboxes) directly into Office Open XML.

Built as a real solution for a care agency client who was considering an expensive website. This approach achieved the same result for free and can be emailed directly to applicants.

---

## What it does

- Takes a standard `.docx` file and rebuilds it as a professional fillable form
- Applicants click each field, type their answer, and email it back — no printing required
- 130 form controls injected: text fields for every input area, clickable checkboxes for Yes/No questions, days available, shifts, and more
- Clean layout with navy section headers, consistent label/input styling, and a footer on every page

---

## How it was built

### Stage 1 — Python XML approach (`build_fillable.py`)
The first attempt used Python to patch the original document by finding underscore placeholders and empty table cells, then replacing them with SDT Content Controls injected via `lxml`.

```python
import zipfile, re
from lxml import etree

# Unzip the docx, parse document.xml, find blank runs, replace with SDT fields
with zipfile.ZipFile("Employment_APPLICATION.docx") as z:
    doc_xml = z.read("word/document.xml")

tree = etree.fromstring(doc_xml)

for r in tree.iter("{...}r"):
    run_text = ''.join(t.text or '' for t in r.findall("{...}t"))
    if re.match(r'^\s*_{5,}\s*$', run_text):
        # Replace with SDT text field
        parent = r.getparent()
        parent.insert(list(parent).index(r), make_text_sdt("Click to enter"))
        parent.remove(r)
```

**Problem:** The original document had broken table structure — narrow cells that caused text to overflow. Patching it made the overflow worse.

### Stage 2 — Full rebuild with docx.js (`build_form.mjs`)
Rather than patching a broken document, the form was rebuilt completely from scratch using `docx` (Node.js library), with SDT fields injected as raw XML via the `xmlData` escape hatch.

```javascript
import { Document, Packer, Table, TableRow, TableCell, Paragraph } from 'docx';

// SDT text field template
function sdtField(placeholder, tag) {
  return `<w:sdt xmlns:w="...">
    <w:sdtPr>
      <w:alias w:val="${placeholder}"/>
      <w:tag w:val="${tag}"/>
      <w:showingPlcHdr/>
      <w:text/>
    </w:sdtPr>
    <w:sdtContent>
      <w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>
        <w:t>${placeholder}</w:t>
      </w:r>
    </w:sdtContent>
  </w:sdt>`;
}

// W14 checkbox template
function sdtCheckbox(label) {
  return `<w:sdt xmlns:w="..." xmlns:w14="...">
    <w:sdtPr>
      <w14:checkbox>
        <w14:checked w14:val="0"/>
        <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
        <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
      </w14:checkbox>
    </w:sdtPr>
    <w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent>
  </w:sdt>`;
}
```

Each input cell wraps an SDT field inside a `Paragraph` with `xmlData`, which docx.js passes through raw into the final XML.

---

## Project structure

```
fillable-word-form/
├── build_fillable.py       # Stage 1: Python patcher (lxml approach)
├── build_form.mjs          # Stage 2: Full rebuild (docx.js + raw XML SDTs)
├── Employment_APPLICATION_Fillable_v2.docx   # Final output
└── README.md
```

---

## Tech stack

| Tool | Purpose |
|------|---------|
| Python | Stage 1 XML patching, zipfile handling |
| lxml | Parsing and editing Office Open XML |
| Node.js | Running the docx.js rebuild script |
| docx (npm) | Document structure — tables, paragraphs, styles |
| Office Open XML (OOXML) | SDT content controls, W14 checkboxes |
| zipfile | Reading and writing `.docx` archives |

---

## Key things learned

**Office Open XML structure** — a `.docx` file is just a ZIP archive containing XML files. `word/document.xml` holds all the content; you can unzip, edit the XML directly, and rezip.

**SDT Content Controls** — Word's form fields are `<w:sdt>` elements (Structured Document Tags). They need a `<w:sdtPr>` block for properties and a `<w:sdtContent>` block for display content. The `<w:showingPlcHdr/>` tag makes the placeholder text show until the user types.

**W14 Checkboxes** — clickable checkboxes use a different namespace (`w14`) and require the `MS Gothic` font for the check/uncheck characters (`✒` = 0x2612, `☐` = 0x2610).

**docx.js `xmlData`** — docx.js doesn't natively support SDT fields, but you can pass raw XML strings via the `xmlData` property on a `Paragraph`. The library passes it straight through into the output file.

**When to patch vs rebuild** — patching works if the source document has clean structure. If the original has layout problems (like overlapping cells or broken column widths), rebuilding from scratch produces a far better result.

---

## How to run

```bash
# Install dependencies
npm install -g docx

# Run the full rebuild
node build_form.mjs

# Output: Employment_APPLICATION_Fillable_v2.docx
```

For the Python stage 1 version:

```bash
pip install lxml
python build_fillable.py
```

---

## Output

The final `.docx` opens in Microsoft Word. Every field is clickable — applicants tab through the form, type into each box, tick the relevant checkboxes, and email it back. No printing, no website, no cost.
