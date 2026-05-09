"""
Build a fillable Word form from Employment_APPLICATION.docx
Uses SDT content controls (text fields + checkboxes) throughout.
"""
import re, copy, zipfile, os, shutil
from lxml import etree

SRC = "/mnt/user-data/uploads/Employment_APPLICATION.docx"
DST = "/mnt/user-data/outputs/Employment_APPLICATION_Fillable.docx"

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
WW = '{' + W + '}'
WW14 = '{' + W14 + '}'

_id_counter = [2000]
def nid():
    _id_counter[0] += 1
    return str(_id_counter[0])

TEXT_SDT_TMPL = '''<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:sdtPr>
    <w:alias w:val="{alias}"/>
    <w:tag w:val="{tag}"/>
    <w:id w:val="{sid}"/>
    <w:showingPlcHdr/>
    <w:text/>
  </w:sdtPr>
  <w:sdtContent>
    <w:r>
      <w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>
      <w:t>{alias}</w:t>
    </w:r>
  </w:sdtContent>
</w:sdt>'''

CHECK_SDT_TMPL = '''<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:sdtPr>
    <w:alias w:val="{alias}"/>
    <w:tag w:val="{tag}"/>
    <w:id w:val="{sid}"/>
    <w14:checkbox>
      <w14:checked w14:val="0"/>
      <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
      <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
    </w14:checkbox>
  </w:sdtPr>
  <w:sdtContent>
    <w:r><w:t>&#x2610;</w:t></w:r>
  </w:sdtContent>
</w:sdt>'''

def text_sdt(alias, tag=None):
    tag = tag or re.sub(r'[^\w]', '_', alias.lower())[:40]
    return etree.fromstring(TEXT_SDT_TMPL.format(alias=alias, tag=tag, sid=nid()))

def check_sdt(alias):
    tag = re.sub(r'[^\w]', '_', alias.lower())[:40]
    return etree.fromstring(CHECK_SDT_TMPL.format(alias=alias, tag=tag, sid=nid()))

def get_text(el):
    return ''.join((t.text or '') for t in el.iter(WW+'t'))

def process_xml(xml_bytes):
    tree = etree.fromstring(xml_bytes)
    
    # 1. Replace underscore-only runs with text SDT
    for r in list(tree.iter(WW+'r')):
        # Don't modify runs inside existing SDTs
        if any(a.tag == WW+'sdt' for a in r.iterancestors()):
            continue
        t_els = r.findall(WW+'t')
        if not t_els:
            continue
        run_text = ''.join((t.text or '') for t in t_els)
        if re.match(r'^\s*_{5,}\s*$', run_text):
            parent = r.getparent()
            if parent is None:
                continue
            idx = list(parent).index(r)
            sdt = text_sdt('Type here')
            parent.remove(r)
            parent.insert(idx, sdt)
    
    # 2. Add text SDTs to empty table cells (input area cells)
    for tc in tree.iter(WW+'tc'):
        if any(a.tag == WW+'sdt' for a in tc.iterancestors()):
            continue
        cell_text = get_text(tc).strip()
        
        # Skip cells that are labels (have meaningful label text ending with colon or known labels)
        if cell_text and len(cell_text) > 2:
            continue  # has content, skip
        
        # Empty cell - add input field  
        paras = tc.findall(WW+'p')
        if not paras:
            continue
        p = paras[0]
        # Only add if no existing SDT or runs with real text
        if p.find(WW+'sdt') is not None:
            continue
        existing_text = get_text(p).strip()
        if existing_text:
            continue
        sdt = text_sdt('Click to enter')
        p.append(sdt)
    
    # 3. Handle Yes/No checkboxes: find patterns like "_Yes" "_No" and replace with checkboxes
    # We process at the run level looking for checkbox marker text
    for r in list(tree.iter(WW+'r')):
        if any(a.tag == WW+'sdt' for a in r.iterancestors()):
            continue
        t_els = r.findall(WW+'t')
        if not t_els:
            continue
        run_text = ''.join((t.text or '') for t in t_els)
        
        # Patterns like "  _Yes  _No" or "_Yes _No" - we want to keep the labels but make checkboxes
        # Replace _Yes with checkbox+Yes and _No with checkbox+No in the run text
        if '_Yes' in run_text or '_No' in run_text:
            parent = r.getparent()
            if parent is None:
                continue
            idx = list(parent).index(r)
            
            # Split run into parts, inserting checkboxes
            parts = re.split(r'(_Yes|_No)', run_text)
            new_elements = []
            for part in parts:
                if part == '_Yes':
                    cb = check_sdt('Yes')
                    lbl = etree.fromstring('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t xml:space="preserve"> Yes </w:t></w:r>')
                    new_elements.extend([cb, lbl])
                elif part == '_No':
                    cb = check_sdt('No')
                    lbl = etree.fromstring('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t xml:space="preserve"> No </w:t></w:r>')
                    new_elements.extend([cb, lbl])
                elif part.strip():
                    el = etree.fromstring(f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t xml:space="preserve">{part}</w:t></w:r>')
                    new_elements.append(el)
            
            parent.remove(r)
            for i, el in enumerate(new_elements):
                parent.insert(idx + i, el)
    
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)

# Copy original zip, process document.xml
shutil.copy(SRC, DST)
with zipfile.ZipFile(DST, 'r') as z:
    doc_xml = z.read('word/document.xml')

new_doc_xml = process_xml(doc_xml)

# Write back to zip
import tempfile
tmp = DST + '.tmp'
with zipfile.ZipFile(DST, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename == 'word/document.xml':
            zout.writestr(item, new_doc_xml)
        else:
            zout.writestr(item, zin.read(item.filename))

os.replace(tmp, DST)
print("Done! Saved to", DST)
