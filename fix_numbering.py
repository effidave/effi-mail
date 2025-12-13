"""
Post-process the PDX DSA document to fix multilevel numbering levels.
Parses paragraph text to detect ordinal patterns and sets appropriate ilvl values.
"""

import zipfile
import os
import re
import shutil
from lxml import etree

# Word namespace
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSMAP = {'w': WORD_NS}

def qn(tag):
    """Create qualified name with Word namespace"""
    return f'{{{WORD_NS}}}{tag}'

def get_paragraph_text(p_elem):
    """Extract text from a paragraph element"""
    texts = []
    for t in p_elem.findall('.//w:t', NSMAP):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

def detect_level(text, prev_level=None):
    """
    Detect the appropriate numbering level based on text patterns.
    Returns: level (0-3) or None
    
    Level 0: Clause headings (ALL CAPS, e.g., "DEFINITIONS", "AGENCY")
    Level 1: Sub-clauses (chapeau text introducing lists, or standalone sub-clauses)
    Level 2: (a), (b), (c) items - definition terms or list items
    Level 3: (i), (ii), (iii) items - sub-list items
    """
    text = text.strip()
    
    if not text:
        return None
    
    # Check for clause headings (ALL CAPS, short text) - Level 0
    # These are things like "DEFINITIONS AND INTERPRETATION", "AGENCY", "GENERAL"
    if text.isupper() and len(text) < 60 and not text.startswith('('):
        return 0
    
    # Check if text starts with a quoted term like "Agreement" - these are definition items (Level 2)
    if re.match(r'^"[A-Z][a-zA-Z\s]+"', text):
        return 2
    
    # Check for (i), (ii), (iii), (iv), etc. - Level 3
    if re.match(r'^\([ivx]+\)\s', text, re.IGNORECASE):
        return 3
    
    # Check for (a), (b), (c), etc. - Level 2
    if re.match(r'^\([a-z]\)\s', text, re.IGNORECASE):
        return 2
    
    # Sub-clause patterns - Level 1
    # Typically longer text that introduces a concept or list
    return 1

def has_numbering(p_elem):
    """Check if paragraph has numbering properties"""
    numPr = p_elem.find('.//w:numPr', NSMAP)
    return numPr is not None

def set_ilvl(p_elem, level):
    """Set the ilvl (indent level) for a paragraph's numbering"""
    numPr = p_elem.find('.//w:numPr', NSMAP)
    if numPr is None:
        return False
    
    ilvl = numPr.find('w:ilvl', NSMAP)
    if ilvl is None:
        ilvl = etree.SubElement(numPr, qn('ilvl'))
    
    ilvl.set(qn('val'), str(level))
    return True

def fix_document_numbering(docx_path):
    """Fix the numbering levels in a Word document based on text patterns"""
    
    # Create backup
    backup_path = docx_path.replace('.docx', '_backup.docx')
    shutil.copy2(docx_path, backup_path)
    print(f"Backup created: {backup_path}")
    
    # Work with temp file
    temp_path = docx_path.replace('.docx', '_temp.docx')
    
    # Open the docx as a zip
    with zipfile.ZipFile(docx_path, 'r') as zip_in:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for item in zip_in.infolist():
                data = zip_in.read(item.filename)
                
                if item.filename == 'word/document.xml':
                    # Parse and fix the document
                    root = etree.fromstring(data)
                    
                    # Find all paragraphs
                    paragraphs = root.findall('.//w:p', NSMAP)
                    
                    changes = []
                    for p in paragraphs:
                        if has_numbering(p):
                            text = get_paragraph_text(p)
                            level = detect_level(text)
                            
                            if level is not None:
                                # Get current level
                                numPr = p.find('.//w:numPr', NSMAP)
                                ilvl_elem = numPr.find('w:ilvl', NSMAP) if numPr is not None else None
                                current_level = ilvl_elem.get(qn('val')) if ilvl_elem is not None else '?'
                                
                                if set_ilvl(p, level):
                                    preview = text[:60] + '...' if len(text) > 60 else text
                                    changes.append(f"Level {current_level} -> {level}: {preview}")
                    
                    # Write modified XML
                    data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                    
                    print(f"\nChanges made ({len(changes)} paragraphs):")
                    for c in changes[:20]:  # Show first 20
                        print(f"  {c}")
                    if len(changes) > 20:
                        print(f"  ... and {len(changes) - 20} more")
                
                zip_out.writestr(item, data)
    
    # Replace original with fixed version
    os.remove(docx_path)
    os.rename(temp_path, docx_path)
    print(f"\nDocument updated: {docx_path}")

if __name__ == '__main__':
    docx_path = r'C:\Users\DavidSant\OneDrive - Harper James Solicitors\Documents\Client Matters\Percayso\PDX DSA - Agency Model.docx'
    fix_document_numbering(docx_path)
    print("\nDone!")
