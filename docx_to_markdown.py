"""
DOCX to Markdown Converter - Compare 3 Methods
Tests Pandoc, Mammoth, and python-docx for numbered list preservation
"""

import subprocess
import sys
from pathlib import Path

# Check and install required packages
def install_if_missing(import_name, package_name=None):
    """Install package if not available. package_name can differ from import_name."""
    if package_name is None:
        package_name = import_name
    try:
        __import__(import_name)
    except ImportError:
        print(f"Installing {package_name}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name, "-q"])

install_if_missing("mammoth")
install_if_missing("docx", "python-docx")  # Import is 'docx', package is 'python-docx'

from docx import Document
from docx.oxml.ns import qn
import mammoth


def convert_with_pandoc(docx_path: str, output_path: str) -> bool:
    """
    Method 1: Pandoc conversion
    Requires pandoc to be installed: winget install pandoc
    """
    try:
        result = subprocess.run(
            ["pandoc", docx_path, "-o", output_path, "--wrap=none", "-t", "gfm"],
            capture_output=True,
            text=True
        )
        if result.returncode == 0:
            print(f"✓ Pandoc: Saved to {output_path}")
            return True
        else:
            print(f"✗ Pandoc error: {result.stderr}")
            return False
    except FileNotFoundError:
        print("✗ Pandoc not installed. Run: winget install pandoc")
        return False


def convert_with_mammoth(docx_path: str, output_path: str) -> bool:
    """
    Method 2: Mammoth conversion
    Semantic conversion that respects Word styles
    """
    try:
        with open(docx_path, "rb") as docx_file:
            result = mammoth.convert_to_markdown(docx_file)
            
        with open(output_path, "w", encoding="utf-8") as md_file:
            md_file.write(result.value)
        
        if result.messages:
            print(f"  Mammoth warnings: {len(result.messages)}")
            for msg in result.messages[:3]:  # Show first 3 warnings
                print(f"    - {msg}")
        
        print(f"✓ Mammoth: Saved to {output_path}")
        return True
    except Exception as e:
        print(f"✗ Mammoth error: {e}")
        return False


def convert_with_python_docx(docx_path: str, output_path: str) -> bool:
    """
    Method 3: python-docx with custom numbered list handling
    Most control over list ordinals and structure
    """
    try:
        doc = Document(docx_path)
        md_lines = []
        list_counters = {}  # Track numbering by (numId, ilvl)
        prev_was_list = False
        
        for para in doc.paragraphs:
            text = para.text.strip()
            
            # Check if paragraph is a numbered/bulleted list item
            numPr = None
            if para._element.pPr is not None:
                numPr_elem = para._element.pPr.find(qn('w:numPr'))
                if numPr_elem is not None:
                    numPr = numPr_elem
            
            if numPr is not None:
                # Get list level and ID
                ilvl_elem = numPr.find(qn('w:ilvl'))
                numId_elem = numPr.find(qn('w:numId'))
                
                ilvl = int(ilvl_elem.get(qn('w:val'))) if ilvl_elem is not None else 0
                numId = int(numId_elem.get(qn('w:val'))) if numId_elem is not None else 0
                
                # Determine if bullet or numbered list
                is_bullet = is_bullet_list(doc, numId, ilvl)
                
                indent = "   " * ilvl
                
                if is_bullet:
                    md_lines.append(f"{indent}- {text}")
                else:
                    # Track counter for this specific list and level
                    key = (numId, ilvl)
                    
                    # Reset deeper levels when going up
                    keys_to_reset = [k for k in list_counters if k[0] == numId and k[1] > ilvl]
                    for k in keys_to_reset:
                        del list_counters[k]
                    
                    list_counters[key] = list_counters.get(key, 0) + 1
                    md_lines.append(f"{indent}{list_counters[key]}. {text}")
                
                prev_was_list = True
                continue
            
            # Handle headings
            if para.style.name.startswith('Heading'):
                level = para.style.name.replace('Heading ', '')
                try:
                    level = int(level)
                    md_lines.append(f"\n{'#' * level} {text}\n")
                    prev_was_list = False
                    list_counters.clear()  # Reset list counters after heading
                    continue
                except ValueError:
                    pass
            
            # Regular paragraph
            if text:
                if prev_was_list:
                    md_lines.append("")  # Add blank line after list
                md_lines.append(text)
                prev_was_list = False
            elif not prev_was_list:
                md_lines.append("")  # Preserve paragraph breaks
        
        # Write output
        with open(output_path, "w", encoding="utf-8") as md_file:
            md_file.write("\n".join(md_lines))
        
        print(f"✓ python-docx: Saved to {output_path}")
        return True
    except Exception as e:
        print(f"✗ python-docx error: {e}")
        import traceback
        traceback.print_exc()
        return False


def is_bullet_list(doc, numId: int, ilvl: int) -> bool:
    """
    Check if a list is a bullet list by examining the numbering definitions
    """
    try:
        # Access numbering part
        numbering_part = doc.part.numbering_part
        if numbering_part is None:
            return False
        
        numbering_xml = numbering_part._element
        
        # Find the abstractNumId for this numId
        for num in numbering_xml.findall(qn('w:num')):
            if num.get(qn('w:numId')) == str(numId):
                abstract_ref = num.find(qn('w:abstractNumId'))
                if abstract_ref is not None:
                    abstract_id = abstract_ref.get(qn('w:val'))
                    
                    # Find the abstract numbering definition
                    for abstract in numbering_xml.findall(qn('w:abstractNum')):
                        if abstract.get(qn('w:abstractNumId')) == abstract_id:
                            # Find the level definition
                            for lvl in abstract.findall(qn('w:lvl')):
                                if lvl.get(qn('w:ilvl')) == str(ilvl):
                                    numFmt = lvl.find(qn('w:numFmt'))
                                    if numFmt is not None:
                                        fmt = numFmt.get(qn('w:val'))
                                        return fmt == 'bullet'
        return False
    except Exception:
        return False


def compare_outputs(base_path: str):
    """Show a brief comparison of the three outputs"""
    methods = ['pandoc', 'mammoth', 'python_docx']
    
    print("\n" + "="*70)
    print("OUTPUT COMPARISON - First 50 lines of each")
    print("="*70)
    
    for method in methods:
        output_path = f"{base_path}_{method}.md"
        if Path(output_path).exists():
            print(f"\n{'='*30} {method.upper()} {'='*30}")
            with open(output_path, "r", encoding="utf-8") as f:
                lines = f.readlines()[:50]
                for i, line in enumerate(lines, 1):
                    print(f"{i:3}: {line.rstrip()}")
            print(f"... ({len(open(output_path, encoding='utf-8').readlines())} total lines)")


def main():
    # Get input file from command line or prompt
    if len(sys.argv) > 1:
        docx_path = sys.argv[1]
    else:
        docx_path = input("Enter path to DOCX file: ").strip().strip('"')
    
    if not Path(docx_path).exists():
        print(f"Error: File not found: {docx_path}")
        return
    
    # Create output paths
    base_name = Path(docx_path).stem
    output_dir = Path(docx_path).parent
    
    print(f"\nConverting: {docx_path}")
    print(f"Output directory: {output_dir}")
    print("="*70)
    
    # Run all three conversions
    pandoc_out = str(output_dir / f"{base_name}_pandoc.md")
    mammoth_out = str(output_dir / f"{base_name}_mammoth.md")
    docx_out = str(output_dir / f"{base_name}_python_docx.md")
    
    print("\n[Method 1] Pandoc (external tool)")
    convert_with_pandoc(docx_path, pandoc_out)
    
    print("\n[Method 2] Mammoth (semantic conversion)")
    convert_with_mammoth(docx_path, mammoth_out)
    
    print("\n[Method 3] python-docx (custom list handling)")
    convert_with_python_docx(docx_path, docx_out)
    
    # Show comparison
    compare_outputs(str(output_dir / base_name))
    
    print("\n" + "="*70)
    print("FILES CREATED:")
    print(f"  1. {pandoc_out}")
    print(f"  2. {mammoth_out}")
    print(f"  3. {docx_out}")
    print("="*70)


if __name__ == "__main__":
    main()
