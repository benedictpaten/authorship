import csv
import argparse
import sys
import zipfile
import os
from xml.sax.saxutils import escape

# ==========================================
# CONSTANTS & TEMPLATES
# ==========================================

# 1. [Content_Types].xml - Defines the file types inside the zip
CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

# 2. _rels/.rels - Defines the relationship to the main document part
RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

# 3. word/document.xml (Header/Footer wrapper)
# We will inject the body content into this template.
DOCUMENT_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        {body_content}
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
    </w:body>
</w:document>"""

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def create_run_xml(text, superscription=False, bold=False):
    """Creates a Word XML <w:r> (run) element containing text."""
    props = ""
    if superscription or bold:
        props += "<w:rPr>"
        if bold:
            props += "<w:b/>"
        if superscription:
            props += "<w:vertAlign w:val=\"superscript\"/>"
        props += "</w:rPr>"
    
    # Escape special characters for XML validity (&, <, >)
    safe_text = escape(text)
    
    return f"<w:r>{props}<w:t xml:space=\"preserve\">{safe_text}</w:t></w:r>"

def create_paragraph_xml(runs):
    """Creates a Word XML <w:p> (paragraph) element containing a list of run XML strings."""
    return f"<w:p>{''.join(runs)}</w:p>"

def read_csv_data(names_file, affiliations_file):
    """Reads the CSV inputs and returns structured lists/dicts."""
    
    # Read Affiliations
    affiliations = {} # {id: name}
    try:
        with open(affiliations_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            next(reader, None) # Skip header
            for row in reader:
                if len(row) >= 2:
                    aff_id, aff_name = row[0], row[1]
                    affiliations[aff_id] = aff_name
    except FileNotFoundError:
        print(f"Error: Affiliations file '{affiliations_file}' not found.")
        sys.exit(1)

    # Read Names
    authors = [] # list of dicts
    try:
        with open(names_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            next(reader, None) # Skip header
            for row in reader:
                if len(row) >= 3:
                    authors.append({
                        'last': row[0],
                        'first': row[1],
                        'affils': row[2]
                    })
    except FileNotFoundError:
        print(f"Error: Names file '{names_file}' not found.")
        sys.exit(1)
        
    return authors, affiliations

def generate_document_xml(authors, affiliations):
    """Generates the content for word/document.xml."""
    
    body_parts = []

    # 1. Main Title
    title_run = create_run_xml("Authors", bold=True)
    body_parts.append(create_paragraph_xml([title_run]))
    
    # 2. Authors Block (Single paragraph, comma separated)
    author_runs = []
    total_authors = len(authors)
    
    for i, auth in enumerate(authors):
        # Name part (e.g., "Jane Doe")
        full_name = f"{auth['first']} {auth['last']}".strip()
        author_runs.append(create_run_xml(full_name))
        
        # Affiliation part (Superscript, e.g., "1,2")
        if auth['affils']:
            author_runs.append(create_run_xml(auth['affils'], superscription=True))
        
        # Comma separator (except for last author)
        if i < total_authors - 1:
            author_runs.append(create_run_xml(", "))
            
    body_parts.append(create_paragraph_xml(author_runs))
    
    # 3. "Affiliations" Header
    # Add an empty line before text
    body_parts.append(create_paragraph_xml([])) 
    aff_header_run = create_run_xml("Affiliations", bold=True)
    body_parts.append(create_paragraph_xml([aff_header_run]))
    
    # 4. Affiliations List (One paragraph per affiliation)
    # Sort by ID (converting to int for correct numerical sorting)
    sorted_ids = sorted(affiliations.keys(), key=lambda x: int(x) if x.isdigit() else x)
    
    for aff_id in sorted_ids:
        # Bold the ID number? The original text looked like "1Arizona..."
        # Usually it's nice to have "1 Arizona..." 
        # Based on previous input parsing, the ID was separate.
        # We will render it as "ID AffiliationText"
        
        # Run 1: The number (superscript or plain? Standard lists usually plain but let's stick to the input style)
        # Input style was "1Arizona State..." (often Superscript in Word doc, but plain text in extraction).
        # We will make the number Superscript to match the author markers visually.
        id_run = create_run_xml(aff_id, superscription=True)
        
        # Run 2: The text
        text_run = create_run_xml(f" {affiliations[aff_id]}")
        
        body_parts.append(create_paragraph_xml([id_run, text_run]))

    return DOCUMENT_XML_TEMPLATE.format(body_content="".join(body_parts))

def write_docx(output_filename, document_xml_content):
    """Writes the valid .docx zip structure."""
    try:
        with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as zf:
            # Add static required files
            zf.writestr('[Content_Types].xml', CONTENT_TYPES_XML)
            zf.writestr('_rels/.rels', RELS_XML)
            
            # Add the generated document content
            zf.writestr('word/document.xml', document_xml_content)
    except Exception as e:
        print(f"Error writing .docx file: {e}")
        sys.exit(1)

# ==========================================
# MAIN EXECUTION
# ==========================================

def main():
    parser = argparse.ArgumentParser(
        description="Convert names.csv and affiliations.csv back into a valid .docx file.",
        epilog="Example: python csv_to_docx.py names.csv affiliations.csv output.docx"
    )
    
    parser.add_argument("names_file", help="Path to the names CSV file (Columns: Last, First, Affil IDs)")
    parser.add_argument("affiliations_file", help="Path to the affiliations CSV file (Columns: ID, Name)")
    parser.add_argument("output_file", help="Path for the output .docx file")
    
    args = parser.parse_args()
    
    print("Reading CSV data...")
    authors, affiliations = read_csv_data(args.names_file, args.affiliations_file)
    print(f"Loaded {len(authors)} authors and {len(affiliations)} affiliations.")
    
    print("Generating XML content...")
    doc_xml = generate_document_xml(authors, affiliations)
    
    print(f"Writing to {args.output_file}...")
    write_docx(args.output_file, doc_xml)
    
    print("Done.")

if __name__ == "__main__":
    main()
