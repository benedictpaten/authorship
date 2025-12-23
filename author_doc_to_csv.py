import re
import csv
import zipfile
import xml.etree.ElementTree as ET
import argparse
import sys
import os


# ==========================================
# DOCX TEXT EXTRACTOR
# ==========================================
def get_text_from_docx(filename):
    """
    Extracts text from a .docx file (zipped XML) without external dependencies.
    """
    if not os.path.exists(filename):
        print(f"Error: The file '{filename}' was not found.")
        sys.exit(1)

    try:
        with zipfile.ZipFile(filename) as docx:
            # content is in word/document.xml
            xml_content = docx.read('word/document.xml')

            # Parse XML
            tree = ET.fromstring(xml_content)

            full_text = []

            # Find all paragraphs (tag ending in 'p')
            for p in tree.iter():
                if p.tag.endswith('}p'):
                    para_text = []
                    # Find all text nodes (tag ending in 't')
                    for t in p.iter():
                        if t.tag.endswith('}t') and t.text:
                            para_text.append(t.text)

                    if para_text:
                        full_text.append("".join(para_text))

            return "\n".join(full_text)

    except Exception as e:
        print(f"Error reading .docx file: {e}")
        sys.exit(1)


# ==========================================
# PARSING LOGIC
# ==========================================

def parse_data(full_text):
    """
    Splits the document text into Authors and Affiliations.
    Assumes:
      1. The first non-empty line is a Title (and is skipped).
      2. The keyword 'Affiliations' separates the author list from the addresses.
    """
    lines = full_text.split('\n')

    author_lines = []
    affiliation_lines = []

    found_split = False
    title_skipped = False

    for line in lines:
        clean_line = line.strip()
        if not clean_line:
            continue

        # Case-insensitive check for the section divider
        if clean_line.lower() == "affiliations":
            found_split = True
            continue

        if not found_split:
            # We are in the Author section.

            # LOGIC CHANGE: Treat the first non-empty line found as the Title and skip it.
            if not title_skipped:
                title_skipped = True
                continue

            author_lines.append(clean_line)
        else:
            # We are in the Affiliation section.
            affiliation_lines.append(clean_line)

    return "\n".join(author_lines), "\n".join(affiliation_lines)


def parse_affiliations_to_dict(text):
    """Parses affiliation text into a dictionary of {id: affiliation_string}."""
    affiliations = {}
    lines = text.strip().split('\n')

    for line in lines:
        line = line.strip()
        # Regex: Start of line -> (Digits) -> whitespace -> (Rest of text)
        match = re.match(r'^(\d+)\s*(.*)', line)
        if match:
            aff_id = match.group(1)
            aff_text = match.group(2)
            affiliations[aff_id] = aff_text

    return affiliations


def parse_authors_to_list(text):
    """Parses author text into a list of dictionaries."""
    # Join lines to treat as a single stream
    clean_text = text.replace('\n', ' ').strip()

    # Regex: Match non-digits (name) followed by digits (affiliation IDs)
    # This handles accents/utf-8 characters correctly.
    pattern = re.compile(r'([^\d]+?)(\d[\d,]*)')

    authors = []

    for match in pattern.finditer(clean_text):
        name_part = match.group(1).strip()
        affils_part = match.group(2).strip()

        # Cleanup leading commas/spaces
        name_part = name_part.lstrip(', ')

        if not name_part:
            continue

        # Split Name into First/Middle and Last
        name_tokens = name_part.split()

        if len(name_tokens) > 1:
            last_name = name_tokens[-1]
            first_middle = " ".join(name_tokens[:-1])
        else:
            last_name = name_part
            first_middle = ""

        # Clean up affiliation string
        affils_part = affils_part.strip(',')

        authors.append({
            'first_middle': first_middle,
            'last_name': last_name,
            'affiliations': affils_part
        })

    return authors


# ==========================================
# MAIN EXECUTION
# ==========================================

def main():
    parser = argparse.ArgumentParser(
        description="Extract authors and affiliations from a DOCX file into structured CSVs.",
        epilog="""
OUTPUT FILES:
  1. names.csv: [Last Name, First Name / Middle, Affiliation IDs]
  2. affiliations.csv: [Affiliation ID, Affiliation Name]

NOTE:
  The script assumes the first non-empty line of the DOCX is a Title and skips it.
""",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("input_file", help="Path to the input .docx file.")

    args = parser.parse_args()

    input_filename = args.input_file
    names_output = "names.csv"
    affiliations_output = "affiliations.csv"

    print(f"Reading file: {input_filename}...")
    full_text = get_text_from_docx(input_filename)

    print("Parsing content (skipping first line as Title)...")
    authors_text, affiliations_text = parse_data(full_text)

    if not affiliations_text:
        print("Warning: No 'Affiliations' section found.")

    # Process Affiliations
    aff_dict = parse_affiliations_to_dict(affiliations_text)

    # Process Authors
    author_list = parse_authors_to_list(authors_text)
    print(f"Found {len(author_list)} authors and {len(aff_dict)} affiliations.")

    # Write Affiliations CSV
    print(f"Writing {affiliations_output}...")
    with open(affiliations_output, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['Affiliation ID', 'Affiliation Name'])
        for aff_id in sorted(aff_dict.keys(), key=lambda x: int(x)):
            writer.writerow([aff_id, aff_dict[aff_id]])

    # Write Names CSV
    print(f"Writing {names_output}...")
    with open(names_output, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['Last Name', 'First Name / Middle', 'Affiliation IDs'])
        for auth in author_list:
            writer.writerow([
                auth['last_name'],
                auth['first_middle'],
                auth['affiliations']
            ])

    print("Conversion complete.")


if __name__ == "__main__":
    main()