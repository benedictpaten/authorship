#!/usr/bin/env python3

import argparse
import subprocess
import sys
import os

def run_step(command, description):
    """Helper to run a shell command and handle errors."""
    print(f"--- {description} ---")
    print(f"Running: {' '.join(command)}")
    try:
        # subprocess.check_call raises CalledProcessError if the script fails
        subprocess.check_call(command)
        print("Success.\n")
    except subprocess.CalledProcessError as e:
        print(f"Error: The step '{description}' failed.")
        sys.exit(1)
    except FileNotFoundError:
        print(f"Error: Could not find the python executable or the script file.")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(
        description="Full Pipeline: Extract authors -> Renumber Affiliations -> Generate New Docx.",
        epilog="""
DEPENDENCIES:
  This script assumes the following files are in the same directory:
  1. author_doc_to_csv.py
  2. renumber_affiliations.py
  3. csv_to_author_doc.py
"""
    )
    
    parser.add_argument("input_docx", help="Path to the original author list (.docx)")
    parser.add_argument("output_docx", help="Path to save the final cleaned (.docx)")
    parser.add_argument("--keep-csv", action="store_true", help="Do not delete the intermediate CSV files after processing.")
    
    args = parser.parse_args()
    
    # Define script names
    # We use sys.executable to ensure we use the same python interpreter running this script
    PYTHON_EXE = sys.executable
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    
    script_extract = os.path.join(SCRIPT_DIR, "author_doc_to_csv.py")
    script_renumber = os.path.join(SCRIPT_DIR, "renumber_affiliations.py")
    script_generate = os.path.join(SCRIPT_DIR, "csv_to_author_doc.py")

    # Intermediate Filenames (Determined by the logic of the previous scripts)
    # docx_to_csv.py outputs these exact names:
    temp_names = "names.csv"
    temp_affils = "affiliations.csv"
    
    # renumber_affiliations.py outputs these names (prefixed "reordered_"):
    reordered_names = "reordered_names.csv"
    reordered_affils = "reordered_affiliations.csv"

    # ==========================================
    # STEP 1: Extract Data
    # ==========================================
    run_step(
        [PYTHON_EXE, script_extract, args.input_docx],
        "Step 1: Extracting data to CSV"
    )

    # ==========================================
    # STEP 2: Renumber Affiliations
    # ==========================================
    run_step(
        [PYTHON_EXE, script_renumber, temp_names, temp_affils],
        "Step 2: Renumbering affiliations"
    )

    # ==========================================
    # STEP 3: Generate New Document
    # ==========================================
    run_step(
        [PYTHON_EXE, script_generate, reordered_names, reordered_affils, args.output_docx],
        "Step 3: Generating final DOCX"
    )

    # ==========================================
    # CLEANUP
    # ==========================================
    if not args.keep_csv:
        print("--- Cleanup ---")
        files_to_remove = [temp_names, temp_affils, reordered_names, reordered_affils]
        for f in files_to_remove:
            if os.path.exists(f):
                os.remove(f)
                print(f"Removed temporary file: {f}")
        print("Cleanup complete.")
    else:
        print("Skipping cleanup (--keep-csv used). Intermediate CSVs preserved.")

    print(f"\nPipeline Finished Successfully. Output saved to: {args.output_docx}")

if __name__ == "__main__":
    main()
