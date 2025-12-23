import csv
import argparse
import sys
import os

# ==========================================
# PROCESSING LOGIC
# ==========================================

def renumber_affiliations(names_data, affiliations_map):
    """
    Renumbers affiliations based on the order of appearance in names_data.
    Returns:
        new_affiliations_list: List of (new_id, text)
        updated_names_data: List of author dicts with updated IDs
    """
    
    old_to_new_id_map = {}
    new_affiliations_list = []
    current_new_id = 1
    
    updated_names_data = []

    # Iterate through authors in the order they appear
    for author in names_data:
        original_ids_str = author['affils']
        
        # Skip if author has no affiliations
        if not original_ids_str.strip():
            updated_names_data.append(author)
            continue
            
        # Parse the comma-separated IDs
        # We filter out empty strings to handle cases like "1, 2,"
        original_ids = [x.strip() for x in original_ids_str.split(',') if x.strip()]
        
        new_ids_for_this_author = []
        
        for oid in original_ids:
            # If this affiliation ID hasn't been seen yet, assign a new number
            if oid not in old_to_new_id_map:
                # Check if the old ID actually exists in the affiliations file
                if oid not in affiliations_map:
                    print(f"Warning: Author {author['last']} references affiliation ID '{oid}', which is not in the affiliations file. Keeping original ID.")
                    # Fallback: map to itself
                    old_to_new_id_map[oid] = oid 
                else:
                    # Assign new sequential ID
                    new_id = str(current_new_id)
                    old_to_new_id_map[oid] = new_id
                    
                    # Store the affiliation text with its new ID
                    new_affiliations_list.append((new_id, affiliations_map[oid]))
                    current_new_id += 1
            
            new_ids_for_this_author.append(old_to_new_id_map[oid])
            
        # Update the author dictionary
        updated_author = author.copy()
        updated_author['affils'] = ",".join(new_ids_for_this_author)
        updated_names_data.append(updated_author)

    # Check for orphaned affiliations (affiliations in the file but never used by an author)
    # Append them to the end of the list with new IDs
    for oid, text in affiliations_map.items():
        if oid not in old_to_new_id_map:
            new_id = str(current_new_id)
            old_to_new_id_map[oid] = new_id
            new_affiliations_list.append((new_id, text))
            current_new_id += 1

    return new_affiliations_list, updated_names_data

# ==========================================
# FILE I/O
# ==========================================

def read_inputs(names_file, affiliations_file):
    # Read Affiliations into a Dictionary
    aff_map = {}
    try:
        with open(affiliations_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader, None) # Skip header
            for row in reader:
                if len(row) >= 2:
                    # row[0] = ID, row[1] = Name
                    aff_map[row[0].strip()] = row[1].strip()
    except Exception as e:
        print(f"Error reading affiliations file: {e}")
        sys.exit(1)

    # Read Names into a List
    names_list = []
    try:
        with open(names_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader, None) # Skip header
            for row in reader:
                if len(row) >= 3:
                    names_list.append({
                        'last': row[0].strip(),
                        'first': row[1].strip(),
                        'affils': row[2].strip()
                    })
    except Exception as e:
        print(f"Error reading names file: {e}")
        sys.exit(1)
        
    return names_list, aff_map

def write_csvs(new_affils, updated_names, output_aff_name, output_names_name):
    # Write new Affiliations
    try:
        with open(output_aff_name, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Affiliation ID', 'Affiliation Name'])
            for aff in new_affils:
                writer.writerow(aff)
        print(f"Successfully wrote: {output_aff_name}")
    except Exception as e:
        print(f"Error writing new affiliations file: {e}")

    # Write updated Names
    try:
        with open(output_names_name, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Last Name', 'First Name / Middle', 'Affiliation IDs'])
            for auth in updated_names:
                writer.writerow([auth['last'], auth['first'], auth['affils']])
        print(f"Successfully wrote: {output_names_name}")
    except Exception as e:
        print(f"Error writing updated names file: {e}")

def get_output_filename(input_path):
    """Generates an output filename by prepending 'reordered_' to the base filename."""
    directory, filename = os.path.split(input_path)
    new_filename = "reordered_" + filename
    return os.path.join(directory, new_filename)

# ==========================================
# MAIN
# ==========================================

def main():
    parser = argparse.ArgumentParser(
        description="Reorder affiliations based on appearance in names CSV. Outputs files prefixed with 'reordered_'.",
        epilog="Example: python renumber_affiliations.py names.csv affiliations.csv -> outputs reordered_names.csv and reordered_affiliations.csv"
    )
    
    parser.add_argument("names_input", help="Path to input names CSV")
    parser.add_argument("affiliations_input", help="Path to input affiliations CSV")
    
    args = parser.parse_args()
    
    # Generate output filenames dynamically
    output_names = get_output_filename(args.names_input)
    output_affiliations = get_output_filename(args.affiliations_input)
    
    # 1. Read Data
    print(f"Reading input files: {args.names_input}, {args.affiliations_input}...")
    names_data, aff_map = read_inputs(args.names_input, args.affiliations_input)
    
    # 2. Process
    print("Renumbering...")
    new_affils_list, updated_names = renumber_affiliations(names_data, aff_map)
    
    # 3. Write Data
    write_csvs(new_affils_list, updated_names, output_affiliations, output_names)
    
    print("Done.")

if __name__ == "__main__":
    main()
