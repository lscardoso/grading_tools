#!/usr/bin/python3
"""
MOODLE EXAM BATCHER & RENAMER
=============================

DESCRIPTION:
    This script automates the preparation of graded exam files for bulk upload 
    to Moodle. It renames local files to match Moodle's specific naming 
    convention and organizes them into compressed zip batches.

FEATURES:
    * Command-Line Interface: Configure paths, files, and delimiters via flags.
    * Smart CSV Parsing: Auto-detects delimiters (comma vs semicolon) and 
      identifies columns by header names (supports English and French Moodle exports).
    * 100MB Batching: Automatically splits files into numbered folders if the 
      total size exceeds 100MB to prevent upload failures.
    * Automatic Compression: Zips each folder batch once it reaches the size 
      limit or completion.
    * Fuzzy Matching: Matches files to students by the student's last name.

HOW IT WORKS:
    1. Scan: Iterates through the provided source directory to index graded files.
    2. Parse: Reads the student CSV, identifying "Full Name" and "Identifier" columns.
    3. Match: For each student in the CSV, it looks for a file starting with their 
       last name (case-insensitive).
    4. Batch: Copies the file to a batch folder while tracking the total folder size.
    5. Zip: Once a folder nears 100MB or the list ends, the folder is compressed 
       into a .zip archive.
"""

import argparse
import csv
import re
import shutil
from pathlib import Path

# Constants
MAX_BATCH_SIZE_MB = 100
MAX_BATCH_SIZE_BYTES = MAX_BATCH_SIZE_MB * 1024 * 1024

def get_args():
    """Configures and parses command line arguments."""
    parser = argparse.ArgumentParser(
        description="Rename exam files for Moodle and package them into 100MB zip batches."
    )
    parser.add_argument("-s", "--source", default="marked", help="Source folder (default: 'marked')")
    parser.add_argument("-d", "--dest", default="moodle", help="Destination base name (default: 'moodle')")
    parser.add_argument("-f", "--file", required=True, dest="grp_file", help="CSV file with student info")
    parser.add_argument(
        "-t", "--delimiter", 
        default=None, 
        help="CSV delimiter (e.g., ',' or ';'). If omitted, it tries to auto-detect."
    )
    return parser.parse_args()

def search_file(starting_string, file_list):
    """Finds the first file that starts with the last name (case-insensitive)."""
    search_term = starting_string.lower().strip()
    for f in file_list:
        if f.lower().startswith(search_term):
            return f
    return None

def main():
    args = get_args()
    source_dir = Path(args.source)
    dest_base = args.dest
    csv_path = Path(args.grp_file)

    if not source_dir.exists() or not csv_path.exists():
        print(f"Error: Path '{source_dir}' or file '{csv_path}' not found.")
        return

    # Get sorted file listing excluding hidden files
    files = sorted([f.name for f in source_dir.iterdir() if f.is_file() and not f.name.startswith('.')])

    batch_number = 1
    current_batch_size = 0
    file_counter = 0
    
    current_dest_dir = Path(f"{dest_base}_{batch_number}")
    current_dest_dir.mkdir(parents=True, exist_ok=True)

    print(f"--- Starting Processing ---")
    print(f"Initial Batch Folder: {current_dest_dir}")

    try:
        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            # Handle Delimiter logic
            if args.delimiter:
                dialect_char = args.delimiter
            else:
                sample = f.read(2048)
                f.seek(0)
                dialect_char = csv.Sniffer().sniff(sample).delimiter
                print(f"Auto-detected CSV delimiter: '{dialect_char}'")

            reader = csv.DictReader(f, delimiter=dialect_char)
            
            # Find column names dynamically (handles various Moodle export languages)
            headers = reader.fieldnames
            name_col = next((h for h in headers if h.lower() in ["nom complet", "full name", "name"]), None)
            id_col = next((h for h in headers if h.lower() in ["identifiant", "identifier", "id number", "id"]), headers[0])

            if not name_col:
                print(f"Error: No 'Name' column found. Headers are: {headers}")
                return

            for row in reader:
                if not row.get(name_col): continue

                # Clean up student info
                moodle_number = "".join(re.findall(r'\d+', row[id_col]))
                full_name = row[name_col].strip()
                last_name = full_name.split(" ")[0].replace('"', '')

                original_filename = search_file(last_name, files)

                if original_filename:
                    source_file = source_dir / original_filename
                    file_size = source_file.stat().st_size

                    # Check if this file pushes the current batch over 100MB
                    if current_batch_size + file_size > MAX_BATCH_SIZE_BYTES:
                        print(f"Batch {batch_number} full (~{current_batch_size/1e6:.1f}MB). Zipping...")
                        shutil.make_archive(str(current_dest_dir), 'zip', current_dest_dir)
                        
                        # Set up new batch
                        batch_number += 1
                        current_batch_size = 0
                        current_dest_dir = Path(f"{dest_base}_{batch_number}")
                        current_dest_dir.mkdir(parents=True, exist_ok=True)

                    # Create Moodle-compliant filename
                    moodle_filename = f"{full_name}_{moodle_number}_assignsubmission_file_{original_filename}"
                    shutil.copyfile(source_file, current_dest_dir / moodle_filename)
                    
                    current_batch_size += file_size
                    file_counter += 1
                else:
                    print(f"  [MISSING] No local file found for student: {full_name}")

        # Final Zip for the remaining batch
        if any(current_dest_dir.iterdir()):
            print(f"Finalizing last batch ({batch_number})...")
            shutil.make_archive(str(current_dest_dir), 'zip', current_dest_dir)

        print("-" * 35)
        print(f"SUCCESS: {file_counter} files processed.")
        print(f"OUTPUT: {batch_number} zip file(s) created.")

    except Exception as e:
        print(f"Critical error during execution: {e}")

if __name__ == "__main__":
    main()