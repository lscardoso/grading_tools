#!python
# Auto prepare exam files for grading

import string
import argparse
import os  # Added to handle directory creation
from os import listdir, rename
import shutil
from os.path import isfile, join
from PyPDF2 import PdfReader, PdfWriter
import unicodedata

def strip_accents(s):
   return ''.join(c for c in unicodedata.normalize('NFD', s)
                  if unicodedata.category(c) != 'Mn')

def split_exam(input_file, output_file):
  """
  Splits an A3 landscape PDF file into A4 portrait pages.

  Args:
      input_file (str): Path to the input PDF file.
      output_file (str): Path to the output PDF file.
  """
  with open(input_file, 'rb') as f:
    reader = PdfReader(f)
    writer = PdfWriter()

    first_page = reader.pages[1]
    # Get page dimensions and check if landscape
    width, height = first_page.mediabox.upper_right
    if width < height:
      # Not landscape, leaving the file alone
      return 0


    # Read odd pages
    for index in range(0,int(len(reader.pages)/2)):
      page_front = reader.pages[2*index]
      page_back = reader.pages[(2*index)+1]

      # Split A3 page vertically into two halves (assuming landscape orientation)
      half_width = width / 2

      # Page 1/4
      page_front.mediabox.lower_left = (half_width, 0)
      page_front.mediabox.upper_right = (width, height)
      writer.add_page(page_front)

      # Page 2/4
      page_back.mediabox.lower_left = (0, 0)
      page_back.mediabox.upper_right = (half_width, height)
      writer.add_page(page_back)

      # Page 3/4
      page_back.mediabox.lower_left = (half_width, 0)
      page_back.mediabox.upper_right = (width, height)
      writer.add_page(page_back)

      # Page 4/4
      page_front.mediabox.lower_left = (0, 0)
      page_front.mediabox.upper_right = (half_width, height)
      writer.add_page(page_front)

    with open(output_file, 'wb') as f:
      writer.write(f)

  return 1

# ---------------------------------------------------------
# Command Line Argument Parsing
# ---------------------------------------------------------
parser = argparse.ArgumentParser(description="Prepare and rename exam files from scans based on a class CSV list.")

# Obligatory argument: CSV Group File
parser.add_argument(
    '--group', '-g', 
    required=True, 
    help='Path to the CSV file containing student names (e.g., ../g62.csv)'
)

# Optional argument: Source Path (Default: scans)
parser.add_argument(
    '--source', '-s', 
    default="scans", 
    help='Folder containing the source PDF scans (default: scans)'
)

# Optional argument: Destination Path (Default: ready_to_grade)
parser.add_argument(
    '--dest', '-d', 
    default="ready_to_grade", 
    help='Folder where prepared PDFs will be saved (default: ready_to_grade)'
)

# Optional argument: Exclude list (Default: empty)
parser.add_argument(
    '--exclude', '-e', 
    nargs='*', 
    default=[], 
    help='List of last names to exclude (e.g., -e dupont richard)'
)

args = parser.parse_args()

# Mapping arguments to the original variable names
source_path = args.source
dest_path = args.dest
grp_file = args.group
exclude = args.exclude

# ---------------------------------------------------------
# Directory Safety Check
# ---------------------------------------------------------
# Create the destination directory if it doesn't exist
if not os.path.exists(dest_path):
    os.makedirs(dest_path, exist_ok=True)
    print(f"Created destination directory: {dest_path}")

# ---------------------------------------------------------
# Original Logic
# ---------------------------------------------------------

# Open the list file
with open(join(grp_file)) as gf:
	content = gf.readlines()[1:] # to skip the CSV header
	content = [x.strip() for x in content]

# Get sorted file listing and avoid all hidden files
files = [f for f in listdir(source_path) if isfile(join(source_path, f)) and f[0] != '.']
files.sort()

file_counter = 0

for i, csv in enumerate(content):

	students_names = csv.split(";")[1].split(" ")

	last_name = "-".join(w.lower() for w in students_names if w.upper() == w)
	first_name = "-".join(w.lower() for w in students_names if w.upper() != w)
  
	last_name = strip_accents(last_name)
	first_name = strip_accents(first_name)

	source_exam_file = join(source_path, files[file_counter])
	dest_exam_file = join(dest_path, last_name+"."+first_name+'.pdf')

	if not any(last_name in name for name in exclude):
		print("Working on: " + last_name + ", " + first_name)
		if not split_exam(source_exam_file, dest_exam_file):
			shutil.copyfile(source_exam_file, dest_exam_file)
		
		file_counter = file_counter+1
	else:
		print("  skipping: " + last_name + ", " + first_name)