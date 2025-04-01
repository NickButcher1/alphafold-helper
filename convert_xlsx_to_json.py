#
# This script converts an XLSX file into a JSON file, for upload to Alphafold Server.
#
# The XLSX file format is two columns: protein name and sequence.
#
# Run with `python convert_xlsx_to_json input_filename target_protein`, where target_protein can
# be in any row in the XLSX file.
#
# Prerequisites: `pip install pandas openpyxl`.

import pandas as pd
import sys
import os

args = sys.argv
if len(args) < 3:
    print("Invalid parameters")
    print("Usage:   python convert_xlsx_to_json input_filename target_protein")
    print('Example: python convert_xlsx_to_json "Lectin sequences2.xlsx" mKLRI1')
    sys.exit(0)
input_filename = args[1]
target_protein = args[2]

# Extract the base name of the input file without the extension.
input_base_name = os.path.splitext(os.path.basename(input_filename))[0]

# Define the JSON filename to include both the input file name and the target protein.
JSON_FILENAME = f"{input_base_name}_{target_protein}.json"

# Read the XLSX file into memory. Set header to None because default assumes first row of XLSX is
# header, not data.
df = pd.read_excel(input_filename, header=None)

# Find the sequence for the target protein.
target_sequence = None
for index, row in df.iterrows():
    if row[df.columns[0]] == target_protein:
        target_sequence = row[df.columns[1]]
if target_sequence is None:
    print(f"Error: target protein '{target_protein}' not found in xlsx file.")
    sys.exit(0)

# Open JSON output file and write the header.
json_file = open(JSON_FILENAME, "w")
json_file.write("[\n")

# Loop over input proteins and, if not the target protein, write the JSON.
for index, row in df.iterrows():
    protein_name = row[df.columns[0]]
    protein_sequence = row[df.columns[1]]
    if protein_name != target_protein:
        json_file.write("  {\n")
        json_file.write('    "name": "' + target_protein + "_" + protein_name + '",\n')
        json_file.write('    "modelSeeds": [],\n')
        json_file.write('    "sequences": [\n')
        json_file.write(
            '      {\n        "proteinChain": {\n          "sequence": "'
            + target_sequence
            + '",\n          "count": 1,\n'
        )
        json_file.write('          "useStructureTemplate": true,\n')
        json_file.write(
            '          "maxTemplateDate": "2025-02-03"\n        }\n      },\n'
        )

        json_file.write(
            '      {\n        "proteinChain": {\n          "sequence": "'
            + protein_sequence
            + '",\n          "count": 1,\n'
        )
        json_file.write('          "useStructureTemplate": true,\n')
        json_file.write(
            '          "maxTemplateDate": "2025-02-03"\n        }\n      }\n'
        )

        json_file.write("    ],\n")
        json_file.write('    "dialect": "alphafoldserver",\n')
        json_file.write('    "version": 1\n')
        if index == (len(df) - 1):
            json_file.write("  }\n")
        else:
            json_file.write("  },\n")

# Write the JSON footer.
json_file.write("]\n")
