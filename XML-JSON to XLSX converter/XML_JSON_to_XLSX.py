import pandas as pd
import json
import xml.etree.ElementTree as ET

def json_to_excel(json_file, excel_file):
    # Load JSON data
    with open(json_file, 'r') as f:
        data = json.load(f)

    # Create DataFrame from JSON data
    df = pd.json_normalize(data)

    # Write DataFrame to Excel file
    df.to_excel(excel_file, index=False)

def xml_to_excel(xml_file, excel_file):
    # Parse XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Extract data from XML and create a list of dictionaries
    data = []
    for child in root:
        item = {}
        for subchild in child:
            item[subchild.tag] = subchild.text
        data.append(item)

    # Create DataFrame from XML data
    df = pd.DataFrame(data)

    # Write DataFrame to Excel file
    df.to_excel(excel_file, index=False)

# Prompt user for input and output file paths
input_file = input("Enter the path to the input file (JSON or XML): ")
output_file = input("Enter the path to the output Excel file: ")

# Determine file types and call the appropriate function
if input_file.lower().endswith('.json'):
    json_to_excel(input_file, output_file)
elif input_file.lower().endswith('.xml'):
    xml_to_excel(input_file, output_file)
else:
    print("Unsupported file type. Please provide a JSON or XML file.")
