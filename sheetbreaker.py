import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path

def remove_sheet_protection(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    for elem in root.findall('main:sheetProtection', ns):
        root.remove(elem)

    tree.write(xml_path, encoding='utf-8', xml_declaration=True)

def unprotect_xlsx(file_path):
    if not file_path.endswith('.xlsx'):
        print("File must be a .xlsx file.")
        return

    base = Path(file_path).stem
    temp_dir = f"{base}_temp"
    output_file = f"{base}_unprotected.xlsx"

    # Step 1: Unzip the .xlsx
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Step 2: Edit all worksheet files
    sheets_dir = Path(temp_dir) / "xl" / "worksheets"
    for sheet_file in sheets_dir.glob("*.xml"):
        remove_sheet_protection(sheet_file)

    # Step 3: Zip again
    shutil.make_archive(base, 'zip', temp_dir)
    os.rename(f"{base}.zip", output_file)

    # Step 4: Clean up
    shutil.rmtree(temp_dir)
    print(f"✅ Unprotected file created: {output_file}")

# Example usage:
# unprotect_xlsx("punchdata.xlsx")