import os
import zipfile
import shutil
import re
from pathlib import Path

def remove_protection(sheet_path):
    try:
        with open(sheet_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        new_content = re.sub(r'<sheetProtection[^>]*/>', '', content)

        with open(sheet_path, 'w', encoding='utf-8') as f:
            f.write(new_content)

    except Exception as e:
        print(f"[!] Failed to clean {sheet_path}: {e}")

def zip_dir_correctly(output_file, source_dir):
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root_dir, _, files in os.walk(source_dir):
            for file in files:
                full_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(full_path, source_dir)
                zipf.write(full_path, rel_path)

def unprotect_xlsx(file_path):
    if not file_path.endswith('.xlsx'):
        print("[!] Not an .xlsx file")
        return

    base = Path(file_path).stem
    temp_dir = f"{base}_extracted"
    output_file = f"{base}_unprotected.xlsx"

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    sheets_dir = Path(temp_dir) / "xl" / "worksheets"
    for sheet_file in sheets_dir.glob("*.xml"):
        remove_protection(sheet_file)

    calc_chain = Path(temp_dir) / "xl" / "calcChain.xml"
    if calc_chain.exists():
        calc_chain.unlink()

    zip_dir_correctly(output_file, temp_dir)
    shutil.rmtree(temp_dir)
    print(f"[+] Done: {output_file}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python sheetbreaker_cli.py <file.xlsx>")
    else:
        unprotect_xlsx(sys.argv[1])