import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path
from colorama import Fore, Style, init
from tqdm import tqdm

init(autoreset=True)

ASCII_BANNER = f"""
{Fore.GREEN}
  ███████╗██╗  ██╗███████╗████████╗██╗  ██╗██████╗ ███████╗██████╗ 
  ██╔════╝██║  ██║██╔════╝╚══██╔══╝██║  ██║██╔══██╗██╔════╝██╔══██╗
  ███████╗███████║█████╗     ██║   ███████║██████╔╝█████╗  ██████╔╝
  ╚════██║██╔══██║██╔══╝     ██║   ██╔══██║██╔═══╝ ██╔══╝  ██╔══██╗
  ███████║██║  ██║███████╗   ██║   ██║  ██║██║     ███████╗██║  ██║
  ╚══════╝╚═╝  ╚═╝╚══════╝   ╚═╝   ╚═╝  ╚═╝╚═╝     ╚══════╝╚═╝  ╚═╝
              {Style.RESET_ALL}{Fore.CYAN}Break the protection. Own the sheet.{Style.RESET_ALL}
"""

def remove_protection(sheet_path):
    try:
        tree = ET.parse(sheet_path)
        root = tree.getroot()
        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        for elem in root.findall('main:sheetProtection', ns):
            root.remove(elem)

        tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    except ET.ParseError:
        print(f"{Fore.YELLOW}[!] Could not parse {sheet_path}{Style.RESET_ALL}")

def unprotect_xlsx(file_path):
    if not file_path.endswith('.xlsx'):
        print(f"{Fore.RED}[-] File must be .xlsx only.{Style.RESET_ALL}")
        return

    base = Path(file_path).stem
    temp_dir = f"{base}_extracted"
    output_file = f"{base}_unprotected.xlsx"

    print(f"{Fore.CYAN}[*] Extracting XLSX contents...{Style.RESET_ALL}")
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    sheets_dir = Path(temp_dir) / "xl" / "worksheets"
    sheet_files = list(sheets_dir.glob("*.xml"))

    print(f"{Fore.CYAN}[*] Removing protection from sheets...{Style.RESET_ALL}")
    for sheet_file in tqdm(sheet_files, desc="Processing Sheets"):
        remove_protection(sheet_file)

    print(f"{Fore.CYAN}[*] Repacking XLSX...{Style.RESET_ALL}")
    shutil.make_archive(base, 'zip', temp_dir)
    os.rename(f"{base}.zip", output_file)

    print(f"{Fore.GREEN}[+] Unprotected file created: {output_file}{Style.RESET_ALL}")
    shutil.rmtree(temp_dir)

if __name__ == "__main__":
    print(ASCII_BANNER)
    path = input(f"{Fore.YELLOW}[?] Enter path to .xlsx file: {Style.RESET_ALL}")
    if os.path.isfile(path):
        unprotect_xlsx(path)
    else:
        print(f"{Fore.RED}[-] File not found. Check the path and try again.{Style.RESET_ALL}")
