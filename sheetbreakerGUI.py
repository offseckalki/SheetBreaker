import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
from tkinter import filedialog, messagebox, Tk, Label, Button, ttk, StringVar
from pathlib import Path

# Setup GUI window
root = Tk()
root.title("SheetBreaker - XLSX Unprotector")
root.geometry("500x250")
root.resizable(False, False)

status_var = StringVar()
status_var.set("Ready")

progress = ttk.Progressbar(root, length=400, mode='determinate')

def remove_protection(sheet_path):
    try:
        ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        tree = ET.parse(sheet_path)
        root = tree.getroot()
        ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        for elem in root.findall('main:sheetProtection', ns):
            root.remove(elem)

        tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    except ET.ParseError:
        print(f"[!] Could not parse: {sheet_path}")

def zip_dir_correctly(output_file, source_dir):
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root_dir, _, files in os.walk(source_dir):
            for file in files:
                full_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(full_path, source_dir)
                zipf.write(full_path, rel_path)

def unprotect_xlsx_gui(file_path):
    if not file_path.endswith('.xlsx'):
        messagebox.showerror("Invalid File", "Please select a .xlsx file only.")
        return

    base = Path(file_path).stem
    temp_dir = f"{base}_extracted"
    output_file = f"{base}_unprotected.xlsx"

    status_var.set("Extracting file...")
    root.update()

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    sheets_dir = Path(temp_dir) / "xl" / "worksheets"
    sheet_files = list(sheets_dir.glob("*.xml"))

    progress["maximum"] = len(sheet_files)
    progress["value"] = 0
    status_var.set("Removing protection...")
    root.update()

    for i, sheet_file in enumerate(sheet_files):
        remove_protection(sheet_file)
        progress["value"] = i + 1
        root.update()

    status_var.set("Repacking file...")
    root.update()
    zip_dir_correctly(output_file, temp_dir)

    shutil.rmtree(temp_dir)
    progress["value"] = 0
    status_var.set("Completed!")
    messagebox.showinfo("Done", f"✅ Unprotected file created:\n{output_file}")

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        unprotect_xlsx_gui(file_path)

# GUI Layout
Label(root, text="🛡 SheetBreaker", font=("Consolas", 20, "bold")).pack(pady=10)
Label(root, text="Remove sheet protection from XLSX files easily!", font=("Consolas", 10)).pack(pady=2)
Button(root, text="Select XLSX File", command=browse_file, width=30, height=2, bg="#222", fg="lime", font=("Consolas", 11)).pack(pady=10)
progress.pack(pady=5)
Label(root, textvariable=status_var, font=("Consolas", 10)).pack(pady=5)

root.mainloop()
