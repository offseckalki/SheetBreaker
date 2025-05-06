import os
import zipfile
import shutil
import re
from tkinter import filedialog, messagebox, Tk, Label, Button, ttk, StringVar, Menu
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

def unprotect_xlsx_gui(file_path, status_var, progress, root):
    if not file_path.endswith('.xlsx'):
        messagebox.showerror("Invalid File", "Please select a .xlsx file only.")
        return

    base = Path(file_path).stem
    temp_dir = f"{base}_extracted"
    output_file = f"{base}_unprotected.xlsx"

    status_var.set("Extracting...")
    root.update()

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    sheets_dir = Path(temp_dir) / "xl" / "worksheets"
    sheet_files = list(sheets_dir.glob("*.xml"))
    progress["maximum"] = len(sheet_files)
    progress["value"] = 0

    for i, sheet_file in enumerate(sheet_files):
        remove_protection(sheet_file)
        progress["value"] = i + 1
        root.update()

    calc_chain = Path(temp_dir) / "xl" / "calcChain.xml"
    if calc_chain.exists():
        calc_chain.unlink()

    status_var.set("Repacking...")
    root.update()
    zip_dir_correctly(output_file, temp_dir)
    shutil.rmtree(temp_dir)
    progress["value"] = 0
    status_var.set("Done!")
    messagebox.showinfo("Success", f"Unprotected file saved as:\n{output_file}")

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        unprotect_xlsx_gui(file_path, status_var, progress, root)

def open_link():
    import webbrowser
    webbrowser.open("https://offseckalki.github.io")

root = Tk()
root.title("SheetBreaker GUI")
root.geometry("500x260")
root.resizable(False, False)

menubar = Menu(root)
offsec_menu = Menu(menubar, tearoff=0)
offsec_menu.add_command(label="offseckalki.github.io", command=open_link)
menubar.add_cascade(label="offseckalki", menu=offsec_menu)
root.config(menu=menubar)

status_var = StringVar()
status_var.set("Ready")

progress = ttk.Progressbar(root, length=400, mode='determinate')

Label(root, text="SheetBreaker", font=("Consolas", 20, "bold")).pack(pady=10)
Label(root, text="Remove protection from XLSX files", font=("Consolas", 10)).pack(pady=2)
Button(root, text="Select XLSX File", command=browse_file, width=30, height=2, bg="#222", fg="lime", font=("Consolas", 11)).pack(pady=10)
progress.pack(pady=5)
Label(root, textvariable=status_var, font=("Consolas", 10)).pack(pady=5)
Label(root, text="Made with love by @offseckalki", font=("Consolas", 9)).pack(pady=3)

root.mainloop()