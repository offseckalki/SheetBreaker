# === README.md ===
# SheetBreaker

**Version:** 1.2

SheetBreaker is a Python tool (CLI + GUI) to remove sheet-level protection from `.xlsx` files without knowing the password. It works by safely extracting the archive, cleaning protection tags from individual XML sheets, and repacking the result.

---

### 🚀 Features
- **Removes** `<sheetProtection>` tags without corrupting file structure.
- **Preserves** all data, formatting, and formulas.
- **Dual Mode**: GUI and CLI support.
- **Progress Feedback** in GUI.
- **Direct Link** to [offseckalki.github.io](https://offseckalki.github.io).

---

### 🔧 Usage (CLI)
```bash
python sheetbreaker_cli.py file.xlsx
```
Output will be saved as `file_unprotected.xlsx`

### 🖥️ Usage (GUI)
```bash
python sheetbreaker_gui.py
```
Select file using GUI and done!

---
