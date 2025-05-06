# SheetBreaker

**Version:** 1.2

SheetBreaker is a Python tool (CLI + GUI) to remove sheet-level protection from `.xlsx` files without knowing the password. It works by safely extracting the archive, cleaning protection tags from individual XML sheets, and repacking the result.

---

### 🚀 Features
- **Removes** `<sheetProtection>` tags without corrupting file structure.
- **Preserves** all data, formatting, and formulas.
- **Dual Mode**: GUI and CLI support.
- **Progress Feedback** in GUI.
- [offseckalki.github.io](https://offseckalki.github.io).

---

### 🔧 Usage (CLI)
```bash
python sheetbreaker_cli.py file.xlsx
```
Output will be saved as `file_unprotected.xlsx`

### 🖥️ Usage (GUI)

**On Windows**

Download the .exe file from [here](https://objects.githubusercontent.com/github-production-release-asset-2e65be/978749497/7c48acac-09cb-4a78-a2c9-90a274fb617b?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=releaseassetproduction%2F20250506%2Fus-east-1%2Fs3%2Faws4_request&X-Amz-Date=20250506T175807Z&X-Amz-Expires=300&X-Amz-Signature=cea8eebf89212ee6d2cac4362be67c0b274e6fe3206f7a1a8fa4ae6a66e5bc1f&X-Amz-SignedHeaders=host&response-content-disposition=attachment%3B%20filename%3DsheetbreakerGUI.exe&response-content-type=application%2Foctet-stream)

Open it 
Select file and Voila !
The sheets are now unprotected.

```bash
python sheetbreaker_gui.py
```
Select file using GUI and done!

---
