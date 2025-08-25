# XLS to XLSX Converter

A simple tool to convert old Excel `.xls` files (Excel 97–2003) into modern `.xlsx` files (Excel 2007+). Supports **all sheets** in the workbook. 

**MAJOR CAVEAT!!!** This does not convert the formatting of your Excel document. The new document will be data only. So this is useful for automation and data extraction, but not for a presentable spreadsheet that you would distribute.

---

## 📦 Usage (with EXE)

If you just want to use the tool (no Python required):

1. Download or copy the `excel-convert-xls-to-xlsx.exe` file (found in the `dist` folder after build).
2. Open a command prompt in the folder with your `.xls` file.
3. Run:

```bash
excel-convert-xls-to-xlsx.exe input.xls
```

This will create `input.xlsx` in the same folder.

You can also specify an output file:

```bash
excel-convert-xls-to-xlsx.exe input.xls output.xlsx
```

---

## 🐼 Usage (with Python)

If you have Python installed, you can run the script directly.

1. Install dependencies:

   ```bash
   pip install pandas xlrd openpyxl
   ```

2. Run:

   ```bash
   python excel-convert-xls-to-xlsx.py input.xls
   ```

3. Or with custom output:

   ```bash
   python excel-convert-xls-to-xlsx.py input.xls output.xlsx
   ```

---

## 🛠️ Build the EXE (for developers)

If you want to create the `.exe` yourself:

1. Install PyInstaller:

   ```bash
   pip install pyinstaller
   ```

2. Build the executable:

   ```bash
   pyinstaller --onefile excel-convert-xls-to-xlsx.py
   ```

3. After the build finishes, the standalone exe will be in:

   ```
   dist/excel-convert-xls-to-xlsx.exe
   ```

4. You can now distribute `excel-convert-xls-to-xlsx.exe` to others — no Python required on their machines.

---

## ⚠️ Notes

- Formatting (fonts, colors, styles, column widths, etc.) is **not preserved**.
- Only raw data is copied into the `.xlsx` file.
- If you need **full formatting preservation**, consider a commercial library like **Aspose.Cells**.

