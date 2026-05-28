# WHONET File Converter (DBF → Excel)

## 📌 Overview

This tool converts WHONET data files (e.g., `.BRH`, `.BRT`, `.CMC`, etc.) into Excel (`.xlsx`) format using Python.

It is designed to batch-process files in a selected folder and export them into a structured Excel output for further analysis (e.g., completeness reporting).

---

## ⚙️ Features

* Select folder via GUI
* Supports user-defined file extensions
* Batch conversion to Excel
* Handles multiple files at once
* Skips unsupported or empty files safely

---

## 📂 Supported Input

* `.DBF` ✅ (Fully supported)

> ⚠️ Note: Not all WHONET file types are true DBF files. Some may fail to convert.

---

## 📥 Installation

### 1. Install Python packages

```bash
pip install pandas dbfread openpyxl
```

---

## ▶️ How to Use

### Run the script:

```bash
python convert_script.py
```

### Steps:

1. Select the folder containing WHONET files

2. Enter the file extension(s) to process
   Example:

   ```
   .BRH
   ```

   or

   ```
   .BRH,.BRT,.CMC
   ```

3. The script will:

   * Read matching files
   * Convert them to Excel
   * Save them in a folder named:

     ```
     converted_excel/
     ```

---

## 📁 Output Structure

```text
your_folder/
   W0125PHL.BRH
   W0225PHL.BRH

   converted_excel/
       W0125PHL.xlsx
       W0225PHL.xlsx
```

---

## ⚠️ Limitations

* `.BRH` files work reliably
* Other extensions may:

  * Produce empty files
  * Fail due to non-DBF structure
* WHONET may require export via its software for full compatibility

---

## 🧪 Troubleshooting

### No files converted

* Ensure correct extension input (e.g., `.BRH`)
* Check folder contains matching files

### Empty Excel output

* File is not a true DBF format

### Errors during processing

* Try using `.BRH` only
* Ensure files are not corrupted

---

## 🚀 Optional: Convert to Executable

You can build a standalone `.exe` using PyInstaller:

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole convert_script.py
```

Output:

```text
dist/convert_script.exe
```

---

## 🧠 Notes

* WHONET uses multiple internal formats
* `.BRH` is the most compatible for direct conversion
* For full accuracy, consider exporting data directly from WHONET

---

## 📄 License

For internal or research use.

---

## 👨‍💻 Author

Custom-built for batch WHONET data processing and analysis.
