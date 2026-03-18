# Skoda & VW After Sales Invoice Extractor

A robust desktop application built for Nagarkot Forwarders Pvt Ltd to extract structured line-item data from **Skoda AG** and **Volkswagen AG** After Sales invoice PDFs.

## Tech Stack
- **Python 3.10+**
- **Tkinter** (GUI)
- **pdfplumber** (PDF Parsing)
- **Pandas & openpyxl** (Excel Generation)
- **Pillow** (Image handling for branding)

---

## Installation

### 1. Project Files
Ensure you have the source code and the `Nagarkot Logo.png` file in the same directory. Note that the application is built with `sys._MEIPASS` support, meaning once compiled to an EXE, the logo will be bundled inside.

### 2. Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment to ensure dependency isolation.

**Create Virtual Environment:**
```bash
python -m venv venv
```

**Activate Environment:**

- **Windows:**
  ```cmd
  venv\Scripts\activate
  ```
- **Mac/Linux:**
  ```bash
  source venv/bin/activate
  ```

**Install Dependencies:**
```bash
pip install -r requirements.txt
```

---

## Usage

1. **Run Application:**
   ```bash
   python Skoda_AfterSales_Extractor_App.py
   ```
2. **Select Format:** Choose between **Skoda AG** or **Volkswagen AG** using the radio buttons. This ensures the correct parsing logic is applied to the specific invoice layout.
3. **Select PDFs:** Click "Select PDFs" and choose one or more invoice files.
4. **Choose Output Folder:** Select where the Excel files should be saved.
5. **Processing Mode:**
   - **Combined:** Merges all selected invoices into a single Excel file (.xlsx).
   - **Individual:** Generates one Excel file (.xlsx) per invoice.
6. **Extract:** Click "Extract & Generate Excel" to process.

---

## Features
- **Multi-Format Support:** Dedicated logic for both Skoda and VW invoice layouts.
- **Intelligent Coordinate Parsing:** Handles complex landscape layouts and multi-line item data.
- **Auto-Formatting:** Converts European number formats (e.g., `2.236,90`) to standard calculation-ready floats.
- **Data Integrity:** Identifiers (Part No, HS Code) are preserved as text to keep leading zeros; numeric values are exported as true numbers for Excel calculations.
- **Nagarkot Branding:** Professional GUI with company colors and embedded logo support.
- **Batch Processing:** Processes hundreds of pages in seconds.

---

## Build Executable

To generate a standalone `.exe` with the logo bundled inside, use the following command (requires `pyinstaller`):

```bash
pyinstaller --name="Nagarkot_Extractor" --onefile --windowed --add-data "Nagarkot Logo.png;." Skoda_AfterSales_Extractor_App.py
```

---

## Notes
- **VAG Numbers**: The script distinguishes between "Net Weight" and "Net Net Weight" to ensure calculations match the line-item totals in the PDF.
- **Separators**: Automatically handles cases where PDF lines or underscores interfere with text extraction.
