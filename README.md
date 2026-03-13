# Skoda After Sales Invoice Extractor

A robust desktop application built for Nagarkot Forwarders Pvt Ltd to extract structured line-item data from Skoda AG After Sales invoice PDFs.

## Tech Stack
- **Python 3.10+**
- **Tkinter** (GUI)
- **pdfplumber** (PDF Parsing)
- **Pillow** (Image handling for branding)

---

## Installation

### 1. Clone or Download
Ensure you have the source code and the `Nagarkot Logo.png` file in the same directory.

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
2. **Select PDFs:** Click "Select PDFs" and choose one or more Skoda After Sales invoice files.
3. **Choose Output Folder:** Select where the CSV files should be saved.
4. **Processing Mode:**
   - **Combined:** Merges all selected invoices into a single CSV.
   - **Individual:** Generates one CSV file per invoice.
5. **Extract:** Click "Extract & Generate CSV" to process the files.

---

## Features
- **Intelligent Parsing:** Handles the complex landscape layout with Order No. / Wrap No. line pairs.
- **Auto Formatting:** Converts European number formats (2.236,90) to standard formats (2,236.90).
- **Nagarkot Branding:** Professional GUI with company colors and logo.
- **Batch Processing:** Handles hundreds of invoices in seconds.

---

## Build Executable (Optional)

If you need to build a standalone `.exe`:

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Build using a Spec file (coming soon) or direct command:
   ```bash
   pyinstaller --name="SkodaAfterSalesExtractor" --onefile --windowed --add-data "Nagarkot Logo.png;." Skoda_AfterSales_Extractor_App.py
   ```
