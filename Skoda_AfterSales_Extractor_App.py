"""
Skoda AG After Sales – Invoice Extractor
Extracts structured line-item data from Skoda AG After Sales invoices (PDF).
Handles the landscape layout with Order No. / Wrap No. line pairs,
EUR-style number formatting, and multi-PDF batch processing.
Outputs CSV with INR/USD-style number formatting.
"""

import os
import sys
import re
import csv
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
from typing import Optional

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None


# ---------- RESOURCE PATH FUNCTION ----------
def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except AttributeError:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


# ---------- EUR-STYLE → STANDARD NUMBER FORMATTING ----------
def convert_eur_to_standard_format(value_str: str) -> str:
    """
    Convert European number format to standard (INR/USD) format.
    European: 2.236,90 → Standard: 2,236.90
    European: 0,297   → Standard: 0.297
    European: 43.760,64 → Standard: 43,760.64
    If already standard-style (period as decimal, no comma), pass through.
    """
    if not value_str or not isinstance(value_str, str):
        return value_str

    value_str = value_str.strip()

    # Case 1: Both '.' and ',' present — European format
    if '.' in value_str and ',' in value_str:
        dot_pos = value_str.rfind('.')
        comma_pos = value_str.rfind(',')

        if comma_pos > dot_pos:
            # European: dots=thousands, comma=decimal  (e.g. 2.236,90)
            converted = value_str.replace('.', '').replace(',', '.')
        else:
            # Already standard: commas=thousands, period=decimal (e.g. 2,236.90)
            return value_str

        try:
            num = float(converted)
            if '.' in converted:
                decimal_places = len(converted.split('.')[-1])
                return f"{num:,.{decimal_places}f}"
            return f"{num:,.0f}"
        except ValueError:
            return value_str

    # Case 2: Only comma present — it's the decimal separator
    elif ',' in value_str and '.' not in value_str:
        converted = value_str.replace(',', '.')
        try:
            num = float(converted)
            decimal_places = len(converted.split('.')[-1])
            return f"{num:,.{decimal_places}f}"
        except ValueError:
            return value_str

    # Case 3: Already in standard format or integer
    else:
        return value_str


def eur_str_to_float(value_str: str) -> float:
    """
    Parse a European-formatted number string to a Python float.
    e.g. '2.236,90' → 2236.90, '0,297' → 0.297
    Also handles standard-style strings correctly.
    """
    if not value_str or not isinstance(value_str, str):
        return 0.0
    value_str = value_str.strip()

    if '.' in value_str and ',' in value_str:
        dot_pos = value_str.rfind('.')
        comma_pos = value_str.rfind(',')
        if comma_pos > dot_pos:
            # European: 2.236,90
            return float(value_str.replace('.', '').replace(',', '.'))
        else:
            # Standard: 2,236.90
            return float(value_str.replace(',', ''))
    elif ',' in value_str:
        return float(value_str.replace(',', '.'))
    else:
        try:
            return float(value_str)
        except ValueError:
            return 0.0


def smart_format_number(value_str: str) -> str:
    """
    Intelligently format a number string from EUR-style to standard.
    Always output standard-style (commas=thousands, period=decimal).
    """
    num = eur_str_to_float(value_str)
    # Detect how many decimal places the original had
    cleaned = value_str.strip()
    if ',' in cleaned and '.' in cleaned:
        # EUR style: decimal part is after comma
        comma_pos = cleaned.rfind(',')
        decimal_places = len(cleaned) - comma_pos - 1
    elif ',' in cleaned:
        # Only comma → decimal
        comma_pos = cleaned.rfind(',')
        decimal_places = len(cleaned) - comma_pos - 1
    elif '.' in cleaned:
        # Standard style or ambiguous
        dot_pos = cleaned.rfind('.')
        decimal_places = len(cleaned) - dot_pos - 1
        # If exactly 3 decimal places and no other dots, could be EUR thousands
        # but we treat as standard since no comma present
    else:
        decimal_places = 0

    if decimal_places > 0:
        return f"{num:,.{decimal_places}f}"
    else:
        return f"{num:,.0f}"


# ---------- COUNTRY CODE MAPPING ----------
COUNTRY_MAP: dict[str, str] = {
    "SK": "Slovakia",
    "CZ": "Czech Republic",
    "DE": "Germany",
    "TR": "Turkey",
    "HU": "Hungary",
    "JP": "Japan",
    "PT": "Portugal",
    "RO": "Romania",
    "ES": "Spain",
    "IT": "Italy",
    "FR": "France",
    "PL": "Poland",
    "AT": "Austria",
    "BE": "Belgium",
    "NL": "Netherlands",
    "SE": "Sweden",
    "GB": "United Kingdom",
    "CN": "China",
    "KR": "South Korea",
    "US": "United States",
    "MX": "Mexico",
    "BR": "Brazil",
    "IN": "India",
    "TH": "Thailand",
    "SI": "Slovenia",
    "RS": "Serbia",
    "BA": "Bosnia and Herzegovina",
    "HR": "Croatia",
    "BG": "Bulgaria",
    "TW": "Taiwan",
    "MY": "Malaysia",
    "ID": "Indonesia",
    "ZA": "South Africa",
    "FI": "Finland",
    "DK": "Denmark",
    "NO": "Norway",
    "CH": "Switzerland",
    "IE": "Ireland",
    "LU": "Luxembourg",
}


# ---------- CORE EXTRACTION LOGIC ----------
def extract_skoda_aftersales_invoice(pdf_path: str) -> dict:
    """
    Extract all line-item data from a single Skoda AG After Sales invoice PDF.

    Invoice layout (Page 1):
        Header rows, then column headers:
            Order No. HS code Quant UoM Unit price Total price
            Wrap. No./Orig.country Name of Goods Wgt./Unit Reference

        Each line item is a PAIR of lines:
            Line 1: PartNumber(with spaces) HSCode Quantity UoM UnitPrice TotalPrice
            Line 2: WrapNo(s)/CountryCode Description Wgt./Unit Reference
            (Optional extra wrap number lines)

    Returns a dict with header info and a list of line items.
    """
    invoice_number: str = ""
    invoice_date: str = ""
    currency: str = "EUR"
    line_items: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        all_text_lines: list[str] = []

        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text_lines.extend(text.split('\n'))

        # --- Extract Header Info ---
        for idx, line in enumerate(all_text_lines):
            stripped = line.strip()

            # Invoice Number: appears on a line by itself as a large number
            # Usually the 3rd or so line, right after "Rechnung" and before "Invoice"
            if not invoice_number:
                # Pattern: standalone 8-digit number (the invoice number)
                if re.match(r'^\d{7,10}$', stripped):
                    invoice_number = stripped

            # Invoice Date: "Date of taxable supply DD.MM.YYYY" or
            # "Den vystavení dokladu - Datum - Date DD.MM.YYYY"
            if not invoice_date:
                date_match = re.search(
                    r'(?:Date of taxable supply|Datum - Date|Date)\s+(\d{2}\.\d{2}\.\d{4})',
                    stripped
                )
                if date_match:
                    invoice_date = date_match.group(1)

            # Also try: "Erfüllungstag - Date of taxable supply DD.MM.YYYY"
            if not invoice_date:
                date_match = re.search(
                    r'Erfüllungstag.*?(\d{2}\.\d{2}\.\d{4})',
                    stripped
                )
                if date_match:
                    invoice_date = date_match.group(1)

            # Currency detection
            if 'EUR' in stripped:
                currency = "EUR"

        # --- Extract Line Items ---
        # The data section starts after the column headers line:
        #   "Order No. HS code Quant UoM Unit price Total price"
        #   "Wrap. No./Orig.country Name of Goods Wgt./Unit Reference"
        #
        # Each item is a pair of lines:
        #   Line 1: 567 857 705 B RAA 870821 2 PC 89.03 178.06
        #   Line 2: 208002329063/RO Three-point aut 1.155 0313191

        # Footer / non-data markers to skip
        skip_markers = [
            'Daňový doklad', 'Rechnung', 'Invoice', 'Kód dodávky',
            'Banka', 'BNP Paribas', 'Č. účtu', 'IBAN', 'BIC',
            'Příjemce', 'Kupující', 'Kaufer', 'Buyer',
            'SKODA AUTO', 'PRIVATE LIMITED', 'E-1, MIDC',
            'VILLAGE NIGOJE', 'CHAKAN TAL', 'PUNE',
            'Dodací list', 'Lieferschein', 'Advice',
            'CPT', 'Dodací podmínky', 'Lieferbedingungen',
            'Způsob dopravy', 'Transport', 'Platební',
            'Zahlungsbedingungen', 'Terms of payment',
            'Splatnost', 'Falligkeit', 'Due date',
            'Místo určení', 'Bestimmungsort', 'Destination',
            'Skoda Auto Volkswagen', 'CLC,E', 'Nigoje',
            'Datum uskut', 'Erfüllungstag', 'Den vystavení',
            'Order No.', 'Wrap. No.', 'Č.DOD.LISTU',
            'Total weight', 'Total price', 'Freight',
            'It is a tax-exempt', 'Es handelt sich',
            'Page ', 'Seite ', 'Strana ',
            'Škoda Auto', 'tř.Václava', 'Tř.Václava',
            'Mladá Boleslav', '293 01', 'IČO:',
            'Skoda Customer Care', '###',
            'State/Loading', 'Cust.overview',
            'Colli:', 'We invoice',
            'DELIVERIES ACCORDING', 'Country of origin',
            'VAV/2', 'road-marit', 'Městský',
            'Rechnungsempfänger', '----------',
            'Goods', 'Maharashtra', 'India',
            'Partner.spol', 'Náložní',
        ]

        # Pattern for Line 1 of a line item pair:
        # PartNumber (may contain spaces), then 6-digit HS code, Quantity, UOM, UnitPrice, TotalPrice
        # Example: "567 857 705 B RAA 870821 2 PC 89.03 178.06"
        # The part number can have letters, digits, spaces – ends right before the 6-digit HS code
        item_line_pattern = re.compile(
            r'^(.+?)\s+'          # Part number (non-greedy, stops before HS code)
            r'(\d{6})\s+'         # HS Code (exactly 6 digits)
            r'([\d.,]+)\s+'       # Quantity
            r'([A-Z]{1,3})\s+'    # UoM (PC, KG, L, etc.)
            r'([\d.,]+)\s+'       # Unit Price
            r'([\d.,]+)$'         # Total Price
        )

        # Pattern for Line 2 (wrap numbers / country code / description / weight):
        # Example: "208002329063/RO Three-point aut 1.155 0313191"
        # WrapNo(s)/CountryCode(2 letters)  Description  Weight  Reference
        detail_line_pattern = re.compile(
            r'^([\d/]+)/([A-Z]{2})\s+'   # Wrap numbers / Country code
            r'(.+?)\s+'                   # Description (non-greedy)
            r'([\d.,]+)\s+'               # Weight per unit
            r'(\d+)$'                     # Reference number
        )

        # Also handle detail lines where description has no reference at end
        detail_line_pattern_alt = re.compile(
            r'^([\d/]+)/([A-Z]{2})\s+'   # Wrap numbers / Country code
            r'(.+?)\s+'                   # Description (non-greedy)
            r'([\d.,]+)$'                 # Weight per unit (no reference)
        )

        i = 0
        while i < len(all_text_lines):
            line = all_text_lines[i].strip()

            # Skip blank lines and header/footer lines
            if not line:
                i += 1
                continue

            # Skip known non-data lines
            should_skip = False
            for marker in skip_markers:
                if line.startswith(marker):
                    should_skip = True
                    break
            if should_skip:
                i += 1
                continue

            # Skip standalone numbers that could be invoice number or other header data
            if re.match(r'^\d{7,10}$', line):
                i += 1
                continue

            # Skip lines that are just 2-digit numbers (like "05" page count, etc.)
            if re.match(r'^\d{1,2}$', line):
                i += 1
                continue

            # Skip country name lines (from page 2 summary)
            # e.g. "RO Romania"
            if re.match(r'^[A-Z]{2}\s+[A-Z][a-z]+', line) and not item_line_pattern.match(line):
                i += 1
                continue

            # Try to match item line (Line 1 of pair)
            match1 = item_line_pattern.match(line)
            if match1 and (i + 1) < len(all_text_lines):
                next_line = all_text_lines[i + 1].strip()
                match2 = detail_line_pattern.match(next_line)
                if not match2:
                    match2 = detail_line_pattern_alt.match(next_line)

                if match2:
                    # Extract from Line 1
                    raw_part_number = match1.group(1).strip()
                    # Remove spaces from part number
                    part_number = raw_part_number.replace(" ", "")
                    hs_code = match1.group(2).strip()
                    quantity_str = match1.group(3).strip()
                    uom = match1.group(4).strip()
                    unit_price_str = match1.group(5).strip()
                    total_price_str = match1.group(6).strip()

                    # Extract from Line 2
                    country_code = match2.group(2).strip()
                    description = match2.group(3).strip()
                    weight_str = match2.group(4).strip()

                    # Format numbers (EUR → standard)
                    formatted_unit_price = convert_eur_to_standard_format(unit_price_str)
                    formatted_total_price = convert_eur_to_standard_format(total_price_str)
                    formatted_weight = convert_eur_to_standard_format(weight_str)
                    formatted_quantity = convert_eur_to_standard_format(quantity_str)

                    # For quantity: if it looks like a whole number, format without decimals
                    try:
                        qty_val = eur_str_to_float(quantity_str)
                        if qty_val == int(qty_val):
                            formatted_quantity = f"{int(qty_val):,}"
                        else:
                            formatted_quantity = smart_format_number(quantity_str)
                    except (ValueError, OverflowError):
                        formatted_quantity = quantity_str

                    item = {
                        "Invoice Number": invoice_number,
                        "Invoice Date": invoice_date,
                        "Part Number": part_number,
                        "Description": description,
                        "Wgt./Unit": formatted_weight,
                        "Country Code": country_code,
                        "HS Code": hs_code,
                        "Default": "AUTOMOTIVE PARTS FOR CAPTIVE CONSUMPTION",
                        "Quantity": formatted_quantity,
                        "UOM": uom,
                        "Unit Price": formatted_unit_price,
                        "Total Price": formatted_total_price,
                        "Currency": currency,
                    }

                    line_items.append(item)

                    # Skip additional wrap number lines (standalone number lines after detail)
                    i += 2
                    while i < len(all_text_lines):
                        extra = all_text_lines[i].strip()
                        # Extra wrap number lines are purely numeric (e.g. "208002329064")
                        if re.match(r'^\d{8,15}$', extra):
                            i += 1
                        else:
                            break
                    continue

            i += 1

    return {
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "currency": currency,
        "items": line_items,
    }


# ---------- CSV OUTPUT ----------
def write_csv(output_path: str, all_records: list[dict]) -> None:
    """Write all extracted records to a single CSV file."""
    if not all_records:
        return

    fieldnames = [
        "Invoice Number",
        "Invoice Date",
        "Part Number",
        "Description",
        "Wgt./Unit",
        "Country Code",
        "HS Code",
        "Default",
        "Quantity",
        "UOM",
        "Unit Price",
        "Total Price",
        "Currency",
    ]

    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for record in all_records:
            # Preserve leading zeros in Part Number for Excel
            safe_record = dict(record)
            part_no = safe_record.get("Part Number", "")
            if part_no and part_no[0] == '0':
                safe_record["Part Number"] = f'="{part_no}"'
            writer.writerow(safe_record)


# ---------- NAGARKOT GUI IMPLEMENTATION ----------
class SkodaAfterSalesExtractorGUI:
    """Skoda AG After Sales Invoice Extractor – Nagarkot Branded GUI."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Skoda AG After Sales – Invoice Extractor")
        self.root.geometry("1200x750")
        self.root.state('zoomed')

        # Nagarkot brand palette
        self.bg_color = "#ffffff"
        self.brand_color = "#0056b3"
        self.root.configure(bg=self.bg_color)

        self.style = ttk.Style()
        self.style.theme_use('clam')

        # --- Style configuration ---
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure(
            "TLabel", background=self.bg_color, font=("Segoe UI", 10)
        )
        self.style.configure(
            "Header.TLabel",
            font=("Helvetica", 18, "bold"),
            foreground=self.brand_color,
            background=self.bg_color,
        )
        self.style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 11),
            foreground="gray",
            background=self.bg_color,
        )
        self.style.configure(
            "Footer.TLabel",
            font=("Segoe UI", 9),
            foreground="#555555",
            background=self.bg_color,
        )
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background=self.brand_color,
            foreground="white",
            borderwidth=0,
            focuscolor=self.brand_color,
        )
        self.style.map("Primary.TButton", background=[('active', '#004494')])
        self.style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            background="#f0f0f0",
            foreground="#333333",
            borderwidth=1,
        )
        self.style.map("Secondary.TButton", background=[('active', '#e0e0e0')])
        self.style.configure("TLabelframe", background=self.bg_color)
        self.style.configure(
            "TLabelframe.Label",
            background=self.bg_color,
            foreground=self.brand_color,
            font=("Segoe UI", 10, "bold"),
        )
        self.style.configure(
            "Treeview", font=("Segoe UI", 9), rowheight=25
        )
        self.style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            foreground=self.brand_color,
        )

        self.setup_ui()
        self.selected_files: list[str] = []

    # ----- UI SETUP -----
    def setup_ui(self) -> None:
        # ---------- HEADER ----------
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", pady=20, padx=20)
        header_frame.columnconfigure(0, weight=0)
        header_frame.columnconfigure(1, weight=1)
        header_frame.columnconfigure(2, weight=0)

        # Logo (Left)
        try:
            if Image and ImageTk:
                logo_path = resource_path("Nagarkot Logo.png")
                if os.path.exists(logo_path):
                    pil_img = Image.open(logo_path)
                    # Scale logo to height=20, preserve aspect ratio
                    h_target = 20
                    h_pct = h_target / float(pil_img.size[1])
                    w_size = int(float(pil_img.size[0]) * h_pct)
                    pil_img = pil_img.resize(
                        (w_size, h_target), Image.Resampling.LANCZOS
                    )
                    self.logo_img = ImageTk.PhotoImage(pil_img)
                    logo_lbl = ttk.Label(header_frame, image=self.logo_img)
                    logo_lbl.grid(
                        row=0, column=0, rowspan=2, sticky="w", padx=(0, 20)
                    )
                else:
                    print("Warning: Nagarkot Logo.png not found.")
                    ttk.Label(
                        header_frame, text="[LOGO]", foreground="gray"
                    ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
            else:
                ttk.Label(
                    header_frame, text="[PIL Missing]", foreground="red"
                ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))
        except Exception as e:
            print(f"Error loading logo: {e}")
            ttk.Label(
                header_frame, text="[LOGO ERROR]", foreground="red"
            ).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 20))

        # Title (Center)
        title_lbl = ttk.Label(
            header_frame,
            text="Skoda AG After Sales – Invoice Extractor",
            style="Header.TLabel",
        )
        title_lbl.grid(row=0, column=1, sticky="n")
        subtitle_lbl = ttk.Label(
            header_frame,
            text="Extract line-item data from Skoda AG After Sales invoices",
            style="Subtitle.TLabel",
        )
        subtitle_lbl.grid(row=1, column=1, sticky="n")

        # ---------- FOOTER (Packed first to reserve bottom space) ----------
        footer_frame = ttk.Frame(self.root, padding="10")
        footer_frame.pack(side="bottom", fill="x")

        copyright_lbl = ttk.Label(
            footer_frame,
            text="© Nagarkot Forwarders Pvt Ltd",
            style="Footer.TLabel",
        )
        copyright_lbl.pack(side="left", anchor="s")

        self.btn_run = ttk.Button(
            footer_frame,
            text="  Extract & Generate CSV  ",
            command=self.run_extraction,
            style="Primary.TButton",
        )
        self.btn_run.pack(side="right", padx=10, pady=5)

        # ---------- MAIN CONTENT ----------
        content_frame = ttk.Frame(self.root, padding="20 10 20 10")
        content_frame.pack(fill="both", expand=True)

        # --- File Selection ---
        file_frame = ttk.LabelFrame(
            content_frame, text="File Selection", padding="15"
        )
        file_frame.pack(fill="x", pady=(0, 15))

        btn_container = ttk.Frame(file_frame)
        btn_container.pack(fill="x")

        self.btn_select = ttk.Button(
            btn_container,
            text="Select PDFs",
            command=self.select_files,
            style="Secondary.TButton",
        )
        self.btn_select.pack(side="left", padx=(0, 10))

        self.btn_clear = ttk.Button(
            btn_container,
            text="Clear List",
            command=self.clear_files,
            style="Secondary.TButton",
        )
        self.btn_clear.pack(side="left")

        self.lbl_count = ttk.Label(
            btn_container, text="No files selected", style="TLabel"
        )
        self.lbl_count.pack(side="left", padx=(20, 0))

        # --- Output Settings ---
        output_frame = ttk.LabelFrame(
            content_frame, text="Output Settings", padding="15"
        )
        output_frame.pack(fill="x", pady=(0, 15))

        # --- Processing Mode (Combined vs Individual) ---
        ttk.Label(output_frame, text="Processing Mode:").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=5
        )

        mode_frame = ttk.Frame(output_frame)
        mode_frame.grid(row=0, column=1, columnspan=2, sticky="w")

        self.mode_var = tk.StringVar(value="combined")

        self.rb_combined = ttk.Radiobutton(
            mode_frame,
            text="Combined (All in one CSV)",
            variable=self.mode_var,
            value="combined",
            command=self.toggle_filename_state,
        )
        self.rb_combined.pack(side="left", padx=(0, 15))

        self.rb_individual = ttk.Radiobutton(
            mode_frame,
            text="Individual (Separate CSV per invoice)",
            variable=self.mode_var,
            value="individual",
            command=self.toggle_filename_state,
        )
        self.rb_individual.pack(side="left")

        # --- Output Folder ---
        ttk.Label(output_frame, text="Output Folder:").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self.output_dir_var = tk.StringVar()
        self.entry_output_dir = ttk.Entry(
            output_frame, textvariable=self.output_dir_var, width=50
        )
        self.entry_output_dir.grid(row=1, column=1, sticky="ew", padx=(0, 10))

        self.btn_browse_out = ttk.Button(
            output_frame,
            text="Browse...",
            command=self.browse_output_dir,
            style="Secondary.TButton",
        )
        self.btn_browse_out.grid(row=1, column=2, sticky="w")

        # --- Output Filename (shows name only, .csv appended automatically) ---
        ttk.Label(output_frame, text="Output Filename:").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self.output_name_var = tk.StringVar(
            value=f"Skoda_AfterSales_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        self.entry_output_name = ttk.Entry(
            output_frame, textvariable=self.output_name_var, width=50
        )
        self.entry_output_name.grid(row=2, column=1, sticky="ew", padx=(0, 10))

        self.lbl_filename_hint = ttk.Label(
            output_frame,
            text="(.csv added automatically)",
            foreground="gray",
        )
        self.lbl_filename_hint.grid(row=2, column=2, sticky="w")

        output_frame.columnconfigure(1, weight=1)

        # --- Data Preview ---
        preview_frame = ttk.LabelFrame(
            content_frame,
            text="Data Preview / Processing Queue",
            padding="15",
        )
        preview_frame.pack(fill="both", expand=True)

        cols = ("File Name", "Status", "Items", "Details")
        self.tree = ttk.Treeview(
            preview_frame, columns=cols, show="headings", selectmode="extended"
        )
        self.tree.heading("File Name", text="File Name")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Items", text="Items Found")
        self.tree.heading("Details", text="Details")

        self.tree.column("File Name", width=350, anchor="w")
        self.tree.column("Status", width=100, anchor="center")
        self.tree.column("Items", width=100, anchor="center")
        self.tree.column("Details", width=400, anchor="w")

        scrollbar_y = ttk.Scrollbar(
            preview_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")

        # --- Status Bar ---
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            content_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            foreground="#666666",
            background="#f5f5f5",
            anchor="w",
            padding="5 2",
        )
        status_bar.pack(fill="x", pady=(10, 0))

    # ----- FILE SELECTION -----
    def select_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select Skoda AG After Sales Invoice PDFs",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if files:
            self.selected_files = list(files)
            self.lbl_count.config(
                text=f"{len(self.selected_files)} file(s) selected"
            )
            # Clear and populate treeview
            for row in self.tree.get_children():
                self.tree.delete(row)
            for fpath in self.selected_files:
                self.tree.insert(
                    "", "end",
                    values=(os.path.basename(fpath), "Pending", "", "")
                )
            self.status_var.set(
                f"{len(self.selected_files)} file(s) loaded. "
                "Click 'Extract & Generate CSV' to process."
            )

            # Auto-set output folder if empty
            if not self.output_dir_var.get():
                first_dir = os.path.dirname(self.selected_files[0])
                self.output_dir_var.set(first_dir)

    def browse_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir_var.set(folder)

    def toggle_filename_state(self) -> None:
        """Enable/Disable filename entry based on mode."""
        if self.mode_var.get() == "individual":
            self.entry_output_name.config(state="disabled")
            self.lbl_filename_hint.config(text="(Auto-named by Invoice No.)")
        else:
            self.entry_output_name.config(state="normal")
            self.lbl_filename_hint.config(text="(.csv added automatically)")

    def clear_files(self) -> None:
        """Clear all selected files and reset output path/filename."""
        self.selected_files = []
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.lbl_count.config(text="No files selected")
        # Clear the output path and reset filename
        self.output_dir_var.set("")
        self.output_name_var.set(
            f"Skoda_AfterSales_Extracted_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        self.status_var.set("File list cleared.")

    # ----- RUN EXTRACTION -----
    def run_extraction(self) -> None:
        if not self.selected_files:
            messagebox.showwarning(
                "No Files", "Please select at least one PDF file."
            )
            return

        # Output setup
        out_dir = self.output_dir_var.get()
        if not out_dir:
            out_dir = os.path.dirname(self.selected_files[0])
            self.output_dir_var.set(out_dir)

        mode = self.mode_var.get()
        combined_records: list[dict] = []
        total_items = 0

        self.btn_run.config(state="disabled")
        self.btn_select.config(state="disabled")
        self.root.update_idletasks()

        tree_rows = self.tree.get_children()

        for idx, fpath in enumerate(self.selected_files):
            fname = os.path.basename(fpath)
            row_id = tree_rows[idx] if idx < len(tree_rows) else None

            try:
                self.status_var.set(f"Processing: {fname} ...")
                self.root.update_idletasks()

                result = extract_skoda_aftersales_invoice(fpath)
                items = result["items"]
                count = len(items)
                inv_no = result.get("invoice_number", "N/A")
                inv_date = result.get("invoice_date", "N/A")

                total_items += count

                # --- INDIVIDUAL MODE ---
                if mode == "individual" and items:
                    # Sanitize invoice number for filename
                    safe_inv = "".join(
                        c for c in inv_no if c.isalnum() or c in ('-', '_')
                    )
                    if safe_inv:
                        indiv_name = f"{safe_inv}.csv"
                    else:
                        base = os.path.splitext(fname)[0]
                        indiv_name = f"{base}_Extracted.csv"

                    indiv_path = os.path.join(out_dir, indiv_name)
                    write_csv(indiv_path, items)
                    detail_msg = f"Saved: {indiv_name} ({count} items)"

                # --- COMBINED MODE ---
                else:
                    combined_records.extend(items)
                    detail_msg = f"Invoice: {inv_no} | Date: {inv_date} | {count} items"

                if row_id:
                    self.tree.item(
                        row_id,
                        values=(fname, "✓ Done", str(count), detail_msg),
                    )

            except Exception as e:
                if row_id:
                    self.tree.item(
                        row_id,
                        values=(fname, "✗ Error", "0", str(e)),
                    )
                self.status_var.set(f"Error processing {fname}: {e}")

            self.root.update_idletasks()

        # Finalize Combined Mode
        if mode == "combined":
            if combined_records:
                out_name = self.output_name_var.get().strip()
                if not out_name:
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    out_name = f"Skoda_AfterSales_Extracted_{timestamp}"

                # Strip any existing .csv to avoid double extension
                if out_name.lower().endswith(".csv"):
                    out_name = out_name[:-4]
                # Re-add .csv
                out_name += ".csv"

                output_path = os.path.join(out_dir, out_name)
                try:
                    write_csv(output_path, combined_records)
                    messagebox.showinfo(
                        "Success",
                        f"Combined extraction complete!\n\n"
                        f"Total items: {total_items}\n"
                        f"Saved to: {output_path}"
                    )
                    self.status_var.set(f"Done. Saved to {out_name}")
                except Exception as e:
                    messagebox.showerror(
                        "Error", f"Could not write combined CSV:\n{e}"
                    )
            else:
                self.status_var.set("No data found to combine.")
                if total_items == 0:
                    messagebox.showwarning(
                        "No Data", "No items extracted from selected files."
                    )

        # Finalize Individual Mode
        else:
            messagebox.showinfo(
                "Success",
                f"Individual extraction complete!\n\n"
                f"Processed {len(self.selected_files)} files.\n"
                f"Total items found: {total_items}\n"
                f"Folder: {out_dir}"
            )
            self.status_var.set(f"Done. Files saved to {out_dir}")

        # Refresh timestamp for next run
        new_ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if self.mode_var.get() == "combined":
            self.output_name_var.set(f"Skoda_AfterSales_Extracted_{new_ts}")

        self._reset_buttons()

    def _reset_buttons(self) -> None:
        self.btn_run.config(state="normal")
        self.btn_select.config(state="normal")

    def run(self) -> None:
        self.root.mainloop()


# ---------- ENTRY POINT ----------
if __name__ == "__main__":
    app = SkodaAfterSalesExtractorGUI()
    app.run()
