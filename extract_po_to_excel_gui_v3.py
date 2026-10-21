"""
PO Extractor - simple desktop app version (v3).

Supports TWO vendor PO formats, auto-detected per PDF:

  FORMAT A - "Accucut"-style (labels: "PO No", "PO Date")
      Item rows read line-by-line with a regex.

  FORMAT B - "Bull Agro"-style (labels: "Order No.", "Order Date")
      Item rows read from the PDF's table grid (handles multi-line
      descriptions and the extra tax columns in this layout).
      Order Date -> written into the "PO Date" column
      Order No.  -> written into the "PO No" column

Lets the user pick a folder of PO PDFs, extracts the fields below, and
saves an Excel file next to the input folder. No command line needed.

Output columns:
    S.No | PO Date | PO No | Customer Name |
    Item Code (prefix before first letter) | Item Description (remainder) |
    PO Qty | Rate
"""

import re
import os
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment

# ---------- Format A ("PO No" / "PO Date") ----------
PO_NO_PATTERN = re.compile(r"PO No\s*:\s*(\S+)")
PO_DATE_PATTERN = re.compile(r"PO Date\s*:\s*(\S+)")
ITEM_LINE_PATTERN = re.compile(
    r"^\s*\d+\s+(.+?)\s+(?:\d{6,8}\s+)?([\d.]+)\s+\w+\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$"
)

# ---------- Format B ("Order No." / "Order Date") ----------
ORDER_NO_PATTERN = re.compile(r"Order No\.?\s*:\s*(.+)")
ORDER_DATE_PATTERN = re.compile(r"Order Date\s*:\s*(.+)")

# Leading run of non-alphabet characters (digits, symbols, spaces) before the
# first alphabet character in a description - used for both formats.
LEADING_NON_ALPHA_PATTERN = re.compile(r"^([^A-Za-z]*)(.*)$", re.DOTALL)

# Matches a day/month/year date regardless of separator (/, ., -, space).
DATE_PATTERN = re.compile(r"(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{4})")


def normalize_date(raw_date):
    """Normalize any dd/mm/yyyy-style date (with /, ., or - separators)
    to a consistent dd/mm/yyyy string. Returns the original text unchanged
    if it doesn't look like a recognizable date."""
    if not raw_date:
        return raw_date
    match = DATE_PATTERN.search(raw_date)
    if not match:
        return raw_date
    day, month, year = match.groups()
    return f"{int(day):02d}/{int(month):02d}/{year}"


def extract_company_name(text):
    """Works for both formats: prefer a 'Limited'/'Ltd' line, else fall
    back to the first ALL-CAPS line that has no digits (e.g. a masthead
    like 'BULL AGRO IMPLEMENTS')."""
    for line in text.splitlines():
        stripped = line.strip()
        if "Limited" in stripped or "Ltd" in stripped:
            return stripped
    for line in text.splitlines():
        stripped = line.strip()
        if stripped and stripped.isupper() and not any(ch.isdigit() for ch in stripped):
            return stripped
    return ""


def split_description(description):
    """
    Used for FORMAT A (Accucut). Split into (prefix, remainder):
      - prefix  = characters before the first alphabet character
      - remainder = everything from the first alphabet character onward
      - if the description already starts with a letter, prefix is left
        blank and remainder holds the full description.
    """
    match = LEADING_NON_ALPHA_PATTERN.match(description)
    prefix = (match.group(1) or "").strip()
    remainder = (match.group(2) or "").strip()
    if prefix == "":
        return "", description.strip()
    return prefix, remainder


def split_description_hyphen(description):
    """
    Used for FORMAT B (Bull Agro). Split into (prefix, remainder):
      - prefix = everything before the first hyphen "-"
      - remainder = everything after that first hyphen
      - if there's no hyphen at all, prefix is left blank and remainder
        holds the full description
    """
    stripped = description.strip()
    if "-" in stripped:
        idx = stripped.index("-")
        prefix = stripped[:idx].strip()
        remainder = stripped[idx + 1:].strip()
        return prefix, remainder
    return "", stripped


def find_col_index(header_row, target):
    """Find a column index in a table header row by fuzzy label match."""
    for i, cell in enumerate(header_row):
        if cell and target.lower() in cell.replace("\n", " ").strip().lower():
            return i
    return None


def detect_format(full_text):
    if re.search(r"PO No\s*:", full_text):
        return "A"
    if re.search(r"Order No\.?\s*:", full_text):
        return "B"
    return None


def extract_format_a(full_text):
    po_no_match = PO_NO_PATTERN.search(full_text)
    po_date_match = PO_DATE_PATTERN.search(full_text)
    po_no = po_no_match.group(1) if po_no_match else ""
    po_date = po_date_match.group(1) if po_date_match else ""

    items = []
    for line in full_text.splitlines():
        match = ITEM_LINE_PATTERN.match(line)
        if match:
            description, qty, rate, disc_amt, amount = match.groups()
            items.append((description.strip(), float(qty), float(rate)))
    return po_no, po_date, items


def extract_format_b(full_text, tables):
    order_no_match = ORDER_NO_PATTERN.search(full_text)
    order_date_match = ORDER_DATE_PATTERN.search(full_text)
    order_no = order_no_match.group(1).strip() if order_no_match else ""
    order_date = order_date_match.group(1).strip() if order_date_match else ""

    items = []
    for table in tables:
        if not table or len(table) < 3:
            continue
        header_row = table[0]
        if not header_row or not any(c and "description" in str(c).lower() for c in header_row):
            continue

        desc_idx = find_col_index(header_row, "description")
        qty_idx = find_col_index(header_row, "ord. qty") or find_col_index(header_row, "qty")
        rate_idx = find_col_index(header_row, "unit cost")
        if desc_idx is None or qty_idx is None or rate_idx is None:
            continue

        for row in table[2:]:
            sl = row[0]
            if sl and str(sl).strip().isdigit():
                raw_desc = row[desc_idx] or ""
                # Strip an "HSN/SAC :..." line that sometimes shares the cell
                clean_desc = re.sub(r"HSN/SAC.*", "", raw_desc, flags=re.IGNORECASE)
                clean_desc = " ".join(clean_desc.split())
                qty_raw = (row[qty_idx] or "0").replace(",", "")
                rate_raw = (row[rate_idx] or "0").replace(",", "")
                try:
                    qty = float(qty_raw)
                    rate = float(rate_raw)
                except ValueError:
                    continue
                items.append((clean_desc, qty, rate))
    return order_no, order_date, items


def extract_fields_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        fmt = detect_format(full_text)

        if fmt == "A":
            po_no, po_date, items = extract_format_a(full_text)
            split_func = split_description
        elif fmt == "B":
            tables = []
            for page in pdf.pages:
                tables.extend(page.extract_tables())
            order_no, order_date, items = extract_format_b(full_text, tables)
            po_date, po_no = order_date, order_no
            split_func = split_description_hyphen
        else:
            raise ValueError("Unrecognized PO format (no 'PO No'/'PO Date' or "
                              "'Order No.'/'Order Date' labels found)")

    company_name = extract_company_name(full_text)
    po_date = normalize_date(po_date)

    rows = []
    for description, qty, rate in items:
        prefix, remainder = split_func(description)
        rows.append({
            "PO Date": po_date,
            "PO No": po_no,
            "Customer Name": company_name,
            "Item Code": prefix,
            "Item Description": remainder,
            "PO Qty": qty,
            "Rate": rate,
        })
    return rows


class App:
    def __init__(self, root):
        self.root = root
        root.title("PO to Excel Extractor")
        root.geometry("560x420")
        root.resizable(False, False)

        self.folder_path = tk.StringVar()

        tk.Label(root, text="1. Choose the folder containing your PO PDFs:",
                 font=("Arial", 10, "bold")).pack(anchor="w", padx=12, pady=(14, 2))

        frame = tk.Frame(root)
        frame.pack(fill="x", padx=12)
        tk.Entry(frame, textvariable=self.folder_path).pack(side="left", fill="x", expand=True)
        tk.Button(frame, text="Browse...", command=self.browse_folder).pack(side="left", padx=6)

        tk.Label(root, text="2. Click Run to extract every PDF in that folder:",
                 font=("Arial", 10, "bold")).pack(anchor="w", padx=12, pady=(16, 2))

        self.run_button = tk.Button(root, text="Run Extraction", bg="#2e7d32", fg="white",
                                     font=("Arial", 11, "bold"), command=self.run_extraction)
        self.run_button.pack(padx=12, pady=4, fill="x")

        tk.Label(root, text="Log:").pack(anchor="w", padx=12, pady=(12, 0))
        self.log = scrolledtext.ScrolledText(root, height=12, state="disabled")
        self.log.pack(fill="both", expand=True, padx=12, pady=(2, 12))

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder with PO PDFs")
        if folder:
            self.folder_path.set(folder)

    def write_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def run_extraction(self):
        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("No folder selected", "Please choose a valid folder first.")
            return
        self.run_button.config(state="disabled", text="Running...")
        threading.Thread(target=self._extract_worker, args=(folder,), daemon=True).start()

    def _extract_worker(self, folder):
        pdf_files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
        if not pdf_files:
            self.write_log("No PDF files found in that folder.")
            self._finish()
            return

        all_rows = []
        group_sizes = []  # number of item rows contributed by each successfully processed PDF
        for filename in pdf_files:
            pdf_path = os.path.join(folder, filename)
            try:
                rows = self._safe_extract(pdf_path)
                if not rows:
                    self.write_log(f"⚠  No item lines matched in: {filename}")
                else:
                    all_rows.extend(rows)
                    group_sizes.append(len(rows))
                self.write_log(f"✔ Processed: {filename} ({len(rows)} item rows)")
            except Exception as e:
                self.write_log(f"✘ Failed on {filename}: {e}")

        if not all_rows:
            self.write_log("Nothing extracted. No Excel file was created.")
            self._finish()
            return

        output_path = os.path.join(folder, "po_extract.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PO Extract"

        data_keys = list(all_rows[0].keys())
        headers = ["S.No"] + data_keys + ["Dispatch Qty", "Balance Qty"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(name="Arial", bold=True)

        # Columns that get merged + centered per source PDF (S.No, PO Date,
        # PO No, Customer Name are always the first four columns).
        MERGE_COLS = [1, 2, 3, 4]
        center_align = Alignment(horizontal="center", vertical="center")

        row_index = 0
        for s_no, size in enumerate(group_sizes, start=1):
            for _ in range(size):
                row = all_rows[row_index]
                ws.append([s_no] + [row[h] for h in data_keys] + [None, None])
                for cell in ws[ws.max_row]:
                    cell.font = Font(name="Arial")
                row_index += 1

        # Apply centering + merging per PDF group. Data starts at row 2.
        current_row = 2
        for size in group_sizes:
            start_row = current_row
            end_row = current_row + size - 1
            for col in MERGE_COLS:
                for r in range(start_row, end_row + 1):
                    ws.cell(row=r, column=col).alignment = center_align
                if end_row > start_row:
                    ws.merge_cells(start_row=start_row, start_column=col,
                                    end_row=end_row, end_column=col)
            current_row = end_row + 1

        widths = [6, 12, 16, 30, 16, 30, 10, 10, 12, 12]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        wb.save(output_path)

        self.write_log(f"\nDone! Saved {len(all_rows)} rows to:\n{output_path}")
        self._finish()

    def _safe_extract(self, pdf_path):
        return extract_fields_from_pdf(pdf_path)

    def _finish(self):
        self.run_button.config(state="normal", text="Run Extraction")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
