"""
PO Extractor - simple desktop app version (v2).

Lets the user pick a folder of PO PDFs, extracts the fields below, and
saves an Excel file next to the input folder. No command line needed.

Output columns:
    S.No | PO Date | PO No | Customer Name |
    Item Desc (prefix before first letter) | Item Desc (remainder) |
    PO Qty | Rate
"""

import re
import os
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

import pdfplumber
import openpyxl
from openpyxl.styles import Font

PO_NO_PATTERN = re.compile(r"PO No\s*:\s*(\S+)")
PO_DATE_PATTERN = re.compile(r"PO Date\s*:\s*(\S+)")
ITEM_LINE_PATTERN = re.compile(
    r"^\s*\d+\s+(.+?)\s+\d{6,8}\s+([\d.]+)\s+\w+\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$"
)
# Leading run of non-alphabet characters (digits, symbols, spaces) before the
# first alphabet character in the description.
LEADING_NON_ALPHA_PATTERN = re.compile(r"^([^A-Za-z]*)(.*)$", re.DOTALL)


def extract_company_name(text):
    for line in text.splitlines():
        line = line.strip()
        if "Limited" in line or "Ltd" in line:
            return line
    return ""


def split_description(description):
    """
    Split into (prefix, remainder):
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


def extract_fields_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    company_name = extract_company_name(full_text)
    po_no_match = PO_NO_PATTERN.search(full_text)
    po_date_match = PO_DATE_PATTERN.search(full_text)
    po_no = po_no_match.group(1) if po_no_match else ""
    po_date = po_date_match.group(1) if po_date_match else ""

    rows = []
    for line in full_text.splitlines():
        match = ITEM_LINE_PATTERN.match(line)
        if match:
            description, qty, rate, disc_amt, amount = match.groups()
            prefix, remainder = split_description(description.strip())
            rows.append({
                "PO Date": po_date,
                "PO No": po_no,
                "Customer Name": company_name,
                "Item Desc (Prefix)": prefix,
                "Item Desc (Description)": remainder,
                "PO Qty": float(qty),
                "Rate": float(rate),
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
        for filename in pdf_files:
            pdf_path = os.path.join(folder, filename)
            try:
                rows = extract_fields_from_pdf(pdf_path)
                if not rows:
                    self.write_log(f"⚠  No item lines matched in: {filename}")
                all_rows.extend(rows)
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

        headers = ["S.No"] + list(all_rows[0].keys())
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(name="Arial", bold=True)

        for i, row in enumerate(all_rows, start=1):
            ws.append([i] + [row[h] for h in list(all_rows[0].keys())])
            for cell in ws[ws.max_row]:
                cell.font = Font(name="Arial")

        widths = [6, 12, 16, 30, 16, 30, 10, 10]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        wb.save(output_path)

        self.write_log(f"\nDone! Saved {len(all_rows)} rows to:\n{output_path}")
        self._finish()

    def _finish(self):
        self.run_button.config(state="normal", text="Run Extraction")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
