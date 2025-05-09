#!/usr/bin/env python
# coding: utf-8

import pytesseract
import numpy as np
from PIL import Image
import os
import re
from typing import List, Dict
import csv
import json
import sqlite3
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import cv2
from pdf2image import convert_from_path
from io import BytesIO
from PIL import ImageTk
import threading


os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
custom_oem_psm_config = r'--oem 3 --psm 12'


# ## OCR extraction functions

def preprocess_image_cv(image_path):
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    _, binary = cv2.threshold(blur, 150, 255, cv2.THRESH_BINARY)
    return Image.fromarray(binary)


def extract_text(image: Image.Image) -> str:
    try:
        return pytesseract.image_to_string(image, lang='eng', config=custom_oem_psm_config)
    except Exception as e:
        print(f"OCR Error: {e}")
        return ""


def extract_invoice_info(text: str) -> Dict[str, str]:
    try:
        bill_match = re.search(r"Invoice no:\s*(\d+)", text)
        bill_id = bill_match.group(1) if bill_match else "Not found"

        date_match = re.search(r"Date of issue:?\s*\n\s*(\d{2}/\d{2}/\d{4})", text)
        date = date_match.group(1) if date_match else "Not found"

        client_match = re.search(r"Client:\s*\n\s*[^\n]+\s*\n\s*([^\n]+)", text)
        client_name = client_match.group(1).strip() if client_match else "Not found"

        amount_match = re.search(r"Total\s*\$\s*[\d\s]+,\d+\s*\$\s*[\d\s]+,\d+\s*\$\s*([\d\s]+,\d+)", text)
        amount = amount_match.group(1) if amount_match else "Not found"

        VAT_match = re.search(r"Total\s*\$\s*[\d\s]+,\d+\s*\$\s*([\d\s]+,\d+)", text)
        VAT = VAT_match.group(1) if VAT_match else "Not found"

        return {
            "Bill Number": bill_id,
            "Date": date,
            "Client Name": client_name,
            "Total Amount": amount,
            "VAT": VAT
        }
    except Exception as e:
        print(f"Error extracting invoice info: {e}")
        return {}


# ## Export functions

def save_to_csv(results, output_path):
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ["Bill Number", "Date", "Client Name", "Total Amount", "VAT"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for data in results:
            writer.writerow(data)

def save_to_json(results, output_path):
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=4)

def save_to_sqlite(results, output_path):
    conn = sqlite3.connect(output_path)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS invoices (
            bill_number TEXT,
            date TEXT,
            client_name TEXT,
            total_amount TEXT,
            vat TEXT
        )
    """)
    for data in results:
        c.execute("INSERT INTO invoices VALUES (?, ?, ?, ?, ?)",
                  (data['Bill Number'], data['Date'], data['Client Name'], data['Total Amount'], data['VAT']))
    conn.commit()
    conn.close()


# ## GUI

class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Extractor")
        self.root.geometry("900x600")
        self.root.configure(bg="#f0f4f7")

        self.results = []
        self.create_widgets()

    def create_widgets(self):
        title = tk.Label(self.root, text="Invoice Information Extractor", font=("Helvetica", 18, "bold"), fg="#333", bg="#f0f4f7")
        title.pack(pady=10)

        btn_frame = tk.Frame(self.root, bg="#f0f4f7")
        btn_frame.pack(pady=10)

        self.select_btn = tk.Button(btn_frame, text="Select Files", command=self.load_files, bg="#4CAF50", fg="white", padx=15)
        self.select_btn.pack(side=tk.LEFT, padx=10)

        self.export_btn = tk.Button(btn_frame, text="Export", command=self.export_dialog, bg="#2196F3", fg="white", padx=15)
        self.export_btn.pack(side=tk.LEFT, padx=10)

        # Progress bar
        progress_frame = tk.Frame(self.root, bg="#f0f4f7")
        progress_frame.pack(pady=10, fill=tk.X, padx=20)
        
        self.progress_label = tk.Label(progress_frame, text="Ready", bg="#f0f4f7")
        self.progress_label.pack()
        
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X)

        # Table
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(table_frame, columns=("Bill Number", "Date", "Client Name", "Total Amount", "VAT"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=160)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

    def load_files(self):
    # Start thread so GUI doesn't freeze
        thread = threading.Thread(target=self._process_files)
        thread.start()

    def _process_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Images and PDF", "*.jpg *.png *.tiff *.pdf")])
        if not file_paths:
            return

        total_files = sum(len(convert_from_path(path)) if path.lower().endswith(".pdf") else 1 for path in file_paths)
        self.results.clear()
        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0
        self.progress_label.config(text="Processing started...")
        self.tree.delete(*self.tree.get_children())

        processed = 0

        for path in file_paths:
            images = convert_from_path(path) if path.lower().endswith(".pdf") else [preprocess_image_cv(path)]
            
            for img in images:
                text = extract_text(img)
                info = extract_invoice_info(text)
                self.results.append(info)
                self.tree.insert("", "end", values=(info["Bill Number"], info["Date"], info["Client Name"], info["Total Amount"], info["VAT"]))
                processed += 1
                percent = int((processed / total_files) * 100)
                self.progress_bar["value"] = processed
                self.progress_label.config(text=f"{processed}/{total_files}  image processed         {percent}%")
                self.root.update_idletasks()

        self.progress_label.config(text="Processing complete ✅")
    
    
    def export_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Export Options")
        dialog.geometry("300x250")

        var_csv = tk.BooleanVar()
        var_json = tk.BooleanVar()
        var_sqlite = tk.BooleanVar()

        tk.Checkbutton(dialog, text="Export to CSV", variable=var_csv).pack(anchor='w', padx=20, pady=5)
        tk.Checkbutton(dialog, text="Export to JSON", variable=var_json).pack(anchor='w', padx=20, pady=5)
        tk.Checkbutton(dialog, text="Export to SQLite", variable=var_sqlite).pack(anchor='w', padx=20, pady=5)

        tk.Label(dialog, text="File name (without extension):").pack(pady=5)
        filename_entry = tk.Entry(dialog)
        filename_entry.pack(pady=5)

        def export():
            folder = filedialog.askdirectory()
            if not folder:
                return
            base = os.path.join(folder, filename_entry.get())
            if var_csv.get():
                save_to_csv(self.results, base + ".csv")
            if var_json.get():
                save_to_json(self.results, base + ".json")
            if var_sqlite.get():
                save_to_sqlite(self.results, base + ".db")
            messagebox.showinfo("Export", "Export completed successfully!")
            dialog.destroy()

        tk.Button(dialog, text="Export", command=export, bg="#4CAF50", fg="white").pack(pady=10)


# ## Main

if __name__ == '__main__':
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()




