import pandas as pd
import glob
import openpyxl
from fpdf import FPDF
from pathlib import Path

files = glob.glob("invoices/*.xlsx")

for file in files:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    file_name = Path(file).stem
    invoice_num = file_name.split("-")[0]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Number {invoice_num}")
    pdf.output(f"pdfs/{file_name}.pdf")
