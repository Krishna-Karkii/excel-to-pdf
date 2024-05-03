import pandas as pd
import glob
from fpdf import FPDF
import pathlib as path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = path.Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    pdf.set_font(family="Helvetica", size=20, style="IB")
    pdf.cell(w=0, h=10, txt=f"Invoice_no.{invoice_nr}")
    pdf.output(f"PDFs/{filename}.pdf")


