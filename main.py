import pandas as pd
import glob
from fpdf import FPDF
import pathlib as path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = path.Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Helvetica", size=16, style="IB")
    pdf.cell(w=0, h=10, txt=f"Invoice_no.{invoice_nr}", ln=1)

    pdf.set_font(family="Helvetica", size=16, style="IB")
    pdf.cell(w=0, h=10, txt=f"Invoice_no.{date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add the header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]

    pdf.set_font(family="Times", size=10, style="IB")

    pdf.cell(w=30, h=10, txt=str(columns[0]), border=1)
    pdf.cell(w=60, h=10, txt=str(columns[1]), border=1)
    pdf.cell(w=30, h=10, txt=str(columns[2]), border=1)
    pdf.cell(w=30, h=10, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=10, txt=str(columns[4]), border=1, ln=1)

    for index,rows in df.iterrows():
        pdf.set_font(family="Times", size=8, style="I")
        pdf.set_text_color(90, 90, 90)

        pdf.cell(w=30, h=10, txt=str(rows["product_id"]), border=1)
        pdf.cell(w=60, h=10, txt=str(rows["product_name"]), border=1)
        pdf.cell(w=30, h=10, txt=str(rows["amount_purchased"]), border=1)
        pdf.cell(w=30, h=10, txt=str(rows["price_per_unit"]), border=1)
        pdf.cell(w=30, h=10, txt=str(rows["total_price"]), border=1, ln=1)

    total_price = df["total_price"].sum()
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=60, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt=f"{total_price}", border=1, ln=1)

    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=10, txt=f"The total price is {total_price}", ln=1)

    pdf.output(f"PDFs/{filename}.pdf")


