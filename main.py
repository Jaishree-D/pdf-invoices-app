import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # print(df["product_id"])
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    # Set the Header
    pdf.set_font(family="Times", style="B", size=12)
    # pdf.set_text_color(0, 0, 254)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", align='L', ln=1)
    pdf.output(f"PDFs/{filename}.pdf")
# print(pdf.output())
