import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1", index_col=False)

    #print(df.get("product_id"))
    #print(filepath, df)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    invoice_date = filename.split("-")[1]
    # Set the Header
    pdf.set_font(family="Times", style="B", size=12)
    # pdf.set_text_color(0, 0, 254)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", align='L', ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {invoice_date}", align='L', ln=1)
    # table header
    pdf.set_font(family="Times", style="B", size=11)
    pdf.set_text_color(80, 80, 80)
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=50, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times",  size=10)
        pdf.set_text_color(80, 80, 80)
        # Add row to the table
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)
        total_sum = df["total_price"].sum()
        # print(tot)
    pdf.cell(w=30, h=8, txt="Total Price", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=f"The total due amount is {total_sum} Euros", border=0, ln=1)
    pdf.cell(w=20, h=8, txt=f"PythonHow", border=0)
    pdf.image("invoices/pythonhow.png", w=5, h=5)
    pdf.output(f"PDFs/{filename}.pdf")
# print(pdf.output())
