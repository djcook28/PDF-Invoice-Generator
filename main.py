import glob
from fpdf import FPDF
import pandas as pd
from pathlib import Path

filepaths = glob.glob('Excel Invoices/*.xlsx')

for filepath in filepaths:
    filename = Path(filepath).stem

    #split divides the string into a list of strings which can either be assigned to a variable or be
    # assigned directly to variables
    #method 1
    # filename_split = filename.split("-")
    # invoice_number = filename_split[0]
    # invoice_date = filename_split[1]
    #method 2
    invoice_number, invoice_date = filename.split("-")

    pdf = FPDF(orientation='P', unit='mm', format='A4')

    pdf.add_page()

    pdf.set_font(family="Times", size=12)
    pdf.set_text_color(110,110,110)

    pdf.cell(50, 12, txt=f'Invoice nr. {invoice_number}', ln=1)
    pdf.cell(0, 12, txt=f'Date {invoice_date}', ln=1)

    invoice_df = pd.read_excel(filepath, sheet_name='Sheet 1')

    column_headers = list(invoice_df.columns)
    pdf.set_font(style='B', family="Times", size=8)

    i = 0
    for header in column_headers:

        header = header.replace("_", " ")
        header = header.title()

        if i == len(column_headers)-1:
            pdf.cell(w=40, h=8, txt=str(header), border=1, ln=1)
        else:
            pdf.cell(w=40, h=8, txt=str(header), border=1)
        i = i+1

    for index, row in invoice_df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.cell(w=40, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=40, h=8, txt=row["product_name"], border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total_sum = invoice_df["total_price"].sum()
    pdf.set_font(family="Times", size=8)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt=str(total_sum), border=1, ln=1)

    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=10, txt=f"The total invoice amount is {total_sum}", ln=1)
    pdf.cell(w=20, h=10, txt="PythonHow")
    pdf.image("pythonhow.png", w=10, h=10)

    pdf.output(f"PDFs/{filename}.pdf")