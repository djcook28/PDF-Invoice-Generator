import glob
from fpdf import FPDF
import pandas as pd
from pathlib import Path

filepaths = glob.glob('Excel Invoices/*.xlsx')

for filepath in filepaths:
    invoice_df = pd.read_excel(filepath, sheet_name='Sheet 1')

    filename = Path(filepath).stem
    filename_split = filename.split("-")
    invoice_number = filename_split[0]
    invoice_date = filename_split[1]

    pdf = FPDF(orientation='P', unit='mm', format='A4')

    pdf.add_page()

    pdf.set_font(family="Times", size=12)
    pdf.set_text_color(110,110,110)

    pdf.cell(50, 12, txt=f'Invoice nr. {invoice_number}')
    pdf.cell(0, 12, txt=f'Date {invoice_date}')

    pdf.output(f"PDFs/{filename}.pdf")