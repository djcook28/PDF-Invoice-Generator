import glob
from fpdf import FPDF
import pandas as pd

filepaths = glob.glob('Excel Invoices/*.xlsx')

for filepath in filepaths:
    invoice_df = pd.read_excel(filepath, sheet_name='Sheet 1')
    print(filepath)

    start_of_filename = filepath.find("\\")
    start_of_date = filepath.find("-")
    end_of_date = filepath.find(".xlsx")
    invoice_number = filepath[start_of_filename+1:start_of_date]
    invoice_date = filepath[start_of_date+1:end_of_date]
    print((invoice_date))

    pdf = FPDF(orientation='P', unit='mm', format='A4')

    pdf.add_page()

    pdf.set_font(family="Times", size=12)
    pdf.set_text_color(110,110,110)

    pdf.cell(50, 12, txt=f'Invoice nr. {invoice_number}')
    pdf.cell(0, 12, txt=f'Date {invoice_date}')

    pdf.output(filepath.replace('.xlsx', '.pdf'))