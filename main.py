import os
from fpdf import FPDF
import pandas as pd

invoices = os.listdir('Excel Invoices')

for invoice in invoices:
    print(f'Excel Invoices/{invoice}')
    invoice_df = pd.read_excel(f'Excel Invoices/{invoice}')

    pdf = FPDF(orientation='P', unit='mm', format='A4')

    pdf.add_page()

    pdf.set_font(family="Times", size=12)
    pdf.set_text_color(110,110,110)

    pdf.cell(0, 12, txt=f'Invoice {invoice}')

    pdf.output(invoice.replace('.xlsx', '.pdf'))