import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths=glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    filename=Path(filepath).stem #string formatina dosturuyor
    invoice_nr,date=filename.split("-") #list'e donusturuyor ve ilk elemanini aliyoruz.

    pdf=FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()
    pdf.set_font(family="Times",size=16,style="B")
    pdf.cell(w=50,h=8,txt=f"Invoices nr. {invoice_nr}",ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {date}",ln=1)
    pdf.cell(w=50, h=8, txt="", ln=1)
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #Adding Header:
    columns=df.columns
    columns=[item.replace("_"," ").title() for item in columns]
    pdf.set_font(family="Times",style="B" ,size=10)
    pdf.set_text_color(200, 100, 100)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=25, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1,ln=1)


    for index,row in df.iterrows():
        pdf.set_font(family="Times",size=10)
        pdf.set_text_color(100,100,100)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=25, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1,ln=1)

    pdf.output(f"PDFs/{filename}.pdf")


