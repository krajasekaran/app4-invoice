import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

files = glob.glob("invoices/*.xlsx")
print(files)

# df = pd.read_excel(file, sheet_name="Sheet 1")
# print(df)

for file in files:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    filename = Path(file).stem
    invoice_nr = (Path(file).stem).split('-')[0]
    print(f"invoice number is {invoice_nr}")
    date = (Path(file).stem).split('-')[1]
    print(f"purchase date is {date}")

    pdf.add_page()
    pdf.set_font(family='arial', style='BI', size=30)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=15, txt='Invoice', border=0, ln=1, align='C')
    pdf.set_line_width(width=2)
    pdf.line(10, 30, 200, 30)


    pdf.ln(30)
    pdf.set_font(family='arial',style='BI',size=10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=10, txt=f"Invoice number: {invoice_nr}", ln=1, align='L')
    pdf.cell(w=0, h=10, txt=f"Date: {date}", ln=1, align='L')


    pdf.ln(40)
    df = pd.read_excel(file, sheet_name="Sheet 1")
    columns = list(df.columns)
    pdf.set_line_width(width=.5)
    pdf.set_font(family='arial',style='BI', size=10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=10, txt=columns[0].title(), ln=0, align='L', border=1)
    pdf.cell(w=60, h=10, txt=columns[1].title(), ln=0, align='L', border=1)
    pdf.cell(w=40, h=10, txt=columns[2].title(), ln=0, align='L', border=1)
    pdf.cell(w=30, h=10, txt=columns[3].title(), ln=0, align='L', border=1)
    pdf.cell(w=30, h=10, txt=columns[4].title(), ln=0, align='L', border=1)

    for index,row in df.iterrows():
        print(row["product_id"])
        pdf.set_line_width(width=.5)
        pdf.set_font(family='arial', style='BI', size=10)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(w=0, h=10, txt='', ln=1, align='L', border=1)
        pdf.cell(w=30, h=10, txt=f"{row['product_id']}", ln=0, align='L', border=1)
        pdf.cell(w=60, h=10, txt=f"{row['product_name']}", ln=0, align='L', border=1)
        pdf.cell(w=40, h=10, txt=f"{row['amount_purchased']}", ln=0, align='L', border=1)
        pdf.cell(w=30, h=10, txt=f"{row['price_per_unit']}", ln=0, align='L', border=1)
        pdf.cell(w=30, h=10, txt=f"{row['total_price']}", ln=0, align='L', border=1)

    total_sum = df["total_price"].sum()

    pdf.set_line_width(width=.5)
    pdf.set_font(family='arial', style='BI', size=10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=10, txt='', ln=1, align='L', border=1)
    pdf.cell(w=30, h=10, txt="", ln=0, align='L', border=1)
    pdf.cell(w=60, h=10, txt="", ln=0, align='L', border=1)
    pdf.cell(w=40, h=10, txt="", ln=0, align='L', border=1)
    pdf.cell(w=30, h=10, txt="", ln=0, align='L', border=1)
    pdf.cell(w=30, h=10, txt=str(total_sum), ln=0, align='L', border=1)


    pdf.ln(30)
    pdf.set_line_width(width=.5)
    pdf.set_font(family='arial', style='BI', size=14)
    pdf.set_text_color(0, 0, 0)
    total_due = f"The total amount due is {total_sum} Euros"
    pdf.cell(w=0, h=10, txt=total_due, ln=1, align='L', border=0)
    pdf.cell(w=30, h=8, txt="PythonHow", ln=0, align='L', border=0)
    pdf.image("images/pythonhow.png", w=10)
    pdf.output(f"PDFs/Invoice_{filename}.pdf")





    #pdf.set_font(family='arial', style='BI', size=30)
    #pdf.set_text_color(0, 0, 0)
    #pdf.cell(w=0, h=15, txt='Invoice', border=0, ln=1, align='C')
    #pdf.set_line_width(width=2)
    #pdf.line(10, 30, 200, 30)

    #pdf.output("Invoice.pdf")
    #for index, row in df.iterrows():
    #    print(row)



#pdf.output("Invoice.pdf")
#pdf = FPDF(orientation='P', unit='mm', format='A4')
#
#pdf.add_page()
#
#pdf.set_font(family='arial',style='BI',size=30)
#pdf.set_text_color(0, 0, 0)
#pdf.cell(w=0, h=15, txt='Invoice',border=0,ln=1,align='C')
#pdf.set_line_width(width=2)
#pdf.line(10, 30, 200, 30)
#

#pdf.ln(70)

#pdf.set_font(family='arial',style='BI',size=10)
#pdf.set_text_color(0, 0, 0)
#pdf.cell(w=20, h=0, txt='S.No', ln=0, align='L')
#
#pdf.set_font(family='arial',style='BI',size=10)
#pdf.set_text_color(0, 0, 0)
#pdf.cell(w=20, h=0, txt='product_id', ln=0, align='L')
#
#pdf.set_font(family='arial',style='BI',size=10)
#pdf.set_text_color(0, 0, 0)
#pdf.cell(w=60, h=0, txt='product_name', ln=0, align='L')
#
#pdf.set_font(family='arial',style='BI',size=10)
#pdf.set_text_color(0, 0, 0)
#pdf.cell(w=40, h=0, txt='amount_purchased', ln=0, align='L')
#
#pdf.set_font(family='arial',style='BI',size=10)
#pdf.set_text_color(0, 0, 0)
#pdf.cell(w=30, h=0, txt='price_per_unit', ln=0, align='L')
#
#pdf.set_font(family='arial',style='BI',size=10)
#pdf.set_text_color(0, 0, 0)
#pdf.cell(w=40, h=0, txt='total_price', ln=0, align='L')
#
#pdf.set_line_width(width=1)
#pdf.line(10, 100, 200, 100)



