import os
import pandas as pd
from fpdf import FPDF, XPos, YPos

if not os.path.exists("./PDFs"):
    os.mkdir("./PDFs")

for i in os.listdir("./Invoices"):
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    filepath = os.path.join("./Invoices", i)

    data = pd.read_excel(filepath)
    total_price = str(data["total_price"].sum())

    pdf.add_page()

    id_number, date = i[:-5].split("-", 1)

    pdf.set_font(family="Times", style="B", size=20)
    pdf.cell(
        w=0, h=12, txt=f"Invoice nr. {id_number}", new_x=XPos.LMARGIN, new_y=YPos.NEXT
    )

    pdf.set_font(family="Times", style="B", size=22)
    pdf.cell(w=0, h=16, txt=f"Date {date}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font(family="Times", style="B", size=10)
    line_height = pdf.font_size_pt
    col_widths = [23, 50, 39, 39, 39]

    for col, width in zip(data.columns, col_widths):
        formated_col = col.replace("_", " ")
        formated_col = formated_col.title()
        pdf.multi_cell(
            w=width,
            h=line_height,
            txt=formated_col,
            border=1,
            new_x="RIGHT",
            new_y="TOP",
        )

    pdf.set_font(family="Times", size=10)
    for index, row in data.iterrows():
        pdf.ln(line_height)
        for value, width in zip(row.values, col_widths):
            pdf.multi_cell(
            w=width,
            h=line_height,
            txt=str(value),
            border=1,
            new_x="RIGHT",
            new_y="TOP",
        )
    pdf.ln(line_height)
    for value, width in zip(["","","","",total_price], col_widths):
            pdf.multi_cell(
            w=width,
            h=line_height,
            txt=str(value),
            border=1,
            new_x="RIGHT",
            new_y="TOP",
        )

    pdf.ln(line_height*2)
    
    content = f"The total due amount is {total_price} Euros."
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=0, h=16, txt=content, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.output(f"./PDFs/{os.path.basename(filepath)}.pdf")
