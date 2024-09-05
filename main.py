import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Get all Excel file paths from the 'invoices' directory
filepaths = glob.glob("invoices/*xlsx")

# Loop through each invoice file
for filepath in filepaths:
    # Initialize a PDF object with A4 page size and portrait orientation
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extract invoice number and date from the filename
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Add Invoice header
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Read the Excel file into a DataFrame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header for the table in the PDF
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)

    # Create table headers for the PDF
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows from the DataFrame to the table in the PDF
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Calculate and display the total sum
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)

    # Add empty cells for alignment
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add a sentence summarizing the total sum
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

    # Add company name and logo at the end of the PDF
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=40, h=8, txt="WorldClass Medic")
    pdf.image("logo.png", w=10)

    # Output the PDF file to the 'PDFs' directory
    pdf.output(f"PDFs/{filename}.pdf")
