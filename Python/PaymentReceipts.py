import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet

# Read data from Excel
excel_file = "data\\sal_data.xlsx"  # Ensure this file is in the same directory
df = pd.read_excel(excel_file)
print(df.columns)

# Convert DataFrame to a list of lists for reportlab Table
DATA = [["Date", "Name", "Description", "Gross (Rs.)"]] + df.values.tolist()
df.columns = df.columns.str.replace(r'[",]', '', regex=True).str.strip()
print(df.columns)  # Check the cleaned column names

# Add subtotal, discount, and total dynamically
subtotal = df["Gross (Rs.)"].sum()
advances = df["advances"].sum()
total = subtotal - advances

DATA.append(["Sub Total", "", "", f"{subtotal:,.2f}/-"])
DATA.append(["Advances", "", "", f"-{advances:,.2f}/-"])
DATA.append(["Total", "", "", f"{total:,.2f}/-"])

# Create PDF document
pdf = SimpleDocTemplate("receipt.pdf", pagesize=A4)
styles = getSampleStyleSheet()
title_style = styles["Heading1"]
title_style.alignment = 1
title = Paragraph("Payment Receipt", title_style)

# Define table style
style = TableStyle([
    ("BOX", (0, 0), (-1, -1), 1, colors.black),
    ("GRID", (0, 0), (-1, -1), 1, colors.black),
    ("BACKGROUND", (0, 0), (-1, 0), colors.gray),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
])

table = Table(DATA, style=style)

# Build the PDF
pdf.build([title, table])

print("Receipt generated successfully!")
