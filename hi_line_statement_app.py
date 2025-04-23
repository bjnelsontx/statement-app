
import streamlit as st
import pandas as pd
import io
import os
from zipfile import ZipFile
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib.pagesizes import LETTER
from reportlab.lib import colors
from reportlab.lib.units import inch
from datetime import datetime

st.set_page_config(page_title="Hi-Line Statement Generator", layout="centered")
st.title("📄 Hi-Line Customer Statement Generator")

uploaded_file = st.file_uploader("Upload your Excel file with statement data", type=["xlsx"])
logo = st.file_uploader("Upload the Hi-Line logo (JPG only)", type=["jpg"])

if uploaded_file and logo:
    with st.spinner("Generating customer statements..."):
        df = pd.ExcelFile(uploaded_file).parse('5 Data Only')
        grouped = df.groupby("customer_id")

        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, 'w') as zipf:
            for customer_id, group in grouped:
                group = group.reset_index(drop=True)
                customer_name = group.loc[0, "bill2_name"].replace(" ", "_").replace("/", "_")
                file_name = f"{customer_name}_{customer_id}.pdf"
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=LETTER)

                margin = 0.5 * inch
                width, height = LETTER
                logo_width = 2.0 * inch
                logo_height = logo_width * 500 / 1536
                rows_per_page = 30
                x_positions = [margin + i * inch for i in range(8)]
                label_x = width - 3.0 * inch
                amount_x = width - 1.75 * inch
                headers = ["Invoice #", "Invoice Date", "Due Date", "PO #", "Contract #", "Charges", "Credits", "Amount Due"]

                total_pages = (len(group) + rows_per_page - 1) // rows_per_page
                as_of_date = pd.to_datetime(group.loc[0, "AS of Date"]).strftime('%m/%d/%Y')
                city_zip = f"{group.loc[0, 'bill2_city']}, {group.loc[0, 'bill2_state']} {group.loc[0, 'bill2_postal_code']}"
                total_due = group.loc[0, "TOTAL_ACT_DUE"]

                for page_num in range(total_pages):
                    start = page_num * rows_per_page
                    end = start + rows_per_page
                    subset = group.iloc[start:end]

                    logo_path = logo
                    c.drawImage(logo_path, margin, height - margin - logo_height, width=logo_width, height=logo_height, mask='auto')

                    # Remit and inquiries
                    remit_y = height - margin - 12
                    for i, line in enumerate(["HI-LINE, INC", "Remit to:", "PO BOX 972081", "Dallas, TX  75397-2081"]):
                        c.drawString(margin + logo_width + 0.2 * inch, remit_y - i * 10, line)
                    for i, line in enumerate(["Other Inquiries:", "2121 Valley View Lane", "Dallas, TX 75234", "United States of America"]):
                        c.drawString(margin + logo_width + 2.0 * inch, remit_y - i * 10, line)

                    # Top right: STATEMENT and info box
                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(width - margin - c.stringWidth("STATEMENT", "Helvetica-Bold", 14), height - margin - 10, "STATEMENT")

                    info = [["DATE", as_of_date], ["Customer ID", str(customer_id)], ["As of Date", as_of_date], ["Page", f"{page_num + 1} of {total_pages}"]]
                    table = Table(info, colWidths=[0.95*inch, 1*inch])
                    table.setStyle(TableStyle([
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                        ('FONT', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 7.5),
                        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
                    ]))
                    table.wrapOn(c, width, height)
                    table.drawOn(c, width - 2.25*inch, height - margin - 1.2 * inch)

                    # Amount Due section
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(width - 2.25*inch, height - margin - 1.45 * inch, "AMOUNT DUE")
                    c.setFont("Helvetica", 10)
                    c.drawString(width - 2.25*inch, height - margin - 1.58 * inch, f"${total_due:.2f}")

                    # Customer address (0.5" from left)
                    addr_x = 0.5 * inch
                    addr_y = height - margin - logo_height - 1.25 * inch
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(addr_x, addr_y, group.loc[0, 'bill2_name'])
                    c.setFont("Helvetica", 10)
                    addr_y -= 12
                    c.drawString(addr_x, addr_y, group.loc[0, 'bill2_address1'])
                    addr_y -= 12
                    if pd.notna(group.loc[0, 'bill2_address2']):
                        c.drawString(addr_x, addr_y, group.loc[0, 'bill2_address2'])
                        addr_y -= 12
                    c.drawString(addr_x, addr_y, city_zip)

                    # Table headers and invoice rows
                    y = addr_y - 30
                    c.setFont("Helvetica-Bold", 9)
                    for i, header in enumerate(headers):
                        c.drawString(x_positions[i], y, header)
                    y -= 14
                    c.setFont("Helvetica", 9)

                    for _, row in subset.iterrows():
                        po_no = "" if pd.isna(row['po_no']) else str(row['po_no'])
                        row_data = [
                            str(row['invoice_no']),
                            pd.to_datetime(row['invoice_date']).strftime('%m/%d/%Y'),
                            pd.to_datetime(row['net_due_date']).strftime('%m/%d/%Y'),
                            po_no,
                            str(row['Contract #']) if pd.notna(row['Contract #']) else "",
                            f"${row['total_amount']:.2f}",
                            f"${row['amount_paid']:.2f}",
                            f"${row['Amt_due']:.2f}"
                        ]
                        for j, val in enumerate(row_data):
                            c.drawString(x_positions[j], y, val)
                        y -= 14

                    # Footer
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(label_x, 0.8 * inch, "Total Amount Due:")
                    c.drawString(amount_x, 0.8 * inch, f"${total_due:.2f}")
                    c.drawString(label_x, 0.6 * inch, "Amount Enclosed:")
                    c.line(amount_x - 0.05 * inch, 0.58 * inch, amount_x + 1.2 * inch, 0.58 * inch)

                    c.showPage()

                c.save()
                zipf.writestr(file_name, buffer.getvalue())

        zip_buffer.seek(0)
        st.success("✅ All statements generated!")
        st.download_button("📦 Download ZIP of Statements", zip_buffer, file_name="Customer_Statements.zip")
