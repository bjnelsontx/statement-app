
import streamlit as st
import pandas as pd
import time
import io
from zipfile import ZipFile
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from datetime import datetime

st.set_page_config(page_title="Hi-Line Statement Generator", layout="centered")
st.title("📄 Hi-Line Statement Generator")

uploaded_excel = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_excel is not None:
    logo_path = "HI-LINE logo DK Red.jpg"
    try:
        logo_image = ImageReader(logo_path)
    except Exception:
        st.error("❌ Could not find or load 'logo.jpg'. Make sure it's in your Streamlit repo.")
        st.stop()

    try:
        df = pd.read_excel(uploaded_excel)
        as_of_col = next(col for col in df.columns if col.strip().lower() == "as of date")
        df[as_of_col] = pd.to_datetime(df[as_of_col], errors='coerce')

        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, "w") as zipf:
            grouped = df.groupby("customer_id")
            progress_bar = st.progress(0)
    total_customers = len(grouped)
    for i, (customer_id, group) in enumerate(grouped):
                group = group.reset_index(drop=True)
                customer_name = group.loc[0, "bill2_name"].replace(" ", "_").replace("/", "_")
                pdf_buffer = io.BytesIO()
                c = canvas.Canvas(pdf_buffer, pagesize=LETTER)

                margin = 0.5 * inch
                width, height = LETTER
                logo_width = 2.0 * inch
                logo_height = logo_width * 500 / 2048
                rows_per_page = 30
                x_positions = [margin + i * inch for i in range(8)]
                label_x = width - 3.0 * inch
                amount_x = width - 1.75 * inch
                headers = ["Invoice #", "Invoice Date", "Due Date", "PO #", "Contract #", "Charges", "Credits", "Amount Due"]

                total_pages = (len(group) + rows_per_page - 1) // rows_per_page
                as_of_date = pd.to_datetime(group.loc[0, as_of_col]).strftime('%m/%d/%Y')
                city_zip = f"{group.loc[0, 'bill2_city']}, {group.loc[0, 'bill2_state']} {group.loc[0, 'bill2_postal_code']}"
                total_due = group.loc[0, "TOTAL_ACT_DUE"]

                for page_num in range(total_pages):
                    start = page_num * rows_per_page
                    end = start + rows_per_page
                    subset = group.iloc[start:end]

                    c.drawImage(logo_image, margin, height - margin - logo_height, width=logo_width, height=logo_height, mask='auto')
                    c.setFont("Helvetica", 9)
                    for j, line in enumerate(["HI-LINE, INC", "Remit to:", "PO BOX 972081", "Dallas, TX  75397-2081"]):
                        c.drawString(margin + logo_width + 0.2 * inch, height - margin - 12 - j * 10, line)
                    for j, line in enumerate(["Other Inquiries:", "2121 Valley View Lane", "Dallas, TX 75234", "United States of America"]):
                        c.drawString(margin + logo_width + 2.0 * inch, height - margin - 12 - j * 10, line)

                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(width - margin - c.stringWidth("STATEMENT", "Helvetica-Bold", 14), height - margin - 10, "STATEMENT")

                    info = [["DATE", as_of_date], ["Customer ID", str(customer_id)], ["As of Date", as_of_date], ["Page", f"{page_num + 1} of {total_pages}"]]
                    table = Table(info, colWidths=[0.9*inch, 0.95*inch])
                    table.setStyle(TableStyle([
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                        ('FONT', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 7),
                        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
                    ]))
                    table.wrapOn(c, width, height)
                    table.drawOn(c, width - 2.25*inch, height - margin - 1.2 * inch)

                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(width - 2.25*inch, height - margin - 1.45 * inch, "AMOUNT DUE")
                    c.setFont("Helvetica", 10)
                    c.drawString(width - 2.25*inch, height - margin - 1.58 * inch, f"${total_due:.2f}")

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

                    y = addr_y - 30
                    c.setFont("Helvetica-Bold", 9)
                    for k, header in enumerate(headers):
                        c.drawString(x_positions[k], y, header)
                    y -= 14
                    c.setFont("Helvetica", 9)

                    for _, row in subset.iterrows():
                        po_no = "" if pd.isna(row['po_no']) else str(row['po_no'])
                        row_data = [
                            str(row['invoice_no']),
                            pd.to_datetime(row['invoice_date']).strftime('%m/%d/%Y'),
                            pd.to_datetime(row['net_due_date']).strftime('%m/%d/%Y'),
                            po_no,
                            str(row['Contract#']) if pd.notna(row['Contract#']) else "",
                            f"${row['total_amount']:.2f}",
                            f"${row['amount_paid']:.2f}",
                            f"${row['Amt_due']:.2f}"
                        ]
                        for j, val in enumerate(row_data):
                            c.drawString(x_positions[j], y, val)
                        y -= 14

                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(label_x, 0.8 * inch, "Total Amount Due:")
                    c.drawString(amount_x, 0.8 * inch, f"${total_due:.2f}")
                    c.drawString(label_x, 0.6 * inch, "Amount Enclosed:")
                    c.line(amount_x - 0.05 * inch, 0.58 * inch, amount_x + 1.2 * inch, 0.58 * inch)

                    c.showPage()

                c.save()
        progress_bar.progress((i + 1) / total_customers)
        time.sleep(0.1)
                zipf.writestr(f"{customer_name}_{customer_id}.pdf", pdf_buffer.getvalue())

        zip_buffer.seek(0)
        st.success("✅ PDF generation complete!")
        st.download_button("📥 Download Statements ZIP", data=zip_buffer.getvalue(), file_name="Customer_Statements.zip")

    except Exception as e:
        st.error(f"Error: {e}")
