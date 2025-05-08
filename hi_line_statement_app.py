import streamlit as st
import pandas as pd
import io
from zipfile import ZipFile
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from datetime import datetime
import time
import base64

# Custom styling and configuration
def set_page_styling():
    # Page configuration
    st.set_page_config(
        page_title="Hi-Line Statement Generator",
        page_icon="📄",
        layout="centered",
        initial_sidebar_state="collapsed"
    )
    
    # Custom CSS
    st.markdown("""
    <style>
        /* Main page styling */
        .main .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
            max-width: 800px;
        }
        
        /* Hi-Line branding colors */
        :root {
            --hiline-red: #8B0000;
            --hiline-gray: #f0f2f6;
        }
        
        /* Primary buttons */
        .stButton>button {
            background-color: var(--hiline-red);
            color: white;
            border: none;
            border-radius: 4px;
            padding: 0.5rem 1rem;
            font-weight: 500;
        }
        
        .stButton>button:hover {
            background-color: darkred;
            color: white;
        }
        
        /* Title styling */
        h1 {
            color: var(--hiline-red) !important;
            font-weight: 700 !important;
        }
        
        /* Success message styling */
        .stSuccess {
            border-left-color: var(--hiline-red) !important;
        }
        
        /* File uploader styling */
        .stFileUploader > div > div {
            border-color: var(--hiline-red) !important;
            border-width: 2px !important;
        }
        
        /* Card-like containers */
        .card {
            background-color: white;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        /* Progress bar styling */
        .stProgress > div > div > div > div {
            background-color: var(--hiline-red);
        }
    </style>
    """, unsafe_allow_html=True)

# Function to add a Hi-Line logo to the app
def add_logo():
    # Logo path
    logo_path = "HI-LINE logo DK Red.jpg"
    # Logo fallback option - embedded base64 logo (you would replace this with actual logo data)
    logo_base64 = "BASE64_LOGO_DATA_HERE"  # Replace with actual base64 encoded logo if needed
    
    try:
        # Try to use the local logo file
        with open(logo_path, "rb") as f:
            logo_data = f.read()
            logo_base64 = base64.b64encode(logo_data).decode()
    except Exception:
        pass  # Use the fallback logo_base64 if file not found
    
    # Display logo at the top
    st.markdown(f"""
    <div style="text-align: center; margin-bottom: 20px;">
        <img src="data:image/jpeg;base64,{logo_base64}" alt="Hi-Line Logo" width="200">
    </div>
    """, unsafe_allow_html=True)

# Function to create a styled container
def styled_container(title, content_function):
    st.markdown(f"""
    <div class="card">
        <h3 style="color: #8B0000; margin-bottom: 15px;">{title}</h3>
    </div>
    """, unsafe_allow_html=True)
    content_function()

# Function to display app info
def show_app_info():
    with st.expander("ℹ️ About this tool", expanded=False):
        st.markdown("""
        This tool generates customer statements in PDF format from your Excel data.
        
        **Instructions:**
        1. Upload your Excel file with customer data
        2. Wait for the processing to complete
        3. Download the ZIP file containing all customer statements
        
        **Required Excel columns:**
        - customer_id
        - bill2_name
        - invoice_no
        - invoice_date
        - and other billing information
        """)

# Function for file processing
def process_excel_file(uploaded_excel):
    # Status indicators
    status_container = st.container()
    
    with status_container:
        status = st.empty()
        status.info("🔍 Analyzing your Excel file...")
        
        logo_path = "HI-LINE logo DK Red.jpg"
        try:
            logo_image = ImageReader(logo_path)
        except Exception:
            status.error("❌ Could not find or load 'HI-LINE logo DK Red.jpg'. Make sure it's in your GitHub repo.")
            st.stop()

        try:
            df = pd.read_excel(uploaded_excel)
            as_of_col = next(col for col in df.columns if col.strip().lower() == "as of date")
            df[as_of_col] = pd.to_datetime(df[as_of_col], errors='coerce')
            status.success(f"✅ Found {df['customer_id'].nunique()} customers in your file")
        except Exception as e:
            status.error(f"Error loading Excel file: {e}")
            st.stop()

        st.markdown("### Processing Progress")
        progress_text = st.empty()
        progress_bar = st.progress(0)
        
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, "w") as zipf:
            grouped = df.groupby("customer_id")
            total_customers = len(grouped)
            
            for i, (customer_id, group) in enumerate(grouped):
                progress_text.text(f"Processing customer {i+1} of {total_customers}: {customer_id}")
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
                zipf.writestr(f"{customer_name}_{customer_id}.pdf", pdf_buffer.getvalue())
                progress_bar.progress((i + 1) / total_customers)
                time.sleep(0.1)

        progress_text.empty()
        status.success("✅ PDF generation complete!")
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.download_button(
                "📥 Download All Statements (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Customer_Statements.zip",
                mime="application/zip",
            )
        with col2:
            # Add a count of PDFs generated
            st.metric("PDFs Generated", f"{total_customers}")

# Main app function
def main():
    set_page_styling()
    add_logo()
    
    st.title("Hi-Line Statement Generator")
    st.markdown("Generate professional PDF statements from your customer data")
    
    show_app_info()
    
    # Create a nice upload section
    st.markdown("### Upload Your Data")
    st.markdown("Please upload your Excel file with customer information:")
    
    uploaded_excel = st.file_uploader(
        "Drop Excel file here",
        type=["xlsx"],
        help="Upload an Excel file containing customer billing data"
    )
    
    if uploaded_excel is not None:
        process_excel_file(uploaded_excel)
    
    # Add footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray; font-size: 12px;'>© 2025 Hi-Line Inc. All rights reserved.</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()