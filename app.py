import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
import zipfile
import tempfile
import re
from docx2pdf import convert

st.set_page_config(page_title="Internship Letter & Report Generator", layout="centered")
st.title("📄 Internship Letter & Report Generator")

# Step 1: Module selection
module = st.selectbox("Choose a Module", [
    "Getin - Intern Acceptance",
    "Getin - Intern Completion Letter",
    "Infonel - Intern Acceptance Letter",
    "Infonel - Intern Completion Letter",
    "Payments Report Merge",
    "Amount Open Merge",
     "Invoice Merge - Amount Open, Amount with Tax, Discount Merge"
])

# Step 2: Upload files
if module == "Payments Report Merge":
    invoice_file = st.file_uploader("Upload 'Invoices Report' Excel", type=["xlsx"], key="invoice_"+module)
    payment_file = st.file_uploader("Upload 'Payments Received' Excel", type=["xlsx"], key="payment_"+module)
elif module == "Amount Open Merge":
    invoice_file = st.file_uploader("Upload 'Invoices' Excel", type=["xlsx"], key="invoice_"+module)
    report_file = st.file_uploader("Upload 'Invoices Report' Excel", type=["xlsx"], key="report_"+module)
elif module == "Invoice Merge - Amount Open, Amount with Tax, Discount Merge":
    invoice_file = st.file_uploader("Upload 'Invoices' Excel", type=["xlsx"], key="invoice_"+module)
    report_file = st.file_uploader("Upload 'Invoices Report' Excel", type=["xlsx"], key="report_"+module)
else:
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"], key="excel_"+module)
    template_file = st.file_uploader("Upload Word Template (DOCX)", type=["docx"], key="template_"+module)

def get_pronouns(gender):
    if isinstance(gender, str):
        gender = gender.lower()
        if gender == "male":
            return {"pronoun_subject": "he", "pronoun_object": "him", "pronoun_possessive": "his"}
        elif gender == "female":
            return {"pronoun_subject": "she", "pronoun_object": "her", "pronoun_possessive": "her"}
    return {"pronoun_subject": "they", "pronoun_object": "them", "pronoun_possessive": "their"}

if st.button("Generate"):
    try:
        if module == "Payments Report Merge":
            if not invoice_file or not payment_file:
                st.warning("Please upload both Invoices Report and Payments Received files.")
            else:
                invoices_df = pd.read_excel(invoice_file, header=1)
                payments_df = pd.read_excel(payment_file, header=1)

                invoices_df.columns = invoices_df.columns.str.strip()
                payments_df.columns = payments_df.columns.str.strip()

                if 'Invoice #' in invoices_df.columns and 'Invoice #' in payments_df.columns:
                    merged_df = payments_df.merge(
                        invoices_df[['Invoice #', 'Branch']],
                        on='Invoice #', how='left')

                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmpfile:
                        merged_df.to_excel(tmpfile.name, index=False)
                        tmpfile.seek(0)
                        with open(tmpfile.name, "rb") as f:
                            st.success("✅ Merge complete!")
                            st.download_button(
                                label="📅 Download Merged Report",
                                data=f,
                                file_name="Payments_Received_With_Branch.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.error("❌ Error: 'Invoice #' column not found in both files.")
                    
        elif module == "Invoice Merge - Amount Open, Amount with Tax, Discount Merge":
            if not invoice_file or not report_file:
                st.warning("Please upload both Invoices and Invoices Report files.")
            else:
                invoices_df = pd.read_excel(invoice_file, header=1)
                invoices_report_df = pd.read_excel(report_file, header=1)

                invoices_df.columns = invoices_df.columns.str.strip().str.lower()
                invoices_report_df.columns = invoices_report_df.columns.str.strip().str.lower()

                if 'amount' in invoices_df.columns:
                    invoices_df.rename(columns={'amount': 'amount with tax'}, inplace=True)

                required_columns = ['invoice #', 'amount', 'discount', 'adjustment', 'amount open']
                missing = [col for col in required_columns if col not in invoices_report_df.columns]
                if missing:
                    st.error(f"❌ Error: Missing columns in Invoices Report: {missing}")
                else:
                    merge_df = invoices_report_df[required_columns]
                    merged_df = invoices_df.merge(merge_df, on='invoice #', how='left')

                    columns_to_remove = ['total tax', 'year', 'project', 'tags']
                    merged_df.drop(columns=[col for col in columns_to_remove if col in merged_df.columns], inplace=True)

                    merged_df.columns = [col.title() for col in merged_df.columns]

                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmpfile:
                        merged_df.to_excel(tmpfile.name, index=False)
                        tmpfile.seek(0)
                        with open(tmpfile.name, "rb") as f:
                            st.success("✅ Merged successfully!")
                            st.download_button(
                                label="📥 Download Merged Excel",
                                data=f,
                                file_name="Invoices_With_All_Merged.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
        elif module == "Amount Open Merge":
            if not invoice_file or not report_file:
                st.warning("Please upload both Invoices and Invoices Report files.")
            else:
                invoices_df = pd.read_excel(invoice_file, header=1)
                report_df = pd.read_excel(report_file, header=1)

                invoices_df.columns = invoices_df.columns.str.strip().str.lower()
                report_df.columns = report_df.columns.str.strip().str.lower()

                if 'invoice #' in invoices_df.columns and 'invoice #' in report_df.columns and 'amount open' in report_df.columns:
                    merged_df = invoices_df.merge(
                        report_df[['invoice #', 'amount open']],
                        on='invoice #', how='left')
                    
                    columns_to_remove = ['total tax', 'year', 'project', 'tags']
                    merged_df.drop(columns=[col for col in columns_to_remove if col in merged_df.columns], inplace=True)
                    
                    merged_df.columns = [col.title() for col in merged_df.columns]

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmpfile:
                        merged_df.to_excel(tmpfile.name, index=False)
                        tmpfile.seek(0)
                        with open(tmpfile.name, "rb") as f:
                            st.success("✅ Merge complete!")
                            st.download_button(
                                label="📅 Download Merged Report",
                                data=f,
                                file_name="Invoices_With_Amount_Open.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.error("❌ Error: Required columns ('invoice #' and 'amount open') not found in both files.")

        else:
            if not excel_file or not template_file:
                st.warning("Please upload both Excel and DOCX template.")
            else:
                df = pd.read_excel(excel_file)
                today_date = datetime.today().strftime("%d/%m/%Y")

                with tempfile.TemporaryDirectory() as tmpdir:
                    zip_path = os.path.join(tmpdir, f"{module.replace(' ', '_')}_certificates.zip")
                    with zipfile.ZipFile(zip_path, "w") as zipf:

                        if module == "Getin - Intern Completion Letter":
                            df['Start Date'] = pd.to_datetime(df['Start Date'], format='%d %B %Y')
                            df['End Date'] = pd.to_datetime(df['End Date'], format='%d %B %Y')
                            for _, row in df.iterrows():
                                doc = DocxTemplate(template_file)
                                pronouns = get_pronouns(row.get('Gender', ''))
                                context = {
                                    'date': today_date,
                                    'name': row['Name'].title(),
                                    'roll_no': row['Roll No'],
                                    'college': row['College Name'],
                                    'position': row['Position'],
                                    'start_date': row['Start Date'].strftime("%d %B %Y"),
                                    'end_date': row['End Date'].strftime("%d %B %Y"),
                                    'pronoun_subject': pronouns['pronoun_subject'],
                                    'pronoun_object': pronouns['pronoun_object'],
                                    'pronoun_possessive': pronouns['pronoun_possessive'],
                                }
                                doc.render(context)
                                filename = f"{row['Name'].replace(' ', '_')}_Completion_Certificate.docx"
                                filepath = os.path.join(tmpdir, filename)
                                doc.save(filepath)

                                # Convert to PDF
                                pdf_filepath = filepath.replace('.docx', '.pdf')
                                convert(filepath, pdf_filepath)
                                
                                zipf.write(filepath, arcname=filename)
                                zipf.write(pdf_filepath, arcname=filename.replace('.docx', '.pdf'))

                        elif module == "Getin - Intern Acceptance":
                            df['Start Date'] = pd.to_datetime(df['Start Date'], format='%d %B %Y')
                            df['End Date'] = pd.to_datetime(df['End Date'], format='%d %B %Y')
                            for _, row in df.iterrows():
                                doc = DocxTemplate(template_file)
                                context = {
                                    'date': today_date,
                                    'name': row['Name'].title(),
                                    'roll_no': row['Roll No'],
                                    'college': row['College Name'],
                                    'city': row['City'].title(),
                                    'postal_code': row['Postal Code'],
                                    'position': row['Position'],
                                    'field': row['Field'],
                                    'location': row['Location'].title(),
                                    'start_date': row['Start Date'].strftime("%d %B %Y"),
                                    'end_date': row['End Date'].strftime("%d %B %Y"),
                                }
                                doc.render(context)
                                filename = f"{row['Name'].replace(' ', '_')}_Internship_Letter.docx"
                                filepath = os.path.join(tmpdir, filename)
                                doc.save(filepath)

                                # Convert to PDF
                                pdf_filepath = filepath.replace('.docx', '.pdf')
                                convert(filepath, pdf_filepath)

                                zipf.write(filepath, arcname=filename)
                                zipf.write(pdf_filepath, arcname=filename.replace('.docx', '.pdf'))

                        elif module == "Infonel - Intern Acceptance Letter":
                            df['Start Date'] = pd.to_datetime(df['Start Date'])
                            df['End Date'] = pd.to_datetime(df['End Date'])
                            for _, row in df.iterrows():
                                doc = DocxTemplate(template_file)
                                context = {
                                    'date': today_date,
                                    'name': row['Name'].title(),
                                    'roll_no': row['Roll No'],
                                    'position': row['Position'],
                                    'start_date': row['Start Date'].strftime("%d %B %Y"),
                                    'end_date': row['End Date'].strftime("%d %B %Y"),
                                    'field': row['Field'],
                                    'location': row['Location'],
                                    'city': row['City'],
                                }
                                doc.render(context)
                                filename = f"{row['Name'].replace(' ', '_')}_Infonel_Internship_Acceptance_Letter.docx"
                                filepath = os.path.join(tmpdir, filename)
                                doc.save(filepath)

                                # Convert to PDF
                                pdf_filepath = filepath.replace('.docx', '.pdf')
                                convert(filepath, pdf_filepath)

                                zipf.write(filepath, arcname=filename)
                                zipf.write(pdf_filepath, arcname=filename.replace('.docx', '.pdf'))

                        elif module == "Infonel - Intern Completion Letter":
                            df['Start Date'] = pd.to_datetime(df['Start Date'], format='%d %B %Y')
                            df['End Date'] = pd.to_datetime(df['End Date'], format='%d %B %Y')
                            for _, row in df.iterrows():
                                doc = DocxTemplate(template_file)
                                context = {
                                    'date': today_date,
                                    'name': row['Name'].title(),
                                    'roll_no': row['Roll No'],
                                    'position': row['Position'],
                                    'start_date': row['Start Date'].strftime("%d %B %Y"),
                                    'end_date': row['End Date'].strftime("%d %B %Y"),
                                }
                                doc.render(context)
                                filename = f"{row['Name'].replace(' ', '_')}_Completion_Certificate.docx"
                                filepath = os.path.join(tmpdir, filename)
                                doc.save(filepath)

                                # Convert to PDF
                                pdf_filepath = filepath.replace('.docx', '.pdf')
                                convert(filepath, pdf_filepath)

                                zipf.write(filepath, arcname=filename)
                                zipf.write(pdf_filepath, arcname=filename.replace('.docx', '.pdf'))

                    # Download the zip file
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label="📥 Download ZIP with DOCX & PDF",
                            data=f,
                            file_name=f"{module.replace(' ', '_')}_Certificates.zip",
                            mime="application/zip"
                        )
    except Exception as e:
        st.error(f"Error: {e}")
