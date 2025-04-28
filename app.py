import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
import zipfile
import tempfile
import re

st.set_page_config(page_title="Internship Letter & Report Generator", layout="centered")
st.title("📄 Internship Letter & Report Generator")

# Step 1: Module selection
module = st.selectbox("Choose a Module", [
    "Getin - Intern Acceptance",
    "Getin - Intern Completion Letter",
    "Infonel - Intern Acceptance Letter",
    "Infonel - Intern Completion Letter",
    "Payments Report Merge"
])

# Step 2: Upload files
if module == "Payments Report Merge":
    invoice_file = st.file_uploader("Upload 'Invoices Report' Excel", type=["xlsx"], key="invoice_"+module)
    payment_file = st.file_uploader("Upload 'Payments Received' Excel", type=["xlsx"], key="payment_"+module)
else:
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"], key="excel_"+module)
    template_file = st.file_uploader("Upload Word Template (DOCX)", type=["docx"], key="template_"+module)

def get_pronouns(gender):
    if isinstance(gender, str):
        gender = gender.lower()
        if gender == "male":
            return {
                "pronoun_subject": "he",
                "pronoun_object": "him",
                "pronoun_possessive": "his",
            }
        elif gender == "female":
            return {
                "pronoun_subject": "she",
                "pronoun_object": "her",
                "pronoun_possessive": "her",
            }
    # Default (for non-binary, unknown, or missing)
    return {
        "pronoun_subject": "they",
        "pronoun_object": "them",
        "pronoun_possessive": "their",
    }

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
                        on='Invoice #',
                        how='left'
                    )

                    # Create a temporary file to save the merged report
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmpfile:
                        merged_df.to_excel(tmpfile.name, index=False)
                        tmpfile.seek(0)
                        # Open the file and send its binary data to the download button
                        with open(tmpfile.name, "rb") as f:
                            st.success("✅ Merge complete!")
                            st.download_button(
                                label="📥 Download Merged Report",
                                data=f,
                                file_name="Payments_Received_With_Branch.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                else:
                    st.error("❌ Error: 'Invoice #' column not found in both files.")

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
                                zipf.write(filepath, arcname=filename)

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
                                zipf.write(filepath, arcname=filename)

                        elif module == "Infonel - Intern Acceptance Letter":
                            df['Start Date'] = pd.to_datetime(df['Start Date'])
                            df['End Date'] = pd.to_datetime(df['End Date'])
                            base_id = 600
                            for _, row in df.iterrows():
                                doc = DocxTemplate(template_file)
                                clean_name = re.sub(r'[^a-zA-Z\s]', '', row['Name']).title()
                                context = {
                                    'date': today_date,
                                    'certificate_id': f"INT/VNR{base_id:03d}",
                                    'name': clean_name,
                                    'roll_no': row['Roll No'],
                                    'college_name': row['College Name'],
                                    'college_location': row['College Location'],
                                    'college_pincode': row['College Pincode'],
                                    'position': row['Position'].title(),
                                    'start_date': row['Start Date'].strftime("%d/%m/%Y"),
                                    'end_date': row['End Date'].strftime("%d/%m/%Y"),
                                }
                                doc.render(context)
                                filename = f"{clean_name.replace(' ', '_')}_Internship_Confirmation_{base_id}.docx"
                                filepath = os.path.join(tmpdir, filename)
                                doc.save(filepath)
                                zipf.write(filepath, arcname=filename)
                                base_id += 1

                        elif module == "Infonel - Intern Completion Letter":
                            df['Start Date'] = pd.to_datetime(df['Start Date'], format='%d %B %Y')
                            df['End Date'] = pd.to_datetime(df['End Date'], format='%d %B %Y')
                            start_id = 501
                            for i, row in df.iterrows():
                                doc = DocxTemplate(template_file)
                                certificate_id = f"INT/KVP{start_id + i:03d}"
                                context = {
                                    'date': today_date,
                                    'certificate_id': certificate_id,
                                    'name': row['Name'].title(),
                                    'roll_no': row['Roll No'],
                                    'college': row['College Name'],
                                    'position': row['Position'],
                                    'start_date': row['Start Date'].strftime('%d %B %Y'),
                                    'end_date': row['End Date'].strftime('%d %B %Y'),
                                    'work_description': row['Work Description'],
                                }
                                doc.render(context)
                                safe_id = certificate_id.replace("/", "_")
                                filename = f"{row['Name'].replace(' ', '_')}_{safe_id}.docx"
                                filepath = os.path.join(tmpdir, filename)
                                doc.save(filepath)
                                zipf.write(filepath, arcname=filename)

                    with open(zip_path, "rb") as f:
                        st.success("✅ Letters generated successfully!")
                        st.download_button("📥 Download All Letters (ZIP)", data=f, file_name=f"{module.replace(' ', '_')}_Letters.zip")

    except Exception as e:
        st.error(f"❌ Error: {e}")
