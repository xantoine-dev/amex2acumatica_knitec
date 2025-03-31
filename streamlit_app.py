import streamlit as st
import pandas as pd
import os
import tempfile
from pathlib import Path
import re

st.set_page_config(page_title="Amex Expense Tool", layout="centered")
st.title("💳 AMEX Expense Claim Generator")

# File upload
uploaded_statement = st.file_uploader("Upload AMEX Excel or CSV File:", type=["xlsx", "xls", "csv"])

amount_column = "Transaction \nAmount \nUSD"
last_name_column = "Supplemental \nCardmember Last \nName"

# Template
uploaded_template = st.file_uploader("Upload Template File (Excel):", type=["xlsx"])

# Optional corporate card file
uploaded_corp_card = st.file_uploader("Upload Employee Corporate Card File (Optional):", type=["xlsx"])

export_format = st.selectbox("Choose export format:", ["excel", "csv"])

if uploaded_statement and uploaded_template:
    try:
        # Read statement
        if uploaded_statement.name.endswith("csv"):
            df = pd.read_csv(uploaded_statement)
        else:
            df = pd.read_excel(uploaded_statement)

        df.columns = df.iloc[13]  # Use row 14 as header
        df = df.iloc[14:].reset_index(drop=True)

        # Clean amount column
        if amount_column in df.columns:
            df[amount_column] = df[amount_column].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False)
            df[amount_column] = pd.to_numeric(df[amount_column], errors='coerce')
            df = df[df[amount_column] >= 0]

        # Remove numbers from Description 4
        df['Transaction \nDescription 4'] = df['Transaction \nDescription 4'].apply(
            lambda x: re.sub(r'\d+', '', x) if isinstance(x, str) else x
        )

        # Load template
        template_df = pd.read_excel(uploaded_template)
        template_columns = template_df.columns.tolist()

        unique_names = df[last_name_column].dropna().unique()

        zip_buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
        output_files = []

        for name in unique_names:
            emp_data = df[df[last_name_column] == name]
            emp_template = pd.DataFrame(columns=template_columns)

            for _, row in emp_data.iterrows():
                new_row = {col: "" for col in emp_template.columns}
                new_row["Date"] = row.get("Transaction Date", "")
                new_row["Ref. Nbr."] = row.get("Transaction \nDescription 4", "")
                new_row["Description"] = row.get("Transaction \nDescription 1", "")
                new_row["Amount"] = row.get(amount_column, "")
                new_row["Claim Amount"] = row.get(amount_column, "")
                new_row["Paid With"] = "Corporate Card, Company Expense"
                new_row["Branch"] = "KEC"
                emp_template = pd.concat([emp_template, pd.DataFrame([new_row])], ignore_index=True)

            out_filename = f"{name}_AMEX_Claim.{ 'xlsx' if export_format == 'excel' else 'csv' }"
            out_path = os.path.join(tempfile.gettempdir(), out_filename)

            if export_format == "excel":
                emp_template.to_excel(out_path, index=False)
            else:
                emp_template.to_csv(out_path, index=False)

            with open(out_path, "rb") as f:
                st.download_button(f"Download {name}'s File", f, file_name=out_filename)

        st.success("✅ All files processed and ready for download.")

    except Exception as e:
        st.error(f"Error processing files: {e}")

else:
    st.info("Please upload both the AMEX statement and the template file.")
