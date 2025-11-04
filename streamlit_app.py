import io
from zipfile import ZipFile

import pandas as pd
import streamlit as st

from amex_tool.pipeline import (
    apply_corporate_cards,
    clean_statement,
    generate_claim_frames,
    load_corporate_mapping,
    load_statement,
    load_template_columns,
    GROUP_COLUMN,
)

st.set_page_config(page_title="AMEX Expense Claim Generator", layout="wide")
st.title("💳 AMEX Expense Claim Generator")

st.markdown(
    "Upload the monthly AMEX statement along with an optional template and corporate "
    "card mapping file. The app will generate one claim workbook per cardholder."
)

with st.form("processor"):
    uploaded_statement = st.file_uploader(
        "AMEX Statement (.csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"]
    )
    uploaded_template = st.file_uploader(
        "Output Template (optional, .csv, .xls, .xlsx)", type=["csv", "xls", "xlsx"]
    )
    uploaded_corporate = st.file_uploader(
        "Corporate Card Mapping (optional, .csv, .xls, .xlsx)",
        type=["csv", "xls", "xlsx"],
    )
    export_format = st.selectbox("Export format", ["excel", "csv"])
    submitted = st.form_submit_button("Generate Claim Files")

if submitted:
    if not uploaded_statement:
        st.error("Please upload an AMEX statement file.")
    else:
        try:
            statement_df = load_statement(uploaded_statement, uploaded_statement.name)
            cleaned_df = clean_statement(statement_df)

            template_columns = (
                load_template_columns(uploaded_template, uploaded_template.name)
                if uploaded_template
                else None
            )

            claims = generate_claim_frames(cleaned_df, template_columns)

            if not claims:
                st.warning(
                    f"No cardholder rows found. Ensure '{GROUP_COLUMN}' is populated."
                )
            else:
                mapping = (
                    load_corporate_mapping(
                        uploaded_corporate, uploaded_corporate.name
                    )
                    if uploaded_corporate
                    else None
                )
                claims = apply_corporate_cards(claims, mapping)

                summary_rows = []
                for last_name, frame in claims.items():
                    totals = pd.to_numeric(
                        frame.get("Amount", pd.Series(dtype=float)), errors="coerce"
                    ).fillna(0)
                    summary_rows.append(
                        {
                            "Last Name": last_name,
                            "Rows": len(frame),
                            "Total Amount": float(totals.sum()),
                            "Corporate Card": (
                                frame["Corporate Card"].iloc[0]
                                if "Corporate Card" in frame.columns
                                and not frame["Corporate Card"].empty
                                else ""
                            ),
                        }
                    )

                summary = pd.DataFrame(summary_rows).sort_values("Last Name")
                st.success(f"Generated {len(claims)} claim files.")
                st.dataframe(summary, use_container_width=True)

                zip_buffer = io.BytesIO()
                with ZipFile(zip_buffer, "w") as zip_file:
                    for last_name, frame in claims.items():
                        file_name = (
                            f"{last_name}_AMEX_Claim.xlsx"
                            if export_format == "excel"
                            else f"{last_name}_AMEX_Claim.csv"
                        )
                        if export_format == "excel":
                            excel_bytes = io.BytesIO()
                            frame.to_excel(excel_bytes, index=False)
                            zip_file.writestr(file_name, excel_bytes.getvalue())
                        else:
                            csv_bytes = frame.to_csv(index=False).encode("utf-8")
                            zip_file.writestr(file_name, csv_bytes)
                zip_buffer.seek(0)

                st.download_button(
                    "⬇️ Download All Claim Files (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="amex_claims.zip",
                    mime="application/zip",
                )

                with st.expander("Preview first claim file"):
                    first_name, first_frame = next(iter(claims.items()))
                    st.write(f"{first_name}_AMEX_Claim")
                    st.dataframe(first_frame.head(), use_container_width=True)
        except Exception as exc:
            st.error(f"Processing failed: {exc}")
