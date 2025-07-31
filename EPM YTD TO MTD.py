import pandas as pd
import numpy as np
import streamlit as st
import os
import re
from datetime import datetime
from io import BytesIO

st.set_page_config(layout="wide")

def get_col_widths(df):
    """Calculate appropriate column widths based on contents."""
    return [max(df[col].astype(str).map(len).max(), len(col)) + 2 for col in df.columns]

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sheet_name ="MTD"
        df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=1, header=False)

        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # Define column settings
        column_settings = [{'header': col} for col in df.columns]
        num_rows, num_cols = df.shape

        # Add an Excel table with filters and style
        worksheet.add_table(0, 0, num_rows, num_cols - 1, {
            'name': "MTD",
            'columns': column_settings,
            'style': 'TableStyleMedium9'  # Optional style
        })

        # Optional: Auto-fit column width
        for i, width in enumerate(get_col_widths(df)):
            worksheet.set_column(i, i, width)

    return output.getvalue()

# === Streamlit UI ===
st.title("📂 Upload Excel DATA to Convert")
uploaded_files = st.file_uploader("", type=["xlsx"], accept_multiple_files=True)
CLOSING_M = st.number_input("Input the latest month:", min_value=1, max_value=12, step=1)
CURRENCY = st.selectbox("Select the currency amount display:", ["LCC and EUR", "LCC only", "EUR only"])
run_btn = st.button("🚀 Run Processing")

# === Run When Button is Clicked ===
if run_btn:
    all_dfs = []
    invalid_files = []

    for file in uploaded_files:
        if file.name.endswith(".xlsx"):
            match = re.search(r"(\d{4})M(\d+)", file.name)
            if match:
                year, month = int(match.group(1)), int(match.group(2))
                try:
                    df = pd.read_excel(file, skiprows=4, na_values=[], keep_default_na=False)
                    df["YEAR"] = year
                    df["MONTH"] = month
                    all_dfs.append(df)
                except Exception as e:
                    invalid_files.append(f"{file.name} - Error: {e}")
            else:
                invalid_files.append(f"{file.name} - Filename does not match 'YYYYMx' format")

    # Show invalid file messages (but still proceed if valid ones exist)
    if invalid_files:
        st.error("Some files could not be processed:")
        for msg in invalid_files:
            st.markdown(f"- {msg}")

    # === Proceed if valid data exists ===
    if all_dfs:
        df = pd.concat(all_dfs, ignore_index=True)
        df["MONTH+1"] = df["MONTH"] + 1

        # Define common column base
        columns_base = [
            "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
            "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
            "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
            "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
            "AccountType"
        ]
        columns_next = columns_base + ["YEAR", "MONTH+1"]
        columns_current = columns_base + ["YEAR", "MONTH"]

        # Group and shift data for next month comparison
        df_next = df.groupby(columns_next).agg({
            "Amount": "sum",
            "Amount In EUR": "sum"
        }).reset_index().rename(columns={
            "Amount": "Amount_Next",
            "Amount In EUR": "Amount In EUR_Next",
            "MONTH+1": "MONTH"
        })

        # Merge current with next month values
        df = df.merge(df_next, how="outer", on=columns_current).fillna(0)

        # Calculate delta values
        df["LCC AMOUNT"] = df["Amount"] - df["Amount_Next"]
        df["EUR AMOUNT"] = df["Amount In EUR"] - df["Amount In EUR_Next"]

        # Clean up and filter
        df = df.drop(columns=["Amount", "Amount In EUR", "Amount_Next", "Amount In EUR_Next", "MONTH+1"])
        df = df[df["MONTH"] <= CLOSING_M]
        df = df[~((df["EUR AMOUNT"] == 0) & (df["LCC AMOUNT"] == 0))]

        # Prepare final output
        columns_final = columns_base + ["LCC AMOUNT", "EUR AMOUNT", "YEAR", "MONTH"]

        if CURRENCY == "LCC and EUR":
            df_final = df[columns_final]
        elif CURRENCY == "LCC only":
            df_final = df[columns_final].drop(columns=["EUR AMOUNT"])
            df_final = df_final[df_final["LCC AMOUNT"] != 0]
        elif CURRENCY == "EUR only":
            df_final = df[columns_final].drop(columns=["LCC AMOUNT"])
            df_final = df_final[df_final["EUR AMOUNT"] != 0]

        df_final = df_final.sort_values(by=["YEAR", "MONTH"])

        # Prepare download
        now = datetime.now()
        date_str = now.strftime("%y%m%d_%H%M")
        max_month = f"{CLOSING_M:02d}"
        currency_code = {
            "LCC only": "LCC",
            "EUR only": "EUR",
            "LCC and EUR": "LCCEUR"
        }.get(CURRENCY, "")

        output_filename = f"FASTCLOSE_{currency_code}_MTD{max_month}_{date_str}.xlsx"

        st.success("✅ Processing completed! Click below to download.")
        st.download_button(
            label="📥 Download Converted File",
            data=to_excel(df_final),
            file_name=output_filename
        )
    else:
        st.warning("⚠️ No valid Excel data found. Please upload the correct file(s).")

