import pandas as pd
import numpy as np
import streamlit as st
import os
import re
from time import time
from datetime import datetime
from io import BytesIO

st.set_page_config(layout="wide")

def get_col_widths(df):
    return [max(df[col].astype(str).map(len).max(), len(col)) + 2 for col in df.columns]

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sheet_name = "MTD"
        df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=1, header=False)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        column_settings = [{'header': col} for col in df.columns]
        num_rows, num_cols = df.shape

        worksheet.add_table(0, 0, num_rows, num_cols - 1, {
            'name': "MTD",
            'columns': column_settings,
            'style': 'TableStyleLight8'
        })

        for i, width in enumerate(get_col_widths(df)):
            worksheet.set_column(i, i, width)

    return output.getvalue()

# === Streamlit UI ===
col1, col2 = st.columns(2)

with col1:
    st.title("EPM Monthly Display Converter")
    uploaded_files = st.file_uploader("", type=["xlsx"], accept_multiple_files=True)

    check_uploaded_files = []

    for file in uploaded_files:
        match = re.search(r"(\d{4})M(\d+)", file.name)
        if match:
            check_uploaded_files.append({
                "File": file.name,
                "YEAR": int(match.group(1)),
                "MONTH": int(match.group(2)),
                "VALID": True             
            })
        else:
            check_uploaded_files.append({
                "File": file.name,
                "YEAR": None,
                "MONTH": None,
                "VALID": False                
            })

    check_uploaded_files = pd.DataFrame(check_uploaded_files)
    check_uploaded_files["CONSECUTIVE"] = False
    # Loop through each row to determine consecutiveness
    for i, row in check_uploaded_files.iterrows():
        year = row["YEAR"]
        month = row["MONTH"]
    
        # Skip rows with invalid or missing data
        if pd.isna(year) or pd.isna(month):
            continue
    
        if month == 1:
            check_uploaded_files.at[i, "CONSECUTIVE"] = True
        else:
            prev_month = month - 1
            next_month = month + 1
    
            same_year_months = check_uploaded_files[check_uploaded_files["YEAR"] == year]["MONTH"].tolist()
    
            if prev_month in same_year_months or next_month in same_year_months:
                check_uploaded_files.at[i, "CONSECUTIVE"] = True

    if check_uploaded_files.empty:
        st.info("📂 Please upload Excel files to begin")
    else:
        st.success(f"📄 {len(check_uploaded_files)} file(s) uploaded")

    run_btn = False
    valid_files = True
    CLOSING_M = len(check_uploaded_files)

    if not check_uploaded_files.empty:
        if (~check_uploaded_files["VALID"]).any():
            st.warning("⚠️ All files must have [yyyy]M[mm] in the name")
            st.dataframe(check_uploaded_files)
            valid_files = False
        
        if check_uploaded_files["YEAR"].nunique() != 1:
            st.warning("⚠️ All files must have the same year")
            st.dataframe(check_uploaded_files)
            valid_files = False

        if check_uploaded_files["MONTH"].min() != 1:
            st.warning("⚠️ Files must start from M1")
            st.dataframe(check_uploaded_files)
            valid_files = False

        if check_uploaded_files["MONTH"].duplicated().any():
            st.warning("⚠️ Files must have unique months")
            st.dataframe(check_uploaded_files)
            valid_files = False

        if (~check_uploaded_files["CONSECUTIVE"]).any():
            st.warning("⚠️ Months within a year must be consecutive")
            st.dataframe(check_uploaded_files)
            valid_files = False

        if valid_files:
            CURRENCY = st.selectbox("Select currency amount:", ["LCC and EUR", "LCC only", "EUR only"])
            run_btn = st.button("🚀 Convert")

# === Run Conversion ===
if run_btn:
    start_time = time()
    with st.spinner("The file is being cooked..."):
        all_dfs = []

        for file in uploaded_files:
            match = re.search(r"(\d{4})M(\d+)", file.name)
            if match:
                year, month = int(match.group(1)), int(match.group(2))
                df = pd.read_excel(file, skiprows=4, na_values=[], keep_default_na=False)
                df["YEAR"] = year
                df["MONTH"] = month
                all_dfs.append(df)

        if all_dfs:
            df = pd.concat(all_dfs, ignore_index=True)
            df["MONTH+1"] = df["MONTH"] + 1

            columns_base = [
                "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
                "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
                "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
                "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
                "AccountType"
            ]
            columns_next = columns_base + ["YEAR", "MONTH+1"]
            columns_current = columns_base + ["YEAR", "MONTH"]

            df_next = df.groupby(columns_next).agg({
                "Amount": "sum",
                "Amount In EUR": "sum"
            }).reset_index().rename(columns={
                "Amount": "Amount_Next",
                "Amount In EUR": "Amount In EUR_Next",
                "MONTH+1": "MONTH"
            })

            df = df.merge(df_next, how="outer", on=columns_current).fillna(0)

            df["LCC AMOUNT"] = df["Amount"] - df["Amount_Next"]
            df["EUR AMOUNT"] = df["Amount In EUR"] - df["Amount In EUR_Next"]

            df = df.drop(columns=["Amount", "Amount In EUR", "Amount_Next", "Amount In EUR_Next", "MONTH+1"])
            df = df[df["MONTH"] <= CLOSING_M]
            df = df[~((df["EUR AMOUNT"] == 0) & (df["LCC AMOUNT"] == 0))]

            columns_final = columns_base + ["LCC AMOUNT", "EUR AMOUNT", "YEAR", "MONTH"]

            if CURRENCY == "LCC only":
                df_final = df[columns_final].drop(columns=["EUR AMOUNT"])
                df_final = df_final[df_final["LCC AMOUNT"] != 0]
            elif CURRENCY == "EUR only":
                df_final = df[columns_final].drop(columns=["LCC AMOUNT"])
                df_final = df_final[df_final["EUR AMOUNT"] != 0]
            else:
                df_final = df[columns_final]

            df_final = df_final.sort_values(by=["YEAR", "MONTH"])

            now = datetime.now()
            date_str = now.strftime("%y%m%d_%H%M")
            max_month = f"{CLOSING_M:02d}"
            currency_code = {
                "LCC only": "LCC",
                "EUR only": "EUR",
                "LCC and EUR": "LCCEUR"
            }[CURRENCY]

            output_filename = f"MTD{max_month}_{currency_code}_{date_str}.xlsx"
            excel_data = to_excel(df_final)

            elapsed_time = time() - start_time

            with col1:
                st.success(f"✅ Processing completed in {elapsed_time:.2f} seconds! Click below to download.")
                st.download_button(
                    label="📥 Download Converted File",
                    data=excel_data,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
