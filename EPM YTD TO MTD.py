import pandas as pd
import numpy as np
import streamlit as st
import os
import re
from datetime import datetime
from io import BytesIO

st.set_page_config(layout="wide")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.title("📂 Upload FASTCLOSE DATA to Convert")
uploaded_files_FastClose = st.file_uploader("", type=["xlsx"], accept_multiple_files=True)

# === 2. Read and combine Excel files ===
all_dfs = []
invalid_files = []

if uploaded_files_FastClose:
    for file in uploaded_files_FastClose:
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

    # Show error if any file failed
    if invalid_files:
        st.error("Some files could not be processed:")
        for msg in invalid_files:
            st.markdown(f"- {msg}")

    if all_dfs:
        # Input fields after successful upload
        CLOSING_M = st.number_input("Input the latest month:", min_value=1, max_value=12, step=1)
        CURRENCY = st.selectbox("Select the currency amount display:", ["LCC and EUR", "LCC only", "EUR only"])

        if CLOSING_M and CURRENCY:
            try:
                df = pd.concat(all_dfs, ignore_index=True)

                # Columns we want to retain
                columns_needed = [
                    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
                    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
                    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
                    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
                    "AccountType", "Amount", "Amount In EUR", "YEAR", "MONTH"
                ]
                df = df[columns_needed]

                # --- EUR Processing ---
                df_EUR = df.copy()
                for m in range(1, 13):
                    df_EUR[str(m)] = np.where(
                        df_EUR["MONTH"] == m, df_EUR["Amount In EUR"],
                        np.where(df_EUR["MONTH"] == m - 1, -df_EUR["Amount In EUR"], 0)
                    )

                df_EUR = df_EUR.melt(
                    id_vars=[col for col in df_EUR.columns if col not in [str(m) for m in range(1, 13)]],
                    value_vars=[str(m) for m in range(1, 13)],
                    var_name="MONTH_NEW_EUR", value_name="EUR AMOUNT"
                )
                df_EUR["MONTH_NEW_EUR"] = df_EUR["MONTH_NEW_EUR"].astype(int)
                df_EUR = df_EUR[(df_EUR["EUR AMOUNT"] != 0) & (df_EUR["MONTH_NEW_EUR"] <= CLOSING_M)]
                df_EUR = df_EUR.drop(columns=["MONTH"]).rename(columns={"MONTH_NEW_EUR": "MONTH"})

                group_keys = [col for col in columns_needed if col not in ["Amount", "Amount In EUR"]]
                df_EUR = df_EUR.groupby(group_keys, dropna=False, as_index=False).agg({"EUR AMOUNT": "sum"})

                # --- LCC Processing ---
                df_LCC = df.copy()
                for m in range(1, 13):
                    df_LCC[str(m)] = np.where(
                        df_LCC["MONTH"] == m, df_LCC["Amount"],
                        np.where(df_LCC["MONTH"] == m - 1, -df_LCC["Amount"], 0)
                    )

                df_LCC = df_LCC.melt(
                    id_vars=[col for col in df_LCC.columns if col not in [str(m) for m in range(1, 13)]],
                    value_vars=[str(m) for m in range(1, 13)],
                    var_name="MONTH_NEW_LCC", value_name="LCC AMOUNT"
                )
                df_LCC["MONTH_NEW_LCC"] = df_LCC["MONTH_NEW_LCC"].astype(int)
                df_LCC = df_LCC[(df_LCC["LCC AMOUNT"] != 0) & (df_LCC["MONTH_NEW_LCC"] <= CLOSING_M)]
                df_LCC = df_LCC.drop(columns=["MONTH"]).rename(columns={"MONTH_NEW_LCC": "MONTH"})
                df_LCC = df_LCC.groupby(group_keys, dropna=False, as_index=False).agg({"LCC AMOUNT": "sum"})

                # --- Final Output ---
                if CURRENCY == "LCC only":
                    df_final = df_LCC
                elif CURRENCY == "EUR only":
                    df_final = df_EUR
                else:  # "LCC and EUR"
                    df_final = pd.merge(df_LCC, df_EUR, on=group_keys + ["MONTH"], how="outer")
                    df_final["LCC AMOUNT"] = df_final["LCC AMOUNT"].fillna(0)
                    df_final["EUR AMOUNT"] = df_final["EUR AMOUNT"].fillna(0)

                df_final = df_final.sort_values(by=["YEAR", "MONTH"])

                # --- Export ---
                now = datetime.now()
                date_str = now.strftime("%y%m%d_%H%M")
                max_month = f"{CLOSING_M:02d}"

                currency_choice = {
                    "LCC only": "LCC",
                    "EUR only": "EUR",
                    "LCC and EUR": "LCCEUR"
                }.get(CURRENCY, "")

                output_filename = f"FASTCLOSE_{currency_choice}_MTD{max_month}_{date_str}.xlsx"
                st.download_button(
                    label="📥 Download Converted File",
                    data=to_excel(df_final),
                    file_name=output_filename
                )

            except Exception as e:
                st.error(f"Processing failed: {e}")
    else:
        st.warning("⚠️No valid Excel data found. Please upload the correct file(s).")
else:
    st.info("📂Please upload your FastClose Excel files to continue.")
