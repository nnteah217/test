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

st.title("📂 Upload excel DATA to Convert")
uploaded_files = st.file_uploader("", type=["xlsx"], accept_multiple_files=True)

# === 2. Read and combine Excel files ===
all_dfs = []
invalid_files = []

if uploaded_files:
    for file in uploaded_files:
        if file.name.endswith(".xlsx"):
            match = re.search(r"(\d{4})M(\d+)", file.name)
            if match:
                year, month = int(match.group(1)), int(match.group(2))
                df = pd.read_excel(file, skiprows=4, na_values=[], keep_default_na=False)
                df["YEAR"] = year
                df["MONTH"] = month
                all_dfs.append(df)

    if all_dfs:
        # Input fields after successful upload
        CLOSING_M = st.number_input("Input the latest month:", min_value=1, max_value=12, step=1)
        CURRENCY = st.selectbox("Select the currency amount display:", ["LCC and EUR", "LCC only", "EUR only"])

        if CLOSING_M is not None and CURRENCY:
                df = pd.concat(all_dfs, ignore_index=True)

                # Columns we want to retain
                columns_origin = [
                    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
                    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
                    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
                    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
                    "AccountType", "Amount", "Amount In EUR", "YEAR", "MONTH"
                ]
                df = df[columns_origin]
                df["MONTH+1"]=df["MONTH"]+1

                columns_next =[
                    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
                    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
                    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
                    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
                    "AccountType", "YEAR", "MONTH+1"
                ]
                
                columns_id =[
                    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
                    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
                    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
                    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
                    "AccountType", "YEAR", "MONTH"
                ]
                # First, create the reference DataFrame with next month's aggregated values
                df_next = df.groupby(columns_next).agg({
                    "Amount": "sum",
                    "Amount In EUR": "sum"
                }).reset_index().rename(columns={
                "Amount" : "Amount_Next",
                "Amount In EUR": "Amount In EUR_Next",
                "MONTH+1" : "MONTH"
                })

                # Now merge back to original dataframe
                df = df.merge(df_next, how="left", on=columns_id).fillna(0)

                # Subtract current month - next month
                df["LCC AMOUNT"] = df["Amount"] - df["Amount_Next"]
                df["EUR AMOUNT"] = df["Amount In EUR"] - df["Amount In EUR_Next"]            

                df = df.drop(columns=["Amount", "Amount In EUR","MONTH+1"])
                df = df[(df["MONTH"] <= CLOSING_M)]   
                df = df[~((df["EUR AMOUNT"] == 0) & (df["LCC AMOUNT"] == 0))]
     
                # --- Final Output ---
                if CURRENCY == "LCC only":
                    df_final = df[
                    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
                    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
                    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
                    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
                    "AccountType", "LCC AMOUNT", "YEAR", "MONTH"
                    ]
                    df_final = df_final[~((df_final["LCC AMOUNT"] == 0))]
                elif CURRENCY == "EUR only":
                    df_final = df[
                    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
                    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
                    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
                    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
                    "AccountType", "EUR AMOUNT", "YEAR", "MONTH"
                    ]
                    df_final = df_final[~((df_final["EUR AMOUNT"] == 0))]
                elif CURRENCY == "LCC and EUR":
                    df_final = df[
                    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
                    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
                    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
                    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
                    "AccountType", "LCC AMOUNT", "EUR AMOUNT", "YEAR", "MONTH"
                    ]
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
