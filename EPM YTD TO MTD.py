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

col1, col2 = st.columns(2)
with col1:
    st.title("📁 Upload FASTCLOSE DATA to Convert")
    uploaded_file_FastClose = st.file_uploader("", type=["xlsx"], accept_multiple_files=True)
# === 2. Read and combine Excel files ===

all_dfs = []
CLOSING_M = st.number_input("Input the highest month:", min_value=1, max_value=12, step=1, format="%d")
CURRENCY = st.selectbox("Select the currency amount:", ["LCC and EUR","LCC only", "EUR only" ])

for file in uploaded_file_FastClose:
    if file.name.endswith(".xlsx"):
        match = re.search(r"(\d{4})M(\d+)", file.name)
        if match:
            year, month = int(match.group(1)), int(match.group(2))
            df = pd.read_excel(file, skiprows=4,na_values=[], keep_default_na=False).assign(YEAR=year, MONTH=month)
            all_dfs.append(df)
# Combine all monthly data
df = pd.concat(all_dfs, ignore_index=True)

# === 3. Select relevant columns ===
columns_needed = [
    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
    "AccountType", "Amount", "Amount In EUR", "YEAR", "MONTH"
]
df = df[columns_needed]

##EUR AMOUNT df
# === 1. Copy original dataframe for EUR processing ===
df_EUR = df.copy()

# === 2. Create 12 monthly MTD adjustment columns for EUR ===
for m in range(1, 13):
    df_EUR[str(m)] = np.where(
        df_EUR["MONTH"] == m, df_EUR["Amount In EUR"],
        np.where(df_EUR["MONTH"] == m - 1, -df_EUR["Amount In EUR"], 0)
    )

# === 3. Prepare list of columns for melting (unpivoting) ===
monthly_columns_EUR = [str(m) for m in range(1, 13)]
id_columns_EUR = [col for col in df_EUR.columns if col not in monthly_columns_EUR]

# === 4. Melt the monthly columns to rows ===
df_EUR = df_EUR.melt(
    id_vars=id_columns_EUR,
    value_vars=monthly_columns_EUR,
    var_name="MONTH_NEW_EUR",
    value_name="EUR AMOUNT"
)

# === 5. Convert MONTH_NEW to integer for filtering ===
df_EUR["MONTH_NEW_EUR"] = df_EUR["MONTH_NEW_EUR"].astype(int)

# === 6. Filter non-zero MTD values and valid months based on closing month ===
df_EUR = df_EUR[
    (df_EUR["EUR AMOUNT"] != 0) &
    (df_EUR["MONTH_NEW_EUR"] <= CLOSING_M)
]

# === 7. Replace original MONTH column with MONTH_NEW_EUR ===
df_EUR = df_EUR.drop(columns=["MONTH"]).rename(columns={"MONTH_NEW_EUR": "MONTH"})

# === 8. Group by all dimensions and sum the EUR values ===
group_keys_EUR = [
    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
    "AccountType", "YEAR", "MONTH"
]

df_EUR = df_EUR.groupby(group_keys_EUR, dropna=False, as_index=False).agg({"EUR AMOUNT": "sum"})

# === 9. Final sort and filter non-zero values ===
df_EUR = df_EUR.sort_values(by=["YEAR", "MONTH"])
df_EUR = df_EUR[df_EUR["EUR AMOUNT"] != 0]

## LCC AMOUNT df
# === 1. Copy original dataframe for LCC processing ===
df_LCC = df.copy()

# === 2. Create 12 monthly MTD adjustment columns for LCC ===
for m in range(1, 13):
    df_LCC[str(m)] = np.where(
        df_LCC["MONTH"] == m, df_LCC["Amount"],
        np.where(df_LCC["MONTH"] == m - 1, -df_LCC["Amount"], 0)
    )

# === 3. Prepare list of columns for melting (unpivoting) ===
monthly_columns_LCC = [str(m) for m in range(1, 13)]
id_columns_LCC = [col for col in df_LCC.columns if col not in monthly_columns_LCC]

# === 4. Melt the monthly columns to rows ===
df_LCC = df_LCC.melt(
    id_vars=id_columns_LCC,
    value_vars=monthly_columns_LCC,
    var_name="MONTH_NEW_LCC",
    value_name="LCC AMOUNT"
)

# === 5. Convert MONTH_NEW to integer for filtering ===
df_LCC["MONTH_NEW_LCC"] = df_LCC["MONTH_NEW_LCC"].astype(int)

# === 6. Filter non-zero MTD values and valid months based on closing month ===
df_LCC = df_LCC[
    (df_LCC["LCC AMOUNT"] != 0) &
    (df_LCC["MONTH_NEW_LCC"] <= CLOSING_M)
]

# === 7. Replace original MONTH column with MONTH_NEW_LCC ===
df_LCC = df_LCC.drop(columns=["MONTH"]).rename(columns={"MONTH_NEW_LCC": "MONTH"})

# === 8. Group by all dimensions and sum the LCC values ===
group_keys_LCC = [
    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
    "AccountType", "YEAR", "MONTH"
]

df_LCC = df_LCC.groupby(group_keys_LCC, dropna=False, as_index=False).agg({"LCC AMOUNT": "sum"})

# === 9. Final sort and filter non-zero values ===
df_LCC = df_LCC.sort_values(by=["YEAR", "MONTH"])
df_LCC = df_LCC[df_LCC["LCC AMOUNT"] != 0]

# Create df_final based on selection
if CURRENCY == "LCC only":
    df_final = df_LCC.copy()
    df_final = df_final[[
    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
    "AccountType","LCC AMOUNT", "YEAR", "MONTH"
]]

elif CURRENCY == "EUR only":
    df_final = df_EUR.copy()
    df_final = df_final[[
    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
    "AccountType","EUR AMOUNT", "YEAR", "MONTH"
]]

elif CURRENCY == "LCC and EUR":
    # Merge both DataFrames on all dimensions
    base_columns = [
    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
    "AccountType", "YEAR", "MONTH"]
    df_final = pd.merge(df_LCC, df_EUR, on=base_columns, how="outer")
    df_final["LCC AMOUNT"] = df_final["LCC AMOUNT"].fillna(0)
    df_final["EUR AMOUNT"] = df_final["EUR AMOUNT"].fillna(0)
    df_final = df_final[[
    "Entity", "Cons", "Scenario", "View", "Account Parent", "Account", "Flow", "Origin", "IC",
    "FinalClient Group", "FinalClient", "Client", "FinancialManager", "Governance Level",
    "Governance", "Commodity", "AuditID", "UD8", "Project", "Employee", "Supplier",
    "InvoiceType", "ContractType", "AmountCurrency", "IntercoType", "ICDetails", "EmployedBy",
    "AccountType", "LCC AMOUNT","EUR AMOUNT", "YEAR", "MONTH"
]]

else:
    st.error("Invalid currency selection.")
    st.stop()

df = df_final.sort_values(by=["YEAR", "MONTH"])

# === 9. Export to Excel ===
# Get current date and time
now = datetime.now()
date_str = now.strftime("%y%m%d_%H%M")

# Get max month from the data
max_month = f"{CLOSING_M:02d}"  # zero-padded to 2 digits

currency_choice=("")
if CURRENCY == "LCC only":
  currency_choice = "LCC"
elif CURRENCY == "EUR only":
  currency_choice = "EUR"
elif CURRENCY == "LCC and EUR":
  currency_choice = "LCCEUR"

output_filename = f"FASTCLOSE_{currency_choice}_MTD{max_month}_{date_str}.xlsx"

with col2:  
        st.download_button(
            label="📥 Download Here",
            data=to_excel(df),
            file_name=output_filename)


