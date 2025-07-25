import pandas as pd
from io import BytesIO
import streamlit as st
st.set_page_config(layout="wide")
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

col1, col2, col3 = st.columns(3)
with col1:
    st.title("📁 File M")
    uploaded_file = st.file_uploader("", type=["xlsx"], key='fileM')
with col2:
    st.title("📁 File M - 1")
    uploaded_file2 = st.file_uploader("", type=["xlsx"], key='fileM-1')

with col3:
    if uploaded_file is not None and uploaded_file2 is not None:
        df1 = pd.read_excel(uploaded_file, skiprows=4)
        df2 = pd.read_excel(uploaded_file2, skiprows=4)
        st.success("✅ Files uploaded successfully!")
        
        if df1 is not None and df2 is not None:
            key_cols = ['Entity', 'Cons', 'Scenario', 'View', 'Account Parent', 'Account',
                        'Flow', 'Origin', 'IC', 'FinalClient Group', 'FinalClient', 'Client',
                        'FinancialManager', 'Governance Level', 'Governance', 'Commodity',
                        'AuditID', 'UD8', 'Project', 'Employee', 'Supplier', 'InvoiceType',
                        'ContractType', 'AmountCurrency', 'IntercoType', 'ICDetails', 'EmployedBy', 'AccountType']
            # Groupby to prevent the Cartesian product due to duplication:

            df1 = df1.groupby(key_cols, as_index=False, dropna=False).agg({'Amount': 'sum', 'Amount In EUR': 'sum'})
            df2 = df2.groupby(key_cols, as_index=False, dropna=False).agg({'Amount': 'sum', 'Amount In EUR': 'sum'})
            
            # Merge February and January on all identifying columns
            merged = pd.merge(df1, df2, on=key_cols, how= 'outer', suffixes=('_M', '_M-1'))
            merged = merged.fillna(0)
            

            # Calculate monthly February amount
            merged['M_Monthly_Amount'] = merged['Amount_M'] - merged['Amount_M-1']

            merged['M_Monthly_EURAmount'] = merged['Amount In EUR_M'] - merged['Amount In EUR_M-1']
            
        st.download_button(
            label="📥 Download Here",
            data=to_excel(merged),
            file_name="merged_output.xlsx",)
            # mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("⬆️ Please upload both files to proceed.")
