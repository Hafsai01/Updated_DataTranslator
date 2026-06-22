import streamlit as st

st.title("Data Translator Portal")
st.write("Welcome! Choose a page from the sidebar to get started.")
st.markdown("""
- **AR Translator** → Customer data processing  
- **AP Translator** → Vendor data processing
- **GL Translator** → General Ledger account mapping (Maestro → Sage)
- **Updated GL Translator** → GL Translator with Debit & Credit support
""")
