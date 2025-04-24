import streamlit as st
import pandas as pd

st.title("📊 Cashflow Model Validator")
st.write("Upload your actuarial cashflow Excel file to verify calculations.")

# Upload file
uploaded_file = st.file_uploader("Choose an Excel file", type=[".xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    st.subheader("Raw Data Preview")
    st.dataframe(df)

    with st.expander("Validation Results"):
        errors = []

        # --- Recalculations ---
        df['Survival rate (calc)'] = 1.0
        for i in range(1, len(df)):
            df.loc[i, 'Survival rate (calc)'] = df.loc[i-1, 'Survival rate (calc)'] * (1 - df.loc[i, 'Death rate'])

        df['Expected Cashflow (calc)'] = df['Cashflow'] * df['Survival rate (calc)']
        df['Discounted Cashflow (calc)'] = df['Expected Cashflow (calc)'] * df['Discount rate.1']

        # --- Compare calculations ---
        df['Survival rate diff'] = abs(df['Survival rate'] - df['Survival rate (calc)'])
        df['Expected CF diff'] = abs(df['Expected Cashflow'] - df['Expected Cashflow (calc)'])
        df['Discounted CF diff'] = abs(df['Discounted cashflow'] - df['Discounted Cashflow (calc)'])

        tol = 1e-6  # tolerance for float comparison

        if any(df['Survival rate diff'] > tol):
            errors.append("Survival rate calculation mismatch detected.")
        if any(df['Expected CF diff'] > tol):
            errors.append("Expected Cashflow mismatch detected.")
        if any(df['Discounted CF diff'] > tol):
            errors.append("Discounted Cashflow mismatch detected.")

        pvfp_calc = df['Discounted Cashflow (calc)'].sum()
        pvfp_sheet = df.loc[0, 'PVFP'] if 'PVFP' in df.columns else None

        if pvfp_sheet is not None:
            if abs(pvfp_calc - pvfp_sheet) > tol:
                errors.append(f"PVFP mismatch. Calculated: {pvfp_calc:.2f}, Sheet: {pvfp_sheet:.2f}")

        # --- Output ---
        if not errors:
            st.success("All calculations are correct. ✅")
        else:
            st.error("Issues found in the following areas:")
            for err in errors:
                st.write(f"- {err}")

        st.subheader("Detailed Comparison Table")
        st.dataframe(df)
