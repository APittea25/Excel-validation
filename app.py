import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile

st.title("📊 Cashflow Model Validator")
st.write("Upload your actuarial cashflow Excel file to verify calculations and review formula logic.")

# Upload file
uploaded_file = st.file_uploader("Choose an Excel file", type=[".xlsx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name

    df = pd.read_excel(tmp_path, sheet_name=0)
    st.subheader("Raw Data Preview")
    st.dataframe(df)

    # Load workbook for formula analysis
    wb = load_workbook(tmp_path, data_only=False)
    ws = wb.active

    # Extract headers
    headers = [cell.value for cell in ws[1]]
    column_formulas = {header: [] for header in headers if header is not None}

    for col_idx, header in enumerate(headers, start=1):
        if header is None:
            continue
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            column_formulas[header].append(cell.value)

    column_analysis = {}
    column_descriptions = {
        "Time": "Represents the time period (e.g., year). This is an input.",
        "Cashflow": "Cash amount expected or paid at each time step. This is an input.",
        "Death rate": "Assumed probability of death in the period. This is an input",
        "Discount rate": "Base rate for discounting future values. Often constant or based on a curve. This is an input",
        "Survival rate": "Calculated as the previous survival rate multiplied by (1 - death rate). The survival rate is assumed to be 1 at Time = 1",
        "Discount factor": "Calculated as the previous discount factor divided by (1 + Discount rate). The discount factor is assumed to be 1 at Time = 1",
        "Expected Cashflow": "Calculated as Cashflow × Survival rate.",
        "Discounted cashflow": "Calculated as Expected Cashflow × Discount Factor.",
        "PVFP": "Present Value of Future Profits. =SUM of discounted cashflows.",
    }

    for header, values in column_formulas.items():
        formulas = [v for v in values if isinstance(v, str) and v.startswith('=')]
        description = column_descriptions.get(header, "No specific description available.")
        if not formulas:
            column_analysis[header] = f"✅ Hardcoded values. {description}"
        else:
            unique_formulas = set(formulas)
            if len(unique_formulas) == 1:
                column_analysis[header] = f"❗ Formula-driven: `{formulas[0]}`. {description}"
            else:
                column_analysis[header] = f"❗ Formula-driven (varied formulas). {description}"

    with st.expander("🔍 Column-by-Column Explanation"):
        for col, explanation in column_analysis.items():
            st.write(f"**{col}**: {explanation}")

    with st.expander("Validation Results"):
        errors = []

        # --- Recalculations ---
        df['Discount factor (calc)'] = 1.0
        for i in range(1, len(df)):
            df.loc[i, 'Discount factor (calc)'] = df.loc[i-1, 'Discount factor (calc)'] / (1 + df.loc[i, 'Discount rate'])

        df['Survival rate (calc)'] = 1.0
        for i in range(1, len(df)):
            df.loc[i, 'Survival rate (calc)'] = df.loc[i-1, 'Survival rate (calc)'] * (1 - df.loc[i, 'Death rate'])

        df['Expected Cashflow (calc)'] = df['Cashflow'] * df['Survival rate (calc)']
        df['Discounted Cashflow (calc)'] = df['Expected Cashflow (calc)'] * df['Discount rate.1']
        df['PVFP (calc)'] = df['Discounted Cashflow (calc)'].sum()

        df['Discount factor diff'] = abs(df['Discount rate.1'] - df['Discount factor (calc)'])

# --- Check differences ---
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

        pvfp_sheet = df.loc[0, 'PVFP'] if 'PVFP' in df.columns else None
        if pvfp_sheet is not None:
            if abs(df['PVFP (calc)'].iloc[0] - pvfp_sheet) > tol:
                errors.append(f"PVFP mismatch. Calculated: {df['PVFP (calc)'].iloc[0]:.2f}, Sheet: {pvfp_sheet:.2f}")

        # --- Output ---
        if not errors:
            st.success("All calculations are correct. ✅")
        else:
            st.error("Issues found in the following areas:")
            for err in errors:
                st.write(f"- {err}")

        st.subheader("Validation Tables")

        

        st.markdown("#### Survival Rate")
        

        st.markdown("#### Discount Factor")
        st.markdown("""**Calculation Description:** The discount factor is initialized at 1.0. For each subsequent period, it is calculated as 1 / (1 + previous period's discount rate).

```python
df['Discount factor (calc)'] = 1.0
for i in range(1, len(df)):
    df.loc[i, 'Discount factor (calc)'] = df.loc[i-1, 'Discount factor (calc)'] / (1 + df.loc[i, 'Discount rate'])
```""")
        st.dataframe(df[['Time', 'Discount rate.1', 'Discount factor (calc)', 'Discount factor diff']])
        st.markdown("""**Calculation Description:** Survival rate is calculated as the previous period's survival rate multiplied by (1 - death rate). The first value is set to 1.0.

```python
df['Survival rate (calc)'] = 1.0
for i in range(1, len(df)):
    df.loc[i, 'Survival rate (calc)'] = df.loc[i-1, 'Survival rate (calc)'] * (1 - df.loc[i, 'Death rate'])
```""")
        st.dataframe(df[['Time', 'Survival rate', 'Survival rate (calc)', 'Survival rate diff']])

        st.markdown("#### Expected Cashflow")
        st.markdown("""**Calculation Description:** Expected Cashflow is calculated by multiplying the Cashflow by the Survival rate.

```python
df['Expected Cashflow (calc)'] = df['Cashflow'] * df['Survival rate (calc)']
```""")
        st.dataframe(df[['Time', 'Expected Cashflow', 'Expected Cashflow (calc)', 'Expected CF diff']])

        st.markdown("#### Discounted Cashflow")
        st.markdown("""**Calculation Description:** Discounted Cashflow is calculated by multiplying the Expected Cashflow by the Discount factor.

```python
df['Discounted Cashflow (calc)'] = df['Expected Cashflow (calc)'] * df['Discount rate.1']
```""")
        st.dataframe(df[['Time', 'Discounted cashflow', 'Discounted Cashflow (calc)', 'Discounted CF diff']])

        if 'PVFP' in df.columns:
            st.markdown("#### PVFP")
            st.markdown("""**Calculation Description:** PVFP (Present Value of Future Profits) is the total of all discounted cashflows.

```python
df['PVFP (calc)'] = df['Discounted Cashflow (calc)'].sum()
```""")
            pvfp_df = pd.DataFrame({
                'PVFP (Excel)': [pvfp_sheet],
                'PVFP (Calculated)': [df['PVFP (calc)'].iloc[0]],
                'Difference': [abs(df['PVFP (calc)'].iloc[0] - pvfp_sheet)]
            })
            st.dataframe(pvfp_df)
