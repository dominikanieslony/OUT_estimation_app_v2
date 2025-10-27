import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import re

st.title("📊 Campaign Estimator (Excel version, cleaned headers)")

@st.cache_data
def load_excel_and_unmerge(file_bytes):
# Załaduj plik z openpyxl i usuń scalenia komórek
wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
ws = wb.active

```
for merged in list(ws.merged_cells.ranges):
    ws.unmerge_cells(str(merged))
    top_left_value = ws[merged.coord.split(':')[0]].value
    for row in ws[str(merged)]:
        for cell in row:
            cell.value = top_left_value

data = ws.values
cols = next(data)
df = pd.DataFrame(data, columns=cols)

# Wypełnij brakujące wartości (po scaleniu komórek)
df = df.ffill(axis=0)

# Oczyść nagłówki kolumn — usuń spacje i niełamliwe znaki
df.columns = (
    df.columns.astype(str)
    .str.strip()
    .str.replace('\u00A0', '', regex=False)
)

return df
```

def clean_demand_column(df):
def parse_demand(val):
if pd.isna(val):
return None
val = str(val).replace('\u00A0', '').replace(' ', '')
val = re.sub(r'[^\d,.-]', '', val)
# Przypadek liczby w formacie europejskim: 54 332 -> 54332
# Zamieniamy przecinek na kropkę, jeśli występuje jako separator dziesiętny
if val.count(',') == 1 and val.count('.') == 0:
val = val.replace(',', '.')
try:
num = float(val)
# Wyeliminuj błędnie sparsowane ogromne liczby
if num > 1e9 or num < -1e9:
return None
return num
except ValueError:
return None

```
if 'Demand' in df.columns:
    df['Demand'] = df['Demand'].apply(parse_demand)
    valid_count = df['Demand'].notna().sum()
    invalid_count = df['Demand'].isna().sum()
    st.info(f"📈 Demand column cleaned — valid: {valid_count}, invalid: {invalid_count}")
else:
    st.warning("⚠️ Column 'Demand' not found in file.")
return df
```

def filter_data(df, name_filter, desc_filter, start_date, end_date):
df_filtered = df.copy()

```
if 'Name' in df_filtered.columns and name_filter and len(name_filter) >= 3:
    mask_name = df_filtered['Name'].astype(str).str.contains(name_filter, case=False, na=False)
    df_filtered = df_filtered[mask_name]

if 'Description' in df_filtered.columns and desc_filter and len(desc_filter) >= 3:
    mask_desc = df_filtered['Description'].astype(str).str.contains(desc_filter, case=False, na=False)
    df_filtered = df_filtered[mask_desc]

if 'Start' in df_filtered.columns and 'End' in df_filtered.columns:
    df_filtered['Start'] = pd.to_datetime(df_filtered['Start'], errors='coerce')
    df_filtered['End'] = pd.to_datetime(df_filtered['End'], errors='coerce')
    df_filtered = df_filtered[
        (df_filtered['Start'] >= pd.to_datetime(start_date)) &
        (df_filtered['End'] <= pd.to_datetime(end_date))
    ]

return df_filtered
```

def estimate_demand(earlier_df, later_df, percentage):
earlier_mean = earlier_df['Demand'].mean() if not earlier_df.empty else 0
later_mean = later_df['Demand'].mean() if not later_df.empty else 0
adjusted_earlier = earlier_mean * (1 + percentage / 100)
if earlier_df.empty and later_df.empty:
return None
elif earlier_df.empty:
return later_mean
elif later_df.empty:
return adjusted_earlier
else:
return (adjusted_earlier + later_mean) / 2

uploaded_file = st.file_uploader("📂 Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
try:
raw_bytes = uploaded_file.read()
df = load_excel_and_unmerge(raw_bytes)

```
    # Upewnij się, że kolumny są prawidłowo rozpoznane po czyszczeniu nagłówków
    st.write("✅ Columns found:", list(df.columns))

    required_cols = {'Start', 'End', 'Name', 'Description', 'Demand'}
    if not required_cols.issubset(df.columns):
        st.error(f"❌ Missing required columns: {required_cols - set(df.columns)}")
    else:
        df = clean_demand_column(df)

        st.subheader("🔍 Filters")
        name_filter = st.text_input("Filter by Name (min 3 letters):")
        desc_filter = st.text_input("Filter by Description (min 3 letters):")

        st.subheader("⏳ Earlier Period")
        earlier_start_date = st.date_input("Start date (Earlier Period):", key='earlier_start')
        earlier_end_date = st.date_input("End date (Earlier Period):", key='earlier_end')

        st.subheader("📈 Target Growth (%)")
        target_growth = st.number_input(
            "Enter growth percentage (can be negative):",
            min_value=-100, max_value=1000, step=1, format="%d"
        )

        st.subheader("⏳ Later Period")
        later_start_date = st.date_input("Start date (Later Period):", key='later_start')
        later_end_date = st.date_input("End date (Later Period):", key='later_end')

        earlier_filtered = filter_data(df, name_filter, desc_filter, earlier_start_date, earlier_end_date)
        later_filtered = filter_data(df, name_filter, desc_filter, later_start_date, later_end_date)

        st.write("Earlier Period Data:")
        st.dataframe(earlier_filtered)
        st.write("Later Period Data:")
        st.dataframe(later_filtered)

        if st.button("📊 Calculate Estimation"):
            if earlier_filtered.empty and later_filtered.empty:
                st.warning("⚠️ No data in selected periods for estimation.")
            else:
                estimation = estimate_demand(earlier_filtered, later_filtered, target_growth)
                if estimation is None:
                    st.warning("⚠️ Unable to calculate estimation with the given data.")
                else:
                    st.success(f"💶 Estimated Demand: **{estimation:,.2f} EUR**")

except Exception as e:
    st.error(f"❌ Error processing file: {e}")
```
