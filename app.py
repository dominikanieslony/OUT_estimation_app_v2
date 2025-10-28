import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="📊 Campaign Estimator", layout="wide")

def load_excel_and_unmerge(file_bytes):
wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
ws = wb.active

```
# Usuń scalone komórki i wypełnij wartości
for merged_range in list(ws.merged_cells.ranges):
    ws.unmerge_cells(str(merged_range))
for col in ws.columns:
    prev_value = None
    for cell in col:
        if cell.value is not None:
            prev_value = cell.value
        else:
            cell.value = prev_value

data = ws.values
columns = next(data)
df = pd.DataFrame(data, columns=columns)
df = df.ffill(axis=0)

# Oczyść nagłówki kolumn
df.columns = (
    df.columns.astype(str)
    .str.strip()
    .str.replace('\u00A0', '', regex=False)
    .str.replace('\u202F', '', regex=False)
)

return df
```

def clean_demand_column(df):
def parse_demand(val):
if pd.isna(val):
return None
val = str(val).strip()
val = val.replace('€', '').replace(' ', '').replace('\u00A0', '').replace('\u202F', '')
val = val.replace(',', '.')
try:
return float(val)
except ValueError:
return None
if 'Demand' in df.columns:
df['Demand'] = df['Demand'].apply(parse_demand)
return df

def filter_data(df, country, search_filter, start_date, end_date, selected_category=None):
df_filtered = df[df['Country'] == country].copy()

```
# Wyczyść tekstowe kolumny
for col in ['Name', 'Description', 'Category']:
    if col in df_filtered.columns:
        df_filtered[col] = (
            df_filtered[col]
            .astype(str)
            .str.strip()
            .str.replace(r'[\u00A0\u202F]', '', regex=True)
        )

# Filtr kategorii
if selected_category and selected_category != "All":
    df_filtered = df_filtered[
        df_filtered['Category'].str.lower() == selected_category.strip().lower()
    ]

# Wspólne pole wyszukiwania (min 3 litery, szuka w Name OR Description)
if search_filter and len(search_filter.strip()) >= 3:
    pattern = search_filter.strip()
    mask_name = df_filtered['Name'].str.contains(pattern, case=False, na=False)
    mask_desc = df_filtered['Description'].str.contains(pattern, case=False, na=False)
    df_filtered = df_filtered[mask_name | mask_desc]

# Daty — pokazuje kampanie, które nachodzą na zakres
df_filtered['Start'] = pd.to_datetime(df_filtered['Start'], dayfirst=True, errors='coerce')
df_filtered['End'] = pd.to_datetime(df_filtered['End'], dayfirst=True, errors='coerce')

df_filtered = df_filtered[
    (df_filtered['End'] >= pd.to_datetime(start_date)) &
    (df_filtered['Start'] <= pd.to_datetime(end_date))
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

def reorder_columns(df):
cols = df.columns.tolist()
if 'Name' in cols and 'Description' in cols:
cols.remove('Description')
idx = cols.index('Name') + 1
cols.insert(idx, 'Description')
return df[cols]
return df

st.title("📊 Campaign Estimator (Excel Version)")

uploaded_file = st.file_uploader("📂 Upload campaign data Excel file", type=["xlsx", "xls"])

if uploaded_file:
try:
raw_bytes = uploaded_file.read()
df = load_excel_and_unmerge(raw_bytes)
df = clean_demand_column(df)

```
    required_cols = {'Country', 'Name', 'Description', 'Start', 'End', 'Demand', 'Category'}
    if not required_cols.issubset(df.columns):
        st.error(f"❌ Missing required columns: {required_cols - set(df.columns)}")
    else:
        country_list = df['Country'].dropna().unique().tolist()
        selected_country = st.selectbox("🌍 Select country:", country_list)

        categories = df['Category'].dropna().unique().tolist()
        categories = sorted(categories)
        selected_category = st.selectbox("🏷️ Select category:", ["All"] + categories)

        search_filter = st.text_input("🔎 Search campaigns by name or description (min 3 letters):")

        st.subheader("⏳ Earlier Period")
        earlier_start_date = st.date_input("Start date (Earlier Period):", key='earlier_start')
        earlier_end_date = st.date_input("End date (Earlier Period):", key='earlier_end')

        st.subheader("📈 Target growth from Earlier Period (%)")
        target_growth = st.number_input(
            "Enter growth percentage (can be negative):",
            min_value=-100, max_value=1000, step=1, format="%d"
        )

        st.subheader("⏳ Later Period")
        later_start_date = st.date_input("Start date (Later Period):", key='later_start')
        later_end_date = st.date_input("End date (Later Period):", key='later_end')

        earlier_filtered = filter_data(df, selected_country, search_filter, earlier_start_date, earlier_end_date, selected_category)
        later_filtered = filter_data(df, selected_country, search_filter, later_start_date, later_end_date, selected_category)

        earlier_filtered = reorder_columns(earlier_filtered)
        later_filtered = reorder_columns(later_filtered)

        st.subheader("Select campaigns to include from Earlier Period:")
        earlier_selections = {
            idx: st.checkbox(
                f"{row['Name']} | {row['Description']} | Start: {row['Start'].date()} | End: {row['End'].date()} | Demand: {row['Demand']}",
                value=True, key=f"earlier_{idx}"
            )
            for idx, row in earlier_filtered.iterrows()
        }

        st.subheader("Select campaigns to include from Later Period:")
        later_selections = {
            idx: st.checkbox(
                f"{row['Name']} | {row['Description']} | Start: {row['Start'].date()} | End: {row['End'].date()} | Demand: {row['Demand']}",
                value=True, key=f"later_{idx}"
            )
            for idx, row in later_filtered.iterrows()
        }

        earlier_selected_df = earlier_filtered.loc[[idx for idx, checked in earlier_selections.items() if checked]]
        later_selected_df = later_filtered.loc[[idx for idx, checked in later_selections.items() if checked]]

        if st.button("📈 Calculate Estimation"):
            if earlier_selected_df.empty and later_selected_df.empty:
                st.warning("⚠️ No campaigns selected in either period for estimation.")
            else:
                estimation = estimate_demand(earlier_selected_df, later_selected_df, target_growth)
                if estimation is None:
                    st.warning("⚠️ Unable to calculate estimation with the given data.")
                else:
                    st.success(f"Estimated Demand: **{estimation:.2f} EUR**")
                    st.markdown("### Data used for estimation:")

                    if not earlier_selected_df.empty:
                        st.write("Earlier Period Campaigns:")
                        st.dataframe(earlier_selected_df)

                    if not later_selected_df.empty:
                        st.write("Later Period Campaigns:")
                        st.dataframe(later_selected_df)

                    combined_df = pd.concat([earlier_selected_df, later_selected_df]).drop_duplicates()
                    csv = combined_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download selected campaigns data as CSV",
                        data=csv,
                        file_name='campaign_estimation_data.csv',
                        mime='text/csv'
                    )
except Exception as e:
    st.error(f"❌ Error processing file: {e}")
```
