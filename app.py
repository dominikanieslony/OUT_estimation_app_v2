import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from io import BytesIO

st.title("📊 Campaign Revenue Predictor (Merged Excel Support)")

# === 1️⃣ Funkcja usuwająca scalenia i uzupełniająca wartości ===
def unmerge_excel_cells(file):
    try:
        in_memory_file = BytesIO(file.read())
        wb = load_workbook(in_memory_file)
        all_dfs = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # Rozpakowanie scalonych komórek
            merged_ranges = list(ws.merged_cells.ranges)
            for merged_range in merged_ranges:
                top_left = merged_range.min_row, merged_range.min_col
                value = ws.cell(*top_left).value
                ws.unmerge_cells(str(merged_range))
                # Wstaw wartość do wszystkich komórek byłego scalenia
                for row in ws.iter_rows(
                    min_row=merged_range.min_row,
                    max_row=merged_range.max_row,
                    min_col=merged_range.min_col,
                    max_col=merged_range.max_col
                ):
                    for cell in row:
                        cell.value = value

            # Konwersja arkusza na DataFrame
            data = ws.values
            columns = next(data)
            df = pd.DataFrame(data, columns=columns)
            all_dfs.append(df)

        wb.close()

        # Połączenie wszystkich arkuszy
        combined_df = pd.concat(all_dfs, ignore_index=True)
        return combined_df

    except Exception as e:
        st.error(f"❌ Error processing merged Excel cells: {e}")
        return pd.DataFrame()

# === 2️⃣ Czyszczenie kolumny Demand ===
def clean_demand_column(df):
    def parse_demand(val):
        if pd.isna(val):
            return None
        val = str(val)
        val = val.replace('€', '').replace(' ', '')
        val = val.replace('.', '').replace(',', '.')
        try:
            return float(val)
        except ValueError:
            return None
    if 'Demand' in df.columns:
        df['Demand'] = df['Demand'].apply(parse_demand)
    return df

# === 3️⃣ Filtrowanie danych po tekście i dacie ===
def filter_data(df, text_filter, start_date, end_date):
    df_filtered = df.copy()

    # Konwersja dat
    if 'Start' in df_filtered.columns:
        df_filtered['Start'] = pd.to_datetime(df_filtered['Start'], errors='coerce')
    if 'End' in df_filtered.columns:
        df_filtered['End'] = pd.to_datetime(df_filtered['End'], errors='coerce')

    # Filtrowanie po dacie
    if 'Start' in df_filtered.columns and 'End' in df_filtered.columns:
        df_filtered = df_filtered[
            (df_filtered['Start'] >= pd.to_datetime(start_date)) &
            (df_filtered['End'] <= pd.to_datetime(end_date))
        ]

    # Filtrowanie po tekście w Name lub Description
    if text_filter and len(text_filter) >= 2:
        mask_name = df_filtered['Name'].astype(str).str.contains(text_filter, case=False, na=False)
        mask_desc = df_filtered['Description'].astype(str).str.contains(text_filter, case=False, na=False)
        df_filtered = df_filtered[mask_name | mask_desc]

    return df_filtered

# === 4️⃣ Obliczanie średniego przychodu ===
def calculate_average_demand(df):
    if df.empty or 'Demand' not in df.columns:
        return None
    valid_values = df['Demand'].dropna()
    if valid_values.empty:
        return None
    return valid_values.mean()

# === 📂 Upload pliku ===
uploaded_file = st.file_uploader("📥 Upload Excel file (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file:
    df = unmerge_excel_cells(uploaded_file)

    if not df.empty:
        df = clean_demand_column(df)

        # Wyświetlenie podglądu
        with st.expander("🔍 Preview loaded data"):
            st.dataframe(df.head(50))

        # === 🔎 Filtry użytkownika ===
        st.subheader("🔧 Filter campaigns")
        text_filter = st.text_input("Search by Name or Description (min 2 letters):")

        st.subheader("📆 Select time period")
        min_date = pd.to_datetime(df['Start'], errors='coerce').min()
        max_date = pd.to_datetime(df['End'], errors='coerce').max()
        start_date = st.date_input("Start date:", min_date if pd.notna(min_date) else datetime(2024, 1, 1))
        end_date = st.date_input("End date:", max_date if pd.notna(max_date) else datetime.today())

        # === 📉 Filtrowanie danych ===
        filtered_df = filter_data(df, text_filter, start_date, end_date)

        if filtered_df.empty:
            st.warning("⚠️ No data found for selected filters.")
        else:
            st.success(f"✅ {len(filtered_df)} records found.")

            st.subheader("📊 Filtered campaigns:")
            st.dataframe(filtered_df)

            # === 📈 Obliczanie średniego przychodu ===
            avg_demand = calculate_average_demand(filtered_df)
            if avg_demand is not None:
                st.success(f"💰 Estimated Average Revenue (Demand): **{avg_demand:.2f} EUR**")
            else:
                st.warning("⚠️ Could not calculate average revenue (missing or invalid Demand values).")

            # === 💾 Pobranie danych ===
            csv = filtered_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download filtered data as CSV",
                data=csv,
                file_name='filtered_campaign_data.csv',
                mime='text/csv'
            )
    else:
        st.error("❌ No data found in Excel file.")

