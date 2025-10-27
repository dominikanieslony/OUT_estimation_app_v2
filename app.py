import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import re
from io import BytesIO

st.set_page_config(page_title="Campaign Estimator", layout="wide")
st.title("📊 Campaign Estimator (Excel - improved Demand parsing)")

@st.cache_data
def load_excel_and_unmerge(uploaded_file_bytes):
    """
    Wczytuje arkusz aktywny z Excela (openpyxl), usuwa scalenia i zwraca DataFrame.
    upload_file_bytes: bytes-like (BytesIO or uploaded_file.read())
    """
    in_memory = BytesIO(uploaded_file_bytes)
    wb = load_workbook(in_memory, data_only=True)
    ws = wb.active

    # Rozbij scalone komórki i wypełnij wartością z lewego górnego rogu
    for merged_range in list(ws.merged_cells.ranges):
        # pobierz wartość z lewego-górnego rogu
        tl = (merged_range.min_row, merged_range.min_col)
        top_left_value = ws.cell(row=tl[0], column=tl[1]).value
        # unmerge i wypełnij cały zakres
        ws.unmerge_cells(range_string=str(merged_range))
        for r in ws.iter_rows(
            min_row=merged_range.min_row,
            max_row=merged_range.max_row,
            min_col=merged_range.min_col,
            max_col=merged_range.max_col,
        ):
            for cell in r:
                cell.value = top_left_value

    # Konwersja arkusza do DataFrame (pierwszy wiersz nagłówki)
    data_iter = ws.values
    try:
        headers = next(data_iter)
    except StopIteration:
        wb.close()
        return pd.DataFrame()
    df = pd.DataFrame(data_iter, columns=headers)
    wb.close()
    # szybkie wypełnienie ewentualnych NaN po scaleniu (dodatkowa ochrona)
    df = df.ffill(axis=0)
    return df

def robust_clean_demand(df, demand_col='Demand'):
    """
    Silne czyszczenie kolumny Demand oraz heurystyczna korekcja ekstremalnie dużych wartości.
    Zwraca (df, diagnostics) gdzie diagnostics to dict z liczbami poprawek.
    """
    diagnostics = {
        'total': 0,
        'parsed_as_number': 0,
        'parsed_as_nan': 0,
        'suspicious_count': 0,
        'corrected_count': 0
    }

    if demand_col not in df.columns:
        return df, diagnostics

    # Funkcja jednego rekordu
    def parse_single(val):
        diagnostics['total'] += 1
        if pd.isna(val):
            diagnostics['parsed_as_nan'] += 1
            return None

        # Jeśli już jest liczba (int/float), pozostaw do dalszej weryfikacji
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            diagnostics['parsed_as_number'] += 1
            return float(val)

        s = str(val)

        # usuń spacje zwykłe i niełamliwe
        s = s.replace('\u00A0', '').replace(' ', '')

        # usuń wszystko poza cyframi, przecinkiem, kropką i minus
        s = re.sub(r'[^\d,.\-]', '', s)

        # jeśli jest jeden przecinek i brak kropki, traktujemy przecinek jako separator dziesiętny
        if s.count(',') == 1 and s.count('.') == 0:
            s = s.replace(',', '.')

        # jeśli nic się nie pozostało, zwróć None
        if s == '' or s == '-' or s == '.' or s == ',': 
            diagnostics['parsed_as_nan'] += 1
            return None

        try:
            num = float(s)
            diagnostics['parsed_as_number'] += 1
            return num
        except ValueError:
            diagnostics['parsed_as_nan'] += 1
            return None

    # Wstępne parsowanie wszystkich wartości (bez korekcji outlierów)
    parsed = df[demand_col].apply(parse_single)

    # Heurystyka: znajdź typową długość cyfr dla "normalnych" wartości
    # bierzemy tylko liczby mniejsze niż 1e8 (czyli sensowny próg)
    normal_numbers = parsed[(parsed.notna()) & (parsed.abs() < 1e8)]
    if not normal_numbers.empty:
        # mediana długości cyfrowej bez separatorów
        lengths = normal_numbers.astype(int).astype(str).str.replace('-', '').str.replace('.', '').str.len()
        typical_len = int(lengths.median()) if not lengths.empty else 5
    else:
        typical_len = 5  # domyślna typowa długość (np. 54332 → 5)

    # Korekcja wartości ekstremalnie dużych: jeśli > 1e8 -> potraktuj jako podejrzane
    sus_threshold = 1e8
    corrected = 0
    suspicious = 0
    final_values = []

    for v in parsed:
        if pd.isna(v):
            final_values.append(None)
            continue
        if abs(v) > sus_threshold:
            suspicious += 1
            # zamień na string cyfr (bez kropki, minus)
            sv = str(int(abs(v)))  # bierzemy część bez ułamka
            # Przytnij do typowej długości i przywróć ewentualny znak minus
            trimmed = sv[:typical_len]
            try:
                new_val = float(trimmed)
                # zachowaj znak
                if v < 0:
                    new_val = -new_val
                final_values.append(new_val)
                corrected += 1
            except Exception:
                final_values.append(None)
        else:
            final_values.append(v)

    diagnostics['suspicious_count'] = suspicious
    diagnostics['corrected_count'] = corrected
    diagnostics['parsed_as_number'] = int(diagnostics['parsed_as_number'])
    diagnostics['parsed_as_nan'] = int(diagnostics['parsed_as_nan'])
    diagnostics['total'] = int(diagnostics['total'])

    df[demand_col] = pd.Series(final_values, index=df.index).astype('float64')

    return df, diagnostics

def filter_data(df, name_filter, desc_filter, start_date, end_date):
    df_filtered = df.copy()
    if name_filter and len(name_filter) >= 2:
        if 'Name' in df_filtered.columns:
            df_filtered = df_filtered[df_filtered['Name'].astype(str).str.contains(name_filter, case=False, na=False)]
    if desc_filter and len(desc_filter) >= 2:
        if 'Description' in df_filtered.columns:
            df_filtered = df_filtered[df_filtered['Description'].astype(str).str.contains(desc_filter, case=False, na=False)]
    if 'Start' in df_filtered.columns and 'End' in df_filtered.columns:
        df_filtered['Start'] = pd.to_datetime(df_filtered['Start'], errors='coerce')
        df_filtered['End'] = pd.to_datetime(df_filtered['End'], errors='coerce')
        df_filtered = df_filtered[
            (df_filtered['Start'] >= pd.to_datetime(start_date)) &
            (df_filtered['End'] <= pd.to_datetime(end_date))
        ]
    return df_filtered

def estimate_demand(earlier_df, later_df, percentage):
    earlier_mean = earlier_df['Demand'].mean() if (earlier_df is not None and not earlier_df.empty) else 0
    later_mean = later_df['Demand'].mean() if (later_df is not None and not later_df.empty) else 0
    adjusted_earlier = earlier_mean * (1 + percentage / 100)
    if (earlier_df is None or earlier_df.empty) and (later_df is None or later_df.empty):
        return None
    if earlier_df is None or earlier_df.empty:
        return later_mean
    if later_df is None or later_df.empty:
        return adjusted_earlier
    return (adjusted_earlier + later_mean) / 2

uploaded_file = st.file_uploader("Upload Excel file (.xlsx/.xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # read raw bytes for caching function
        raw_bytes = uploaded_file.read()

        df = load_excel_and_unmerge(raw_bytes)

        if df.empty:
            st.error("No data read from Excel.")
        else:
            required_cols = {'Start', 'End', 'Name', 'Description', 'Demand'}
            missing = required_cols - set(df.columns)
            if missing:
                st.error(f"Missing required columns: {missing}. Columns found: {list(df.columns)}")
            else:
                # clean and robust parse of Demand
                df, diag = robust_clean_demand(df, demand_col='Demand')
                st.info(f"Demand parsed: total={diag['total']}, parsed_numbers={diag['parsed_as_number']}, parsed_missing={diag['parsed_as_nan']}")
                if diag['suspicious_count'] > 0:
                    st.warning(f"Detected {diag['suspicious_count']} suspiciously large Demand values, corrected: {diag['corrected_count']} (heuristic).")

                with st.expander("Preview data (first 100 rows)"):
                    st.dataframe(df.head(100))

                st.subheader("Filters")
                name_filter = st.text_input("Filter by Name (min 2 chars):")
                desc_filter = st.text_input("Filter by Description (min 2 chars):")
                min_start = pd.to_datetime(df['Start'], errors='coerce').min()
                max_end = pd.to_datetime(df['End'], errors='coerce').max()
                start_date = st.date_input("Earlier period start:", min_value=min_start.date() if pd.notna(min_start) else None, value=min_start.date() if pd.notna(min_start) else datetime.today().date())
                end_date = st.date_input("Earlier period end:", min_value=min_start.date() if pd.notna(min_start) else None, value=max_end.date() if pd.notna(max_end) else datetime.today().date())

                st.subheader("Later period (for comparison)")
                later_start = st.date_input("Later period start:", value=start_date, key='later_start_date')
                later_end = st.date_input("Later period end:", value=end_date, key='later_end_date')

                st.subheader("Target growth (%)")
                target_growth = st.number_input("Growth percent:", min_value=-100, max_value=1000, value=0)

                earlier_df = filter_data(df, name_filter, desc_filter, start_date, end_date)
                later_df = filter_data(df, name_filter, desc_filter, later_start, later_end)

                st.write("Earlier period sample:")
                st.dataframe(earlier_df.head(50))
                st.write("Later period sample:")
                st.dataframe(later_df.head(50))

                if st.button("Calculate estimation"):
                    if earlier_df.empty and later_df.empty:
                        st.warning("No records in selected periods.")
                    else:
                        est = estimate_demand(earlier_df, later_df, target_growth)
                        if est is None:
                            st.warning("Unable to estimate (not enough data).")
                        else:
                            st.success(f"Estimated Demand (average): {est:,.2f} EUR")

    except Exception as e:
        st.error(f"Error processing file: {e}")
