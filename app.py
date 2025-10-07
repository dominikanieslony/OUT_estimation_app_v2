import streamlit as st
import pandas as pd
import chardet
import io
from datetime import date

st.set_page_config(page_title="Campaign Estimator", page_icon="📊", layout="wide")

# ------------------ Funkcje pomocnicze ------------------
def detect_encoding_and_read(uploaded_file):
    raw = uploaded_file.read()
    enc = chardet.detect(raw).get("encoding", "utf-8")
    try:
        text = raw.decode(enc)
    except Exception:
        text = raw.decode("utf-8", errors="replace")
    uploaded_file.seek(0)

    # Wykrywanie separatora na podstawie drugiego wiersza (pierwszy może być opisem)
    lines = text.splitlines()
    if len(lines) < 2:
        sep = ","
    else:
        second_line = lines[1]
        sep = "\t" if "\t" in second_line else ","

    # Wczytanie CSV z drugiego wiersza jako header
    df = pd.read_csv(io.StringIO(text), sep=sep, decimal=",", header=1)

    # Usuń spacje wokół nazw kolumn
    df.columns = [col.strip() for col in df.columns]
    return df

def clean_demand_series(s: pd.Series) -> pd.Series:
    def parse_val(val):
        if pd.isna(val):
            return None
        v = str(val).strip()
        v = v.replace("€", "").replace("\u00A0", "").replace(" ", "")
        # usuń przecinki jako separator tysięcy
        v = v.replace(",", "")
        # zostaw tylko cyfry, kropkę i minus
        cleaned = ''.join(ch for ch in v if (ch.isdigit() or ch in ".-"))
        if cleaned in ["", ".", "-", "-."]:
            return None
        try:
            return float(cleaned)
        except Exception:
            return None
    return s.apply(parse_val)

def similar_title_mask(df, phrase):
    if not phrase:
        return pd.Series([True]*len(df), index=df.index)
    pref = phrase[:3].lower()
    name_col = df["Name"].fillna("").astype(str).str.lower()
    desc_col = df["Description"].fillna("").astype(str).str.lower()
    starts = name_col.str.startswith(pref) | desc_col.str.startswith(pref)
    if starts.any():
        return starts
    return name_col.str.contains(pref, na=False) | desc_col.str.contains(pref, na=False)

def estimate_demand(earlier_df, later_df, pct):
    earlier_mean = earlier_df["Demand"].mean() if not earlier_df.empty else 0
    later_mean = later_df["Demand"].mean() if not later_df.empty else 0
    adjusted_earlier = earlier_mean * (1 + pct/100.0)
    if earlier_df.empty and later_df.empty:
        return None
    if earlier_df.empty:
        return later_mean
    if later_df.empty:
        return adjusted_earlier
    return (adjusted_earlier + later_mean) / 2.0

# ------------------ Streamlit UI ------------------
st.title("📊 Campaign Estimator")

uploaded_file = st.file_uploader(
    "Wgraj plik CSV z kolumnami: Start, End, Country, Name, Description, Category, Demand",
    type=["csv"]
)
if not uploaded_file:
    st.info("Prześlij plik CSV, aby rozpocząć.")
    st.stop()

try:
    df = detect_encoding_and_read(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać pliku: {e}")
    st.stop()

# upewnij się, że kolumny istnieją
required_cols = {"Start", "End", "Country", "Name", "Description", "Demand"}
if not required_cols.issubset(df.columns):
    st.error(f"Brakuje wymaganych kolumn: {required_cols - set(df.columns)}")
    st.stop()

# czyszczenie danych
df["Start"] = pd.to_datetime(df["Start"], dayfirst=True, errors="coerce")
df["End"] = pd.to_datetime(df["End"], dayfirst=True, errors="coerce")
df["Demand"] = clean_demand_series(df["Demand"])

# wybór kraju
countries = sorted(df["Country"].dropna().astype(str).unique().tolist())
selected_country = st.selectbox("🌍 Wybierz kraj:", countries)

# kategoria (opcjonalnie)
categories = df["Category"].dropna().unique().tolist() if "Category" in df.columns else []
if categories:
    categories = sorted(categories)
    selected_category = st.selectbox("🏷️ Wybierz kategorię:", ["All"] + categories)
else:
    selected_category = "All"

# fraza
campaign_filter = st.text_input("🔎 Nazwa kampanii (fraza, min. 3 znaki):")

# okresy
st.subheader("⏳ Earlier Period")
earlier_start_date = st.date_input("Start date (Earlier Period):", key="earlier_start")
earlier_end_date = st.date_input("End date (Earlier Period):", key="earlier_end")

st.subheader("📈 Target growth from Earlier Period (%)")
target_growth = st.number_input(
    "Enter growth percentage (can be negative):",
    min_value=-100, max_value=1000, step=1, format="%d"
)

st.subheader("⏳ Later Period")
later_start_date = st.date_input("Start date (Later Period):", key="later_start")
later_end_date = st.date_input("End date (Later Period):", key="later_end")

# filtracja danych
df_filtered = df[df["Country"] == selected_country].copy()

if selected_category != "All" and "Category" in df.columns:
    df_filtered = df_filtered[df_filtered["Category"].astype(str).str.lower() == selected_category.lower()]

if campaign_filter and len(campaign_filter) >= 3:
    df_filtered = df_filtered[similar_title_mask(df_filtered, campaign_filter)]

earlier_filtered = df_filtered[
    (df_filtered["Start"] >= pd.to_datetime(earlier_start_date)) &
    (df_filtered["End"] <= pd.to_datetime(earlier_end_date))
]

later_filtered = df_filtered[
    (df_filtered["Start"] >= pd.to_datetime(later_start_date)) &
    (df_filtered["End"] <= pd.to_datetime(later_end_date))
]

# wybór kampanii
st.subheader("Select campaigns to include from Earlier Period:")
earlier_selections = {
    idx: st.checkbox(
        f"{row['Name']} | {row['Description']} | Start: {row['Start'].date() if pd.notna(row['Start']) else 'n/a'} | "
        f"End: {row['End'].date() if pd.notna(row['End']) else 'n/a'} | Demand: {row['Demand']}",
        value=True, key=f"earlier_{idx}"
    )
    for idx, row in earlier_filtered.iterrows()
}

st.subheader("Select campaigns to include from Later Period:")
later_selections = {
    idx: st.checkbox(
        f"{row['Name']} | {row['Description']} | Start: {row['Start'].date() if pd.notna(row['Start']) else 'n/a'} | "
        f"End: {row['End'].date() if pd.notna(row['End']) else 'n/a'} | Demand: {row['Demand']}",
        value=True, key=f"later_{idx}"
    )
    for idx, row in later_filtered.iterrows()
}

earlier_selected_df = earlier_filtered.loc[[idx for idx, checked in earlier_selections.items() if checked]]
later_selected_df = later_filtered.loc[[idx for idx, checked in later_selections.items() if checked]]

# obliczanie estymacji
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
                file_name="campaign_estimation_data.csv",
                mime="text/csv"
            )
