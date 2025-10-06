import streamlit as st
import pandas as pd
import chardet
from datetime import datetime

def load_data(uploaded_file):
    raw = uploaded_file.read()
    encoding = chardet.detect(raw)['encoding']
    uploaded_file.seek(0)
    
    # Try different separators since your file uses semicolon
    try:
        df = pd.read_csv(uploaded_file, encoding=encoding, sep=';', dayfirst=True)
    except:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, encoding=encoding, sep='\t', dayfirst=True)
    
    return df

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
    
    # Map the column names from your new CSV
    if 'Demand' in df.columns:
        df['Demand'] = df['Demand'].apply(parse_demand)
    return df

def map_column_names(df):
    """Map the column names from your new CSV to the expected names"""
    column_mapping = {
        'Start': 'Date Start',
        'End': 'Date End', 
        'Name': 'Campaign name',
        'Description': 'Description',
        'Country': 'Country',
        'Category': 'Category_name',
        'Demand': 'Demand'
    }
    
    # Rename columns that exist in the dataframe
    for old_name, new_name in column_mapping.items():
        if old_name in df.columns and new_name not in df.columns:
            df[new_name] = df[old_name]
    
    return df

def filter_data(df, country, campaign_filter, start_date, end_date, selected_category=None):
    df_filtered = df[df['Country'] == country].copy()

    if selected_category and selected_category != "All":
        df_filtered = df_filtered[df_filtered['Category_name'].str.strip().str.lower() == selected_category.strip().lower()]

    if campaign_filter and len(campaign_filter) >= 3:
        mask_desc = df_filtered['Description'].str.contains(campaign_filter, case=False, na=False)
        # Use 'Campaign name' if available, otherwise use 'Name'
        campaign_col = 'Campaign name' if 'Campaign name' in df_filtered.columns else 'Name'
        mask_camp = df_filtered[campaign_col].str.contains(campaign_filter, case=False, na=False)
        df_filtered = df_filtered[mask_desc | mask_camp]

    # Handle date columns - try different possible column names
    date_start_col = 'Date Start' if 'Date Start' in df_filtered.columns else 'Start'
    date_end_col = 'Date End' if 'Date End' in df_filtered.columns else 'End'
    
    df_filtered[date_start_col] = pd.to_datetime(df_filtered[date_start_col], dayfirst=True, errors='coerce')
    df_filtered[date_end_col] = pd.to_datetime(df_filtered[date_end_col], dayfirst=True, errors='coerce')

    df_filtered = df_filtered[
        (df_filtered[date_start_col] >= pd.to_datetime(start_date)) &
        (df_filtered[date_end_col] <= pd.to_datetime(end_date))
    ]

    return df_filtered

def estimate_demand(earlier_df, later_df, percentage):
    earlier_mean = earlier_df['Demand'].mean() if not earlier_df.empty and 'Demand' in earlier_df.columns else 0
    later_mean = later_df['Demand'].mean() if not later_df.empty and 'Demand' in later_df.columns else 0
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
    # Create a list of desired columns in order, keeping only those that exist
    desired_order = ['Campaign name', 'Description', 'Date Start', 'Date End', 'Country', 'Category_name', 'Demand']
    existing_cols = [col for col in desired_order if col in df.columns]
    
    # Add any remaining columns
    remaining_cols = [col for col in df.columns if col not in existing_cols]
    final_order = existing_cols + remaining_cols
    
    return df[final_order]

st.title("📊 Campaign Estimator")

uploaded_file = st.file_uploader("Upload campaign data CSV file", type="csv")

if uploaded_file:
    try:
        df = load_data(uploaded_file)
        
        # Map column names to expected format
        df = map_column_names(df)
        
        # Check for required columns
        required_cols = {'Country', 'Description', 'Demand'}
        date_cols = {'Date Start', 'Date End'} | {'Start', 'End'}
        name_cols = {'Campaign name'} | {'Name'}
        
        if not required_cols.issubset(df.columns):
            st.error(f"❌ Missing required columns. Found: {list(df.columns)}")
            st.info("Expected columns: Country, Description, Demand, and date columns")
        else:
            df = clean_demand_column(df)

            # Display basic info about the data
            st.sidebar.subheader("Data Overview")
            st.sidebar.write(f"Total rows: {len(df)}")
            st.sidebar.write(f"Columns: {list(df.columns)}")
            st.sidebar.write(f"Date range: {df['Date Start'].min() if 'Date Start' in df.columns else df['Start'].min()} to {df['Date End'].max() if 'Date End' in df.columns else df['End'].max()}")

            country_list = df['Country'].dropna().unique().tolist()
            selected_country = st.selectbox("🌍 Select country:", country_list)

            # Handle category column - use 'Category_name' or 'Category'
            category_col = 'Category_name' if 'Category_name' in df.columns else 'Category'
            categories = df[category_col].dropna().unique().tolist()
            categories = sorted(categories)
            selected_category = st.selectbox("🏷️ Select category:", ["All"] + categories)

            campaign_filter = st.text_input("🔎 Filter campaigns (contains, min 3 letters):")

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

            earlier_filtered = filter_data(df, selected_country, campaign_filter, earlier_start_date, earlier_end_date, selected_category)
            later_filtered = filter_data(df, selected_country, campaign_filter, later_start_date, later_end_date, selected_category)

            earlier_filtered = reorder_columns(earlier_filtered)
            later_filtered = reorder_columns(later_filtered)

            st.subheader("Select campaigns to include from Earlier Period:")
            if not earlier_filtered.empty:
                earlier_selections = {
                    idx: st.checkbox(
                        f"{row['Campaign name'] if 'Campaign name' in row else row['Name']} | {row['Description']} | Start: {row['Date Start'].date() if 'Date Start' in row else row['Start'].date()} | End: {row['Date End'].date() if 'Date End' in row else row['End'].date()} | Demand: {row['Demand'] if 'Demand' in row else 'N/A'}",
                        value=True, key=f"earlier_{idx}"
                    )
                    for idx, row in earlier_filtered.iterrows()
                }
                earlier_selected_df = earlier_filtered.loc[[idx for idx, checked in earlier_selections.items() if checked]]
            else:
                st.info("No campaigns found in the earlier period with the selected filters.")
                earlier_selected_df = pd.DataFrame()

            st.subheader("Select campaigns to include from Later Period:")
            if not later_filtered.empty:
                later_selections = {
                    idx: st.checkbox(
                        f"{row['Campaign name'] if 'Campaign name' in row else row['Name']} | {row['Description']} | Start: {row['Date Start'].date() if 'Date Start' in row else row['Start'].date()} | End: {row['Date End'].date() if 'Date End' in row else row['End'].date()} | Demand: {row['Demand'] if 'Demand' in row else 'N/A'}",
                        value=True, key=f"later_{idx}"
                    )
                    for idx, row in later_filtered.iterrows()
                }
                later_selected_df = later_filtered.loc[[idx for idx, checked in later_selections.items() if checked]]
            else:
                st.info("No campaigns found in the later period with the selected filters.")
                later_selected_df = pd.DataFrame()

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

                        if not earlier_selected_df.empty or not later_selected_df.empty:
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
        st.info("Please check that your CSV file has the correct format with columns like: Start, End, Country, Name, Description, Category, Demand")