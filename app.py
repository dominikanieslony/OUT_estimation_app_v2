def clean_demand_column(df):
    import re

    def parse_demand(val):
        if pd.isna(val):
            return None
        val = str(val)

        # usuń spacje zwykłe i niełamliwe (U+00A0)
        val = val.replace('\u00A0', '').replace(' ', '')

        # usuń waluty i znaki nieliczbowe
        val = re.sub(r'[^\d,.-]', '', val)

        # zamień przecinek na kropkę (jeśli jest separatorem dziesiętnym)
        # ale tylko jeśli nie występują dwa przecinki lub dwie kropki
        if val.count(',') == 1 and val.count('.') == 0:
            val = val.replace(',', '.')

        # czasem Excel zapisze "54 332" jako "54332." lub "54332,00"
        try:
            return float(val)
        except ValueError:
            return None

    if 'Demand' in df.columns:
        df['Demand'] = df['Demand'].apply(parse_demand)

        # usuń absurdalne wartości (np. > 1e9)
        df.loc[df['Demand'] > 1e9, 'Demand'] = None

    return df
