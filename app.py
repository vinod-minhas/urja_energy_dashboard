import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import os
import glob

# ========================================================================
# PAGE CONFIG & STYLING
# ========================================================================
st.set_page_config(page_title="RME Energy Command Center", page_icon="⚡", layout="wide")

# ========================================================================
# CONSTANTS
# ========================================================================
SITES = ['DEL5', 'HYD1', 'BLR4', 'BOM5', 'MAA3', 'CCU2', 'PNQ1', 'AMD1',
         'JAI1', 'LKO1', 'COK1', 'IDR1', 'NAG1', 'VGA1', 'GHY1', 'PAT1',
         'RAN1', 'BHU1', 'LUH1', 'KOL3', 'GUR2']

COLORS = {
    'eb': '#2E86C1',
    'dg': '#E74C3C',
    'solar': '#27AE60',
    'target': '#F39C12',
    'water': '#3498DB',
    'fuel': '#D35400',
    'scrap': '#8E44AD'
}

EUI_TARGETS = {
    1: 0.26, 2: 0.32, 3: 0.46, 4: 0.67, 5: 0.72, 6: 0.76,
    7: 0.80, 8: 0.84, 9: 0.74, 10: 0.68, 11: 0.52, 12: 0.48
}

WORKDOCS_SYNC_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")


# ========================================================================
# HELPER FUNCTIONS
# ========================================================================
def get_layout(title="", height=400, showlegend=True, xaxis_title="", yaxis_title="", margin=None):
    layout = dict(
        title=dict(text=title, font=dict(size=14)),
        height=height,
        showlegend=showlegend,
        template="plotly_white",
        paper_bgcolor='#FFFFFF',
        plot_bgcolor='#F8F9FA',
        margin=margin if margin else dict(l=40, r=40, t=50, b=40),
        xaxis_title=xaxis_title,
        yaxis_title=yaxis_title
    )
    return layout


def get_files_hash(path):
    """Get combined modification time of all Excel files as cache key."""
    files = glob.glob(os.path.join(path, "**", "*.xlsx"), recursive=True)
    mod_times = []
    for f in files:
        if not os.path.basename(f).startswith('~'):
            mod_times.append(os.path.getmtime(f))
    return sum(mod_times) if mod_times else 0

# ========================================================================
# CSS STYLING
# ========================================================================
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        padding: 1.5rem;
        border-radius: 10px;
        margin-bottom: 1rem;
        color: white;
        text-align: center;
    }
    .main-header h1 { font-size: 1.8rem; margin: 0; }
    .main-header p { font-size: 0.9rem; opacity: 0.8; margin: 0.5rem 0 0 0; }
    .section-header {
        font-size: 1.1rem;
        font-weight: 600;
        margin: 1rem 0 0.5rem 0;
        padding: 0.5rem;
        background: #f0f2f6;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

# ========================================================================
# DATA LOADING FUNCTIONS
# ========================================================================
@st.cache_data
def load_all_site_data(path, _file_hash):
    """Load all site data from Excel files in subfolders."""
    all_energy = []
    all_feeders = []
    all_wdr = []
    all_submeters = []
    all_stp = []
    all_water = []
    all_fuel = []
    all_scrap = []
    site_info_list = []

    # Search recursively in subfolders
    files = glob.glob(os.path.join(path, "**", "*_Consumption_Template.xlsx"), recursive=True)
    for file in files:
        filename = os.path.basename(file)
        if filename.startswith('~'):
            continue
        site_code = filename.split('_')[0].upper()

        try:
            xl = pd.ExcelFile(file)
            sheets = xl.sheet_names

            # ---- ENERGY - MASTER ----
            if 'Energy - Master' in sheets:
                energy_df = pd.read_excel(xl, 'Energy - Master', skiprows=1)
                energy_df.columns = range(len(energy_df.columns))
                rename_map = {
                    0: 'Date',
                    1: 'EB_kWh',
                    2: 'DG_kWh',
                    3: 'Solar_kWh',
                    4: 'Total_Power_kWh',
                    5: 'HSD_Ltr',
                    6: 'EB_MTD',
                    7: 'DG_MTD',
                    8: 'Solar_MTD',
                    9: 'Total_Power_MTD',
                    10: 'Budget_Unit',
                    11: 'Budget_MTD',
                    12: 'EUI_Target',
                    13: 'EUI_Actual',
                    14: 'EUI_MTD',
                    15: 'Raw_Water_KL',
                    16: 'STP_Inlet_KL',
                    17: 'STP_Outlet_KL',
                    18: 'Total_Water_KL',
                    19: 'Comments'
                }
                energy_df.rename(columns=rename_map, inplace=True)
                energy_df['Site'] = site_code

                if 'Date' in energy_df.columns:
                    energy_df['Date'] = pd.to_datetime(energy_df['Date'], dayfirst=True, errors='coerce')
                    energy_df = energy_df.dropna(subset=['Date'])

                # Convert numeric columns
                numeric_cols = ['EB_kWh', 'DG_kWh', 'Solar_kWh', 'Total_Power_kWh', 'HSD_Ltr',
                               'EUI_Target', 'EUI_Actual', 'EUI_MTD', 'Raw_Water_KL', 'STP_Inlet_KL', 'STP_Outlet_KL',
                               'Total_Water_KL', 'Budget_Unit', 'Budget_MTD', 'EB_MTD', 'DG_MTD',
                               'Solar_MTD', 'Total_Power_MTD']
                for col in numeric_cols:
                    if col in energy_df.columns:
                        energy_df[col] = pd.to_numeric(energy_df[col], errors='coerce')

                # Drop rows where all main consumption values are NaN
                main_cols = ['EB_kWh', 'DG_kWh', 'Solar_kWh', 'Total_Power_kWh']
                existing = [c for c in main_cols if c in energy_df.columns]
                if existing:
                    energy_df = energy_df.dropna(subset=existing, how='all')

                # Add WDR calculation
                if 'STP_Inlet_KL' in energy_df.columns and 'STP_Outlet_KL' in energy_df.columns:
                    energy_df['WDR_Percent'] = np.where(
                        energy_df['STP_Inlet_KL'] > 0,
                        (energy_df['STP_Outlet_KL'] / energy_df['STP_Inlet_KL']) * 100,
                        0
                    )
                else:
                    energy_df['WDR_Percent'] = 0

                # Add Treated Water
                if 'STP_Outlet_KL' in energy_df.columns:
                    energy_df['Treated_Water_KL'] = energy_df['STP_Outlet_KL']
                else:
                    energy_df['Treated_Water_KL'] = 0

                all_energy.append(energy_df)

            # ---- LT PANEL FEEDER READINGS ----
            if 'LT Panel Feeder Readings' in sheets:
                try:
                    feeders_raw = pd.read_excel(xl, 'LT Panel Feeder Readings', header=None, skiprows=3)
                    feeder_names_row = pd.read_excel(xl, 'LT Panel Feeder Readings', header=None, skiprows=1, nrows=1)

                    feeder_data = pd.DataFrame()
                    feeder_data['Date'] = pd.to_datetime(feeders_raw.iloc[:, 0], errors='coerce')

                    feeder_names = feeder_names_row.iloc[0, 2:].dropna().tolist()
                    consumption_cols = list(range(4, feeders_raw.shape[1], 3))

                    for i, col_idx in enumerate(consumption_cols):
                        if col_idx < feeders_raw.shape[1] and i < len(feeder_names):
                            name = str(feeder_names[i]).replace(chr(10), ' ').strip()
                            feeder_data[name] = pd.to_numeric(feeders_raw.iloc[:, col_idx], errors='coerce')

                    feeder_data['Site'] = site_code
                    feeder_data = feeder_data.dropna(subset=['Date'])
                    all_feeders.append(feeder_data)
                except Exception:
                    pass

            # ---- SUBMETER READINGS ----
            if 'Submeter Readings' in sheets:
                try:
                    sub_raw = pd.read_excel(xl, 'Submeter Readings', header=None, skiprows=3)
                    sub_names_row = pd.read_excel(xl, 'Submeter Readings', header=None, skiprows=1, nrows=1)

                    sub_data = pd.DataFrame()
                    sub_data['Date'] = pd.to_datetime(sub_raw.iloc[:, 0], errors='coerce')

                    sub_names = sub_names_row.iloc[0, 2:].dropna().tolist()
                    consumption_cols = list(range(4, sub_raw.shape[1], 3))

                    for i, col_idx in enumerate(consumption_cols):
                        if col_idx < sub_raw.shape[1] and i < len(sub_names):
                            name = str(sub_names[i]).replace(chr(10), ' ').strip()
                            sub_data[name] = pd.to_numeric(sub_raw.iloc[:, col_idx], errors='coerce')

                    sub_data['Site'] = site_code
                    sub_data = sub_data.dropna(subset=['Date'])
                    all_submeters.append(sub_data)
                except Exception:
                    pass

            # ---- FUEL STATUS ----
            if 'Fuel Status' in sheets:
                try:
                    fuel_raw = pd.read_excel(xl, 'Fuel Status', header=None, skiprows=3)
                    fuel_data = pd.DataFrame()
                    fuel_data['Date'] = pd.to_datetime(fuel_raw.iloc[:, 0], errors='coerce')
                    fuel_data['DG1_Consumption'] = pd.to_numeric(fuel_raw.iloc[:, 4], errors='coerce').fillna(0)
                    fuel_data['DG2_Consumption'] = pd.to_numeric(fuel_raw.iloc[:, 9], errors='coerce').fillna(0)
                    if fuel_raw.shape[1] > 14:
                        fuel_data['DG3_Consumption'] = pd.to_numeric(fuel_raw.iloc[:, 14], errors='coerce').fillna(0)
                    else:
                        fuel_data['DG3_Consumption'] = 0
                    fuel_data['Total_HSD'] = fuel_data['DG1_Consumption'] + fuel_data['DG2_Consumption'] + fuel_data['DG3_Consumption']
                    if fuel_raw.shape[1] > 28:
                        fuel_data['Total_Stock'] = pd.to_numeric(fuel_raw.iloc[:, 28], errors='coerce').fillna(0)
                    fuel_data['Site'] = site_code
                    fuel_data = fuel_data.dropna(subset=['Date'])
                    all_fuel.append(fuel_data)
                except Exception:
                    pass
            # ---- RAW WATER ----
            if 'Raw Water' in sheets:
                try:
                    rw_raw = pd.read_excel(xl, 'Raw Water', header=None, skiprows=3)
                    rw_data = pd.DataFrame()
                    rw_data['Date'] = pd.to_datetime(rw_raw.iloc[:, 0], errors='coerce')
                    rw_data['Raw_Water_KL'] = pd.to_numeric(rw_raw.iloc[:, 20], errors='coerce')
                    rw_data['Borewell_KL'] = pd.to_numeric(rw_raw.iloc[:, 4], errors='coerce')
                    rw_data['Site'] = site_code
                    rw_data = rw_data.dropna(subset=['Date'])
                    all_water.append(rw_data)
                except Exception:
                    pass

            # ---- STP ----
            if 'STP' in sheets:
                try:
                    stp_raw = pd.read_excel(xl, 'STP', header=None, skiprows=3)
                    stp_data = pd.DataFrame()
                    stp_data['Date'] = pd.to_datetime(stp_raw.iloc[:, 0], errors='coerce')
                    stp_data['STP_Inlet_KL'] = pd.to_numeric(stp_raw.iloc[:, 4], errors='coerce')
                    stp_data['Flushing_Ltrs'] = pd.to_numeric(stp_raw.iloc[:, 7], errors='coerce')
                    stp_data['Horticulture_Ltrs'] = pd.to_numeric(stp_raw.iloc[:, 10], errors='coerce')
                    stp_data['STP_Outlet_KL'] = pd.to_numeric(stp_raw.iloc[:, 21], errors='coerce')
                    stp_data['Site'] = site_code
                    stp_data = stp_data.dropna(subset=['Date'])
                    all_stp.append(stp_data)
                except Exception:
                    pass

            # ---- SCRAP CONSUMPTION ----
            if 'Scrap Consumption' in sheets:
                try:
                    scrap_raw = pd.read_excel(xl, 'Scrap Consumption', skiprows=1)
                    if 'Category' in scrap_raw.columns:
                        id_cols = ['Category', 'Description', 'Unit']
                        date_cols = [c for c in scrap_raw.columns if c not in id_cols]
                        scrap_long = scrap_raw.melt(id_vars=id_cols, value_vars=date_cols,
                                                   var_name='Date', value_name='Quantity')
                        scrap_long['Date'] = pd.to_datetime(scrap_long['Date'], errors='coerce')
                        scrap_long['Quantity'] = pd.to_numeric(scrap_long['Quantity'], errors='coerce')
                        scrap_long['Site'] = site_code
                        scrap_long = scrap_long.dropna(subset=['Date', 'Quantity'])
                        scrap_long = scrap_long[scrap_long['Quantity'] != 0]
                        all_scrap.append(scrap_long)
                except Exception:
                    pass
                # ---- WDR % from Scrap ----
                try:
                    scrap_raw2 = pd.read_excel(xl, 'Scrap Consumption', header=None)
                    dates_row = scrap_raw2.iloc[1, 3:]
                    wdr_row = scrap_raw2.iloc[90, 3:]
                    wdr_df = pd.DataFrame({
                        'Date': pd.to_datetime(dates_row.values, errors='coerce'),
                        'WDR_Pct': pd.to_numeric(wdr_row.values, errors='coerce')
                    })
                    wdr_df = wdr_df.dropna(subset=['Date'])
                    wdr_df['WDR_Pct'] = wdr_df['WDR_Pct'] * 100
                    wdr_df['Site'] = site_code
                    all_wdr.append(wdr_df)
                except Exception:
                    pass

            # ---- SITE INFORMATION ----
            if 'Site Information' in sheets:
                try:
                    info = pd.read_excel(xl, 'Site Information', header=None)
                    gsf_row = info[info.iloc[:, 0].astype(str).str.contains('Capacity', na=False)]
                    if not gsf_row.empty:
                        gsf = gsf_row.iloc[0, 1]
                        site_info_list.append({'Site': site_code, 'GSF': gsf})
                except Exception:
                    pass

        except Exception as e:
            st.warning(f"Error loading {filename}: {e}")
    if all_energy:
        energy_combined = pd.concat(all_energy, ignore_index=True)
        if all_water:
            water_combined = pd.concat(all_water, ignore_index=True)
            energy_combined = energy_combined.merge(
                water_combined[['Date', 'Site', 'Raw_Water_KL', 'Borewell_KL']],
                on=['Date', 'Site'], how='left', suffixes=('_drop', '')
            )
            drop_cols = [c for c in energy_combined.columns if c.endswith('_drop')]
            energy_combined.drop(columns=drop_cols, inplace=True)
        if all_stp:
            stp_combined = pd.concat(all_stp, ignore_index=True)
            energy_combined = energy_combined.merge(
                stp_combined[['Date', 'Site', 'STP_Inlet_KL', 'Flushing_Ltrs', 'Horticulture_Ltrs', 'STP_Outlet_KL']],
                on=['Date', 'Site'], how='left', suffixes=('_drop', '')
            )
            drop_cols = [c for c in energy_combined.columns if c.endswith('_drop')]
            energy_combined.drop(columns=drop_cols, inplace=True)

    # Merge WDR after all sites are combined
    final_energy = pd.concat(all_energy, ignore_index=True) if all_energy else pd.DataFrame()
    if all_wdr and not final_energy.empty:
        wdr_combined = pd.concat(all_wdr, ignore_index=True)
        final_energy = final_energy.merge(
            wdr_combined[['Date', 'Site', 'WDR_Pct']],
            on=['Date', 'Site'], how='left', suffixes=('_drop', '')
        )
        drop_cols = [c for c in final_energy.columns if c.endswith('_drop')]
        final_energy.drop(columns=drop_cols, inplace=True)
    import streamlit as st
    st.write(f"DEBUG WDR: final_energy shape={final_energy.shape}")
    st.write(f"DEBUG WDR: WDR_Pct in columns = {'WDR_Pct' in final_energy.columns}")
    if 'WDR_Pct' in final_energy.columns:
        st.write(f"DEBUG WDR: non-null WDR = {final_energy['WDR_Pct'].notna().sum()}")
        st.write(f"DEBUG WDR: sites with WDR = {final_energy[final_energy['WDR_Pct'].notna()]['Site'].unique().tolist()}")


    return {
        'energy': final_energy,
        'feeders': pd.concat(all_feeders, ignore_index=True) if all_feeders else pd.DataFrame(),
        'submeters': pd.concat(all_submeters, ignore_index=True) if all_submeters else pd.DataFrame(),
        'fuel': pd.concat(all_fuel, ignore_index=True) if all_fuel else pd.DataFrame(),
        'scrap': pd.concat(all_scrap, ignore_index=True) if all_scrap else pd.DataFrame(),
        'site_info': site_info_list
    }


def generate_demo_data():
    """Generate demo data for testing when WorkDocs is not available."""
    dates = pd.date_range(start='2026-01-01', end='2026-07-10', freq='D')
    records = []
    for date in dates:
        for site in SITES[:5]:
            base_eb = np.random.randint(8000, 25000)
            season = 1 + 0.5 * np.sin((date.month - 1) * np.pi / 6)
            eb = int(base_eb * season + np.random.randint(-2000, 2000))
            dg = int(np.random.choice([0]*7 + [np.random.randint(100, 3000)]*3))
            base_solar = np.random.randint(2000, 6000)
            solar = int(base_solar * (0.8 + 0.4 * np.sin((date.month - 1) * np.pi / 6)) + np.random.randint(-500, 500))
            solar = max(0, solar)
            total = eb + dg + solar
            hsd = int(dg * 0.35) if dg > 0 else 0
            gsf = 1194597
            eui = total / gsf
            budget = EUI_TARGETS.get(date.month, 0.50)
            base_water = np.random.uniform(40, 80)
            water = base_water + np.random.uniform(-8, 8)
            stp_inlet = water * 0.8
            stp_outlet = stp_inlet * np.random.uniform(0.75, 0.90)
            wdr = (stp_outlet / stp_inlet * 100) if stp_inlet > 0 else 0

            records.append({
                'Date': date, 'Site': site, 'EB_kWh': eb, 'DG_kWh': dg,
                'Solar_kWh': solar, 'Total_Power_kWh': total, 'HSD_Ltr': hsd,
                'EUI_Day': eui, 'Budget_Unit': budget,
                'Raw_Water_KL': water, 'Treated_Water_KL': stp_outlet,
                'WDR_Percent': wdr, 'STP_Inlet_KL': stp_inlet,
                'STP_Outlet_KL': stp_outlet
            })

    return pd.DataFrame(records)


# ========================================================================
# LOAD DATA
# ========================================================================
if os.path.exists(WORKDOCS_SYNC_PATH):
    file_hash = get_files_hash(WORKDOCS_SYNC_PATH)
    data = load_all_site_data(WORKDOCS_SYNC_PATH, file_hash)
    df = data['energy']
    feeders_df = data['feeders']
    submeters_df = data['submeters']
    fuel_df = data['fuel']
    scrap_df = data['scrap']
    site_info_list = data['site_info']
    data_source = "WorkDocs"
else:
    df = generate_demo_data()
    feeders_df = pd.DataFrame()
    submeters_df = pd.DataFrame()
    fuel_df = pd.DataFrame()
    scrap_df = pd.DataFrame()
    site_info_list = []
    data_source = "Demo"

# ========================================================================
# SIDEBAR
# ========================================================================
with st.sidebar:
    st.markdown("### ⚡ URJA Command Center")
    st.caption(f"📡 {data_source} | {len(df['Site'].unique()) if not df.empty else 0} Sites")
    st.markdown("---")

    all_sites = sorted(df['Site'].unique().tolist()) if not df.empty else SITES
    select_all = st.checkbox("🌐 All Sites", value=True)
    if select_all:
        selected_sites = all_sites
    else:
        selected_sites = st.multiselect("Choose Sites", all_sites, default=all_sites[:3])

    st.markdown("---")

    if not df.empty:
        if 'Date' not in df.columns:
            date_col = [c for c in df.columns if 'date' in str(c).lower() or 'dd' in str(c).lower()]
            if date_col:
                df.rename(columns={date_col: 'Date'}, inplace=True)
                df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
        max_date = df['Date'].max().date()
        min_date = df['Date'].min().date()
    else:
        max_date = datetime(2026, 7, 10).date()
        min_date = datetime(2026, 1, 1).date()


    date_opt = st.radio("📅 Period",
                        ["Today", "7 Days", "30 Days", "This Month", "All Data", "Custom"],
                        index=4)

    if date_opt == "Today":
        start_date, end_date = max_date, max_date
    elif date_opt == "7 Days":
        start_date, end_date = max_date - timedelta(days=7), max_date
    elif date_opt == "30 Days":
        start_date, end_date = max_date - timedelta(days=30), max_date
    elif date_opt == "This Month":
        start_date, end_date = max_date.replace(day=1), max_date
    elif date_opt == "All Data":
        start_date, end_date = min_date, max_date
    else:
        start_date = st.date_input("From", min_date)
        end_date = st.date_input("To", max_date)

    st.markdown("---")
    st.success(f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    if st.button("🔄 Refresh Now", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ========================================================================
# APPLY FILTERS
# ========================================================================
if not df.empty:
    filtered = df[
        (df['Site'].isin(selected_sites)) &
        (df['Date'].dt.date >= start_date) &
        (df['Date'].dt.date <= end_date)
    ]
else:
    filtered = pd.DataFrame()

# ========================================================================
# HEADER
# ========================================================================
st.markdown(f"""
<div class="main-header">
    <h1>⚡ URJA ENERGY COMMAND CENTER</h1>
    <p>🏢 {len(selected_sites)} Sites | 📅 {start_date.strftime('%d/%m/%Y')} → {end_date.strftime('%d/%m/%Y')} |
    🔄 Auto-refresh: Hourly | 📡 Source: {data_source}</p>
</div>
""", unsafe_allow_html=True)

if data_source == "Demo":
    st.info("⚠️ **Demo Mode** — WorkDocs path not found. Showing sample data. "
           "Update WORKDOCS_SYNC_PATH in code to connect live data.")

# ========================================================================
# MAIN DASHBOARD — 7 TABS
# ========================================================================
if not filtered.empty:

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "⚡ Power & EUI", "🔌 Feeder Analysis", "💧 Water & STP",
        "⛽ Fuel/DG", "♻️ Scrap & Waste", "🔍 Site Deep Dive", "🏥 Health"
    ])

    # ==================================================================
    # TAB 1: POWER & EUI (UPDATED - Option 1: Sum of daily values)
    # ==================================================================
    with tab1:
        # --- KPI CARDS: Show SUM for entire selected period ---
        eb_total = filtered['EB_kWh'].sum()
        dg_total = filtered['DG_kWh'].sum()
        solar_total = filtered['Solar_kWh'].sum()
        total_power = eb_total + dg_total + solar_total
        hsd_total = filtered['HSD_Ltr'].sum()
        if 'EUI_Actual' in filtered.columns:
            eui_data = filtered[['Date', 'EUI_Actual']].dropna(subset=['EUI_Actual'])
            if not eui_data.empty:
                eui_data['Month'] = eui_data['Date'].dt.to_period('M')
                eui_monthly_avg = eui_data.groupby('Month')['EUI_Actual'].mean()
                eui_avg = eui_monthly_avg.sum()
            else:
                eui_avg = 0
        else:
            eui_avg = 0

        # Calculate period-over-period change (compare with previous equal period)
        period_days = (end_date - start_date).days + 1
        prev_start = start_date - timedelta(days=period_days)
        prev_end = start_date - timedelta(days=1)
        prev_data = df[
            (df['Site'].isin(selected_sites)) &
            (df['Date'].dt.date >= prev_start) &
            (df['Date'].dt.date <= prev_end)
        ]

        prev_eb = prev_data['EB_kWh'].sum() if not prev_data.empty else 0
        prev_dg = prev_data['DG_kWh'].sum() if not prev_data.empty else 0
        prev_solar = prev_data['Solar_kWh'].sum() if not prev_data.empty else 0
        prev_hsd = prev_data['HSD_Ltr'].sum() if not prev_data.empty else 0

        # EUI Target for current month
        latest_date = filtered['Date'].max()
        if 'EUI_Target' in filtered.columns:
            tgt_data = filtered[['Date', 'EUI_Target']].dropna(subset=['EUI_Target'])
            if not tgt_data.empty:
                tgt_data['Month'] = tgt_data['Date'].dt.to_period('M')
                tgt_monthly_avg = tgt_data.groupby('Month')['EUI_Target'].mean()
                eui_target = tgt_monthly_avg.sum()
            else:
                eui_target = EUI_TARGETS.get(latest_date.month, 0.50)
        else:
            eui_target = EUI_TARGETS.get(latest_date.month, 0.50)



        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            delta_eb = eb_total - prev_eb if prev_eb > 0 else None
            st.metric("⚡ EB (kWh)", f"{eb_total:,.0f}",
                     f"{delta_eb:+,.0f}" if delta_eb else None, delta_color="inverse")
        with c2:
            delta_dg = dg_total - prev_dg if prev_dg > 0 else None
            st.metric("🔋 DG (kWh)", f"{dg_total:,.0f}",
                     f"{delta_dg:+,.0f}" if delta_dg else None, delta_color="inverse")
        with c3:
            delta_solar = solar_total - prev_solar if prev_solar > 0 else None
            st.metric("☀️ Solar (kWh)", f"{solar_total:,.0f}",
                     f"{delta_solar:+,.0f}" if delta_solar else None)
        with c4:
            st.metric("🔌 Total Power", f"{total_power:,.0f}")
        with c5:
            delta_hsd = hsd_total - prev_hsd if prev_hsd > 0 else None
            st.metric("⛽ HSD (Ltr)", f"{hsd_total:,.0f}",
                     f"{delta_hsd:+,.0f}" if delta_hsd else None, delta_color="inverse")
        with c6:
            st.metric("EUI (Actual)", f"{eui_avg:.3f}", f"Target: {eui_target:.2f}")

        st.markdown("---")

        # --- Power Consumption Chart + Energy Mix ---
        col_left, col_right = st.columns([3, 1])
        with col_left:
            daily = filtered.groupby('Date').agg({
                'EB_kWh': 'sum', 'DG_kWh': 'sum', 'Solar_kWh': 'sum'
            }).reset_index()
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=daily['Date'], y=daily['EB_kWh'], name='EB (Grid)',
                                    fill='tonexty', line=dict(color=COLORS['eb'], width=2), stackgroup='power'))
            fig.add_trace(go.Scatter(x=daily['Date'], y=daily['DG_kWh'], name='DG (Diesel)',
                                    fill='tonexty', line=dict(color=COLORS['dg'], width=2), stackgroup='power'))
            fig.add_trace(go.Scatter(x=daily['Date'], y=daily['Solar_kWh'], name='Solar',
                                    fill='tonexty', line=dict(color=COLORS['solar'], width=2), stackgroup='power'))
            fig.update_layout(**get_layout(title="📈 Daily Power Consumption (Stacked)", height=380))
            st.plotly_chart(fig, use_container_width=True)

        with col_right:
            # Energy Mix for ENTIRE selected period (not just last day)
            total_all = eb_total + dg_total + solar_total
            fig = go.Figure(data=[go.Pie(
                labels=['EB', 'DG', 'Solar'], values=[eb_total, dg_total, solar_total],
                marker_colors=[COLORS['eb'], COLORS['dg'], COLORS['solar']],
                hole=0.6, textinfo='percent', textfont_size=11
            )])
            fig.update_layout(**get_layout(title="Energy Mix (Period)", height=200, showlegend=False,
                                          margin=dict(l=10, r=10, t=40, b=10)))
            st.plotly_chart(fig, use_container_width=True)
            if total_all > 0:
                st.markdown(f"☀️ **Solar:** {solar_total/total_all*100:.1f}%")
                st.markdown(f"🔋 **DG:** {dg_total/total_all*100:.1f}%")
                st.markdown(f"⚡ **Grid:** {eb_total/total_all*100:.1f}%")

        # --- EUI Chart ---
        st.markdown('<div class="section-header">🎯 EUI/Day — Actual vs Target</div>', unsafe_allow_html=True)
        eui_cols = {}
        if 'EUI_Actual' in filtered.columns:
            eui_cols['EUI_Actual'] = 'mean'
        if 'EUI_Target' in filtered.columns:
            eui_cols['EUI_Target'] = 'mean'
        if eui_cols:
            eui_daily = filtered.groupby('Date').agg(eui_cols).reset_index()
            fig = go.Figure()
            if 'EUI_Actual' in eui_daily.columns:
                fig.add_trace(go.Scatter(x=eui_daily['Date'], y=eui_daily['EUI_Actual'],
                                        name='Actual EUI/Day', line=dict(color=COLORS['eb'], width=2.5)))
            if 'EUI_Target' in eui_daily.columns:
                fig.add_trace(go.Scatter(x=eui_daily['Date'], y=eui_daily['EUI_Target'],
                                        name='Target EUI/Day', line=dict(color=COLORS['target'], width=2, dash='dash')))
            fig.update_layout(**get_layout(title="EUI/Day - Actual vs Target", height=300, yaxis_title="kWh/GSF/Day"))
            st.plotly_chart(fig, use_container_width=True)

        # Monthly EUI Summary (with MTD for current month)
        if 'EUI_Actual' in filtered.columns:
            monthly_eui = filtered.copy()
            monthly_eui['Month'] = monthly_eui['Date'].dt.to_period('M')
            current_month = pd.Timestamp.now().to_period('M')

            eui_monthly = monthly_eui.groupby('Month').agg(
                Actual=('EUI_Actual', 'mean'),
                Target=('EUI_Target', 'mean')
            ).reset_index()
            eui_monthly['Month'] = eui_monthly['Month'].astype(str)
            eui_monthly['Label'] = eui_monthly['Month'].apply(
                lambda x: x + ' (MTD)' if x == str(current_month) else x
            )

            fig = go.Figure()
            fig.add_trace(go.Bar(x=eui_monthly['Label'], y=eui_monthly['Actual'],
                                name='Actual EUI', marker_color=COLORS['eb'],
                                text=[f"{v:.3f}" for v in eui_monthly['Actual']],
                                textposition='outside'))
            fig.add_trace(go.Scatter(x=eui_monthly['Label'], y=eui_monthly['Target'],
                                    name='Target EUI', line=dict(color=COLORS['target'], width=2, dash='dash'),
                                    mode='lines+markers'))
            fig.update_layout(**get_layout(title="Monthly EUI - Actual vs Target", height=350, yaxis_title="kWh/GSF/Day"))
            st.plotly_chart(fig, use_container_width=True)


        # --- Power Heatmap (only when multiple sites selected) ---
        if len(selected_sites) > 1:
            st.markdown('<div class="section-header">🏢 Site-wise Power Heatmap</div>', unsafe_allow_html=True)
            st.caption("Compare daily power consumption across sites — darker = higher consumption")
            site_daily = filtered.groupby(['Date', 'Site'])['Total_Power_kWh'].sum().reset_index()
            pivot = site_daily.pivot(index='Site', columns='Date', values='Total_Power_kWh')
            fig = go.Figure(data=go.Heatmap(
                z=pivot.values, x=pivot.columns.strftime('%d/%m'),
                y=pivot.index, colorscale='YlOrRd', colorbar=dict(title="kWh")
            ))
            fig.update_layout(**get_layout(title="Power Heatmap (Site x Date)", height=max(350, len(pivot) * 40)))
            st.plotly_chart(fig, use_container_width=True)


    # ==================================================================
    # TAB 2: FEEDER ANALYSIS
    # ==================================================================
    with tab2:
        st.markdown('<div class="section-header">LT Panel Feeder Breakdown</div>', unsafe_allow_html=True)
        if not feeders_df.empty:
            # Use sidebar site selection instead of separate dropdown
            if len(selected_sites) == 1:
                site_sel = selected_sites[0]
            else:
                site_sel = st.selectbox("Select Site for Feeder View", selected_sites, key='feeder_site')

            site_feeders = feeders_df[feeders_df['Site'] == site_sel].copy()

            # Filter by date range from sidebar
            if not site_feeders.empty and 'Date' in site_feeders.columns:
                site_feeders = site_feeders[
                    (site_feeders['Date'].dt.date >= start_date) &
                    (site_feeders['Date'].dt.date <= end_date)
                ]

            if not site_feeders.empty:
                feeder_cols = [c for c in site_feeders.columns if c not in ['Date', 'Site']]

                # Bar chart - Latest day consumption for ALL feeders
                latest = site_feeders[site_feeders['Date'] == site_feeders['Date'].max()]
                if not latest.empty:
                    vals = latest[feeder_cols].iloc[0].dropna()
                    vals = vals[vals > 0].sort_values(ascending=True)
                    if not vals.empty:
                        fig = go.Figure(go.Bar(
                            x=vals.values, y=[str(n)[:30] for n in vals.index],
                            orientation='h', marker_color=COLORS['eb']
                        ))
                        fig.update_layout(**get_layout(
                            title=f"All Feeders - {site_sel} (Date: {site_feeders['Date'].max().strftime('%d/%m/%Y')})",
                            height=max(400, len(vals) * 25), xaxis_title="kWh (Daily Consumption)"))
                        st.plotly_chart(fig, use_container_width=True)

                # Trend chart - Top 5 feeders over time
                st.markdown("---")
                top5 = site_feeders[feeder_cols].sum().nlargest(5).index.tolist()
                if top5:
                    fig = go.Figure()
                    colors_list = ['blue', 'red', 'green', 'orange', 'purple']
                    for i, f in enumerate(top5):
                        fig.add_trace(go.Scatter(
                            x=site_feeders['Date'], y=site_feeders[f],
                            name=str(f)[:25],
                            line=dict(color=colors_list[i % 5], width=2)))
                    fig.update_layout(**get_layout(
                        title=f"Top 5 Feeders Trend - {site_sel}",
                        height=400, yaxis_title="kWh (Daily Consumption)"))
                    st.plotly_chart(fig, use_container_width=True)

                # Summary table
                st.markdown("---")
                st.markdown("**Period Summary (All Feeders)**")
                summary = site_feeders[feeder_cols].sum().reset_index()
                summary.columns = ['Feeder', 'Total kWh']
                summary = summary.sort_values('Total kWh', ascending=False)
                summary['Total kWh'] = summary['Total kWh'].apply(lambda x: f"{x:,.0f}")
                st.dataframe(summary, use_container_width=True, hide_index=True)
            else:
                st.info(f"No feeder data for {site_sel} in selected date range.")
        else:
            st.info("Feeder data not available. Connect WorkDocs to see live feeder data.")

        st.markdown("---")
        st.markdown('<div class="section-header">Submeter Readings</div>', unsafe_allow_html=True)
        if not submeters_df.empty:
            # Use same site from above (no separate dropdown)
            site_sub = submeters_df[submeters_df['Site'] == site_sel].copy()

            # Filter by date range
            if not site_sub.empty and 'Date' in site_sub.columns:
                site_sub = site_sub[
                    (site_sub['Date'].dt.date >= start_date) &
                    (site_sub['Date'].dt.date <= end_date)
                ]

            sub_cols = [c for c in site_sub.columns if c not in ['Date', 'Site']]
            if sub_cols and not site_sub.empty:
                # Bar chart - Latest day
                latest_sub = site_sub[site_sub['Date'] == site_sub['Date'].max()]
                if not latest_sub.empty:
                    vals = latest_sub[sub_cols].iloc[0].dropna()
                    vals = vals[vals > 0].sort_values(ascending=True)
                    if not vals.empty:
                        fig = go.Figure(go.Bar(
                            x=vals.values, y=[str(n)[:30] for n in vals.index],
                            orientation='h', marker_color=COLORS['solar']
                        ))
                        fig.update_layout(**get_layout(
                            title=f"Submeter Readings - {site_sel} (Date: {site_sub['Date'].max().strftime('%d/%m/%Y')})",
                            height=max(300, len(vals) * 25), xaxis_title="kWh"))
                        st.plotly_chart(fig, use_container_width=True)

                # Trend chart - Top 5 submeters
                top5_sub = site_sub[sub_cols].sum().nlargest(5).index.tolist()
                if top5_sub:
                    fig = go.Figure()
                    colors_list = ['blue', 'red', 'green', 'orange', 'purple']
                    for i, f in enumerate(top5_sub):
                        fig.add_trace(go.Scatter(
                            x=site_sub['Date'], y=site_sub[f],
                            name=str(f)[:25],
                            line=dict(color=colors_list[i % 5], width=2)))
                    fig.update_layout(**get_layout(
                        title=f"Top 5 Submeters Trend - {site_sel}",
                        height=400, yaxis_title="kWh"))
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info(f"No submeter data for {site_sel} in selected date range.")
        else:
            st.info("Submeter data not available. Connect WorkDocs to see live submeter data.")


    # ==================================================================
    # TAB 3: WATER & STP
    # ==================================================================
    with tab3:
        st.markdown('<div class="section-header">Water & STP Consumption</div>', unsafe_allow_html=True)

        water_total = filtered['Raw_Water_KL'].sum() if 'Raw_Water_KL' in filtered.columns else 0
        stp_inlet_total = filtered['STP_Inlet_KL'].sum() if 'STP_Inlet_KL' in filtered.columns else 0
        stp_outlet_total = filtered['STP_Outlet_KL'].sum() if 'STP_Outlet_KL' in filtered.columns else 0
        flushing_total = filtered['Flushing_Ltrs'].sum() if 'Flushing_Ltrs' in filtered.columns else 0
        horticulture_total = filtered['Horticulture_Ltrs'].sum() if 'Horticulture_Ltrs' in filtered.columns else 0

        w1, w2, w3, w4, w5 = st.columns(5)
        with w1:
            st.metric("Raw Water (KL)", f"{water_total:,.1f}")
        with w2:
            st.metric("STP Inlet (KL)", f"{stp_inlet_total:,.1f}")
        with w3:
            st.metric("STP Outlet (KL)", f"{stp_outlet_total:,.1f}")
        with w4:
            st.metric("Flushing (KL)", f"{flushing_total:,.0f}")
        with w5:
            st.metric("Horticulture (KL)", f"{horticulture_total:,.0f}")

        st.markdown("---")

        water_cols_agg = {}
        if 'Raw_Water_KL' in filtered.columns:
            water_cols_agg['Raw_Water_KL'] = 'sum'
        if 'STP_Inlet_KL' in filtered.columns:
            water_cols_agg['STP_Inlet_KL'] = 'sum'
        if 'STP_Outlet_KL' in filtered.columns:
            water_cols_agg['STP_Outlet_KL'] = 'sum'

        if water_cols_agg:
            water_daily = filtered.groupby('Date').agg(water_cols_agg).reset_index()
            fig = go.Figure()
            if 'Raw_Water_KL' in water_daily.columns:
                fig.add_trace(go.Bar(x=water_daily['Date'], y=water_daily['Raw_Water_KL'],
                                    name='Raw Water (KL)', marker_color='dodgerblue'))
            if 'STP_Inlet_KL' in water_daily.columns:
                fig.add_trace(go.Bar(x=water_daily['Date'], y=water_daily['STP_Inlet_KL'],
                                    name='STP Inlet (KL)', marker_color='steelblue'))
            if 'STP_Outlet_KL' in water_daily.columns:
                fig.add_trace(go.Bar(x=water_daily['Date'], y=water_daily['STP_Outlet_KL'],
                                    name='STP Outlet (KL)', marker_color='seagreen'))
            fig.update_layout(**get_layout(title="Daily Water & STP Consumption", height=350, yaxis_title="KL"))
            fig.update_layout(barmode='group')
            st.plotly_chart(fig, use_container_width=True)

        flush_cols_agg = {}
        if 'Flushing_Ltrs' in filtered.columns:
            flush_cols_agg['Flushing_Ltrs'] = 'sum'
        if 'Horticulture_Ltrs' in filtered.columns:
            flush_cols_agg['Horticulture_Ltrs'] = 'sum'

        if flush_cols_agg:
            flush_daily = filtered.groupby('Date').agg(flush_cols_agg).reset_index()
            fig = go.Figure()
            if 'Flushing_Ltrs' in flush_daily.columns:
                fig.add_trace(go.Scatter(x=flush_daily['Date'], y=flush_daily['Flushing_Ltrs'],
                                        name='Flushing (KL)', line=dict(color='orange', width=2)))
            if 'Horticulture_Ltrs' in flush_daily.columns:
                fig.add_trace(go.Scatter(x=flush_daily['Date'], y=flush_daily['Horticulture_Ltrs'],
                                        name='Horticulture (KL)', line=dict(color='green', width=2)))
            fig.update_layout(**get_layout(title="STP Outlet - Flushing & Horticulture Trend", height=300, yaxis_title="Litres"))
            st.plotly_chart(fig, use_container_width=True)

    # ==================================================================
    # TAB 4: FUEL / DG
    # ==================================================================
    with tab4:
        st.markdown('<div class="section-header">⛽ Fuel & DG Analysis</div>', unsafe_allow_html=True)

        f1, f2, f3 = st.columns(3)
        with f1:
            st.metric("⛽ HSD (Ltr)", f"{hsd_total:,.0f}")
        with f2:
            st.metric("🔋 DG (kWh)", f"{dg_total:,.0f}")
        with f3:
            eff = dg_total / hsd_total if hsd_total > 0 else 0
            st.metric("⚡ Efficiency (kWh/Ltr)", f"{eff:.2f}")

        st.markdown("---")

        daily_fuel = filtered.groupby('Date').agg({'DG_kWh': 'sum', 'HSD_Ltr': 'sum'}).reset_index()
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=daily_fuel['Date'], y=daily_fuel['HSD_Ltr'],
                            name='HSD (Ltr)', marker_color=COLORS['fuel']), secondary_y=False)
        fig.add_trace(go.Scatter(x=daily_fuel['Date'], y=daily_fuel['DG_kWh'],
                                name='DG (kWh)', line=dict(color=COLORS['dg'], width=2)), secondary_y=True)
        fig.update_layout(**get_layout(title="Daily Fuel Consumption vs DG Generation", height=380))
        fig.update_yaxes(title_text="HSD (Ltr)", secondary_y=False)
        fig.update_yaxes(title_text="DG (kWh)", secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)

        # DG by site
        if len(selected_sites) > 1:
            dg_sites = filtered.groupby('Site')['DG_kWh'].sum().reset_index().sort_values('DG_kWh', ascending=False)
            fig = go.Figure(go.Bar(x=dg_sites['Site'], y=dg_sites['DG_kWh'],
                                  marker_color=COLORS['dg']))
            fig.update_layout(**get_layout(title="DG Consumption by Site", height=300, yaxis_title="kWh"))
            st.plotly_chart(fig, use_container_width=True)

    # ==================================================================
    # TAB 5: SCRAP & WASTE
    # ==================================================================
    with tab5:
        st.markdown('<div class="section-header">WDR % - Monthly</div>', unsafe_allow_html=True)

        if 'WDR_Pct' in filtered.columns:
            wdr_data = filtered[['Date', 'WDR_Pct']].dropna(subset=['WDR_Pct'])
            if not wdr_data.empty:
                wdr_data['Month'] = wdr_data['Date'].dt.to_period('M')
                current_month = pd.Timestamp.now().to_period('M')

                wdr_monthly = wdr_data.groupby('Month')['WDR_Pct'].mean().reset_index()
                wdr_monthly['Month'] = wdr_monthly['Month'].astype(str)

                wdr_monthly['Label'] = wdr_monthly['Month'].apply(
                    lambda x: x + ' (MTD)' if x == str(current_month) else x
                )

                latest_month_wdr = wdr_monthly.iloc[-1]['WDR_Pct']
                st.metric("Current Month WDR % (MTD)", f"{latest_month_wdr:.1f}%")

                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=wdr_monthly['Label'], y=wdr_monthly['WDR_Pct'],
                    marker_color='teal', text=[f"{v:.1f}%" for v in wdr_monthly['WDR_Pct']],
                    textposition='outside'
                ))
                fig.add_hline(y=95, line_dash="dash", line_color="red",
                             annotation_text="Target 95%")
                fig.update_layout(**get_layout(title="WDR % Monthly Average", height=350, yaxis_title="WDR %"))
                fig.update_layout(yaxis=dict(range=[0, 110]))
                st.plotly_chart(fig, use_container_width=True)
                st.markdown("---")
            else:
                st.info("No WDR data available for selected period.")
        else:
            st.info("WDR data not loaded.")

        st.markdown('<div class="section-header">♻️ Scrap & Waste Management</div>', unsafe_allow_html=True)
        if not scrap_df.empty:
            # Filter scrap by selected sites and dates
            scrap_filtered = scrap_df[
                (scrap_df['Site'].isin(selected_sites)) &
                (scrap_df['Date'].dt.date >= start_date) &
                (scrap_df['Date'].dt.date <= end_date)
            ]
            if not scrap_filtered.empty:
                cat_summary = scrap_filtered.groupby('Category')['Quantity'].sum().reset_index()
                cat_summary = cat_summary.sort_values('Quantity', ascending=False).head(15)
                fig = go.Figure(go.Bar(x=cat_summary['Category'], y=cat_summary['Quantity'],
                                      marker_color=COLORS['scrap']))
                fig.update_layout(**get_layout(title="Scrap by Category (Top 15)", height=350, yaxis_title="Quantity"))
                st.plotly_chart(fig, use_container_width=True)

                if len(selected_sites) > 1:
                    site_scrap = scrap_filtered.groupby('Site')['Quantity'].sum().reset_index()
                    fig = go.Figure(go.Bar(x=site_scrap['Site'], y=site_scrap['Quantity'],
                                          marker_color=COLORS['scrap']))
                    fig.update_layout(**get_layout(title="Scrap by Site", height=300, yaxis_title="Quantity"))
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No scrap data for selected filters.")
        else:
            st.info("📋 **Scrap data not available.** Connect WorkDocs to see live scrap data.")

    # ==================================================================
    # TAB 6: SITE DEEP DIVE
    # ==================================================================
    with tab6:
        st.markdown('<div class="section-header">🔍 Site Deep Dive</div>', unsafe_allow_html=True)
        site_sel3 = st.selectbox("Select Site for Deep Dive", sorted(filtered['Site'].unique()), key='deep_site')
        site_data = filtered[filtered['Site'] == site_sel3]

        if not site_data.empty:
            # Site summary KPIs
            s1, s2, s3, s4, s5 = st.columns(5)
            with s1:
                st.metric("⚡ EB Total", f"{site_data['EB_kWh'].sum():,.0f} kWh")
            with s2:
                st.metric("🔋 DG Total", f"{site_data['DG_kWh'].sum():,.0f} kWh")
            with s3:
                st.metric("☀️ Solar Total", f"{site_data['Solar_kWh'].sum():,.0f} kWh")
            with s4:
                st.metric("Avg EUI (Actual)", f"{site_data['EUI_Actual'].mean():.3f}" if 'EUI_Actual' in site_data.columns else "N/A")
            with s5:
                st.metric("💧 Water Total", f"{site_data['Raw_Water_KL'].sum():,.1f} KL")

            st.markdown("---")

            fig = make_subplots(rows=3, cols=1, shared_xaxes=True,
                              subplot_titles=("Power Consumption", "Water Usage", "EUI Tracking"),
                              vertical_spacing=0.08)

            # Power
            fig.add_trace(go.Scatter(x=site_data['Date'], y=site_data['EB_kWh'],
                                    name='EB', line=dict(color=COLORS['eb'])), row=1, col=1)
            fig.add_trace(go.Scatter(x=site_data['Date'], y=site_data['DG_kWh'],
                                    name='DG', line=dict(color=COLORS['dg'])), row=1, col=1)
            fig.add_trace(go.Scatter(x=site_data['Date'], y=site_data['Solar_kWh'],
                                    name='Solar', line=dict(color=COLORS['solar'])), row=1, col=1)

            # Water
            fig.add_trace(go.Scatter(x=site_data['Date'], y=site_data['Raw_Water_KL'],
                                    name='Raw Water', line=dict(color=COLORS['water'])), row=2, col=1)
            if 'Treated_Water_KL' in site_data.columns:
                fig.add_trace(go.Scatter(x=site_data['Date'], y=site_data['Treated_Water_KL'],
                                        name='Treated', line=dict(color=COLORS['solar'])), row=2, col=1)

            # EUI
            if 'EUI_Actual' in site_data.columns:
                fig.add_trace(go.Scatter(x=site_data['Date'], y=site_data['EUI_Actual'],
                                        name='Actual EUI', line=dict(color=COLORS['eb'])), row=3, col=1)
            if 'EUI_Target' in site_data.columns:
                fig.add_trace(go.Scatter(x=site_data['Date'], y=site_data['EUI_Target'],
                                        name='Target', line=dict(color=COLORS['target'], dash='dash')), row=3, col=1)


            fig.update_layout(height=800, title_text=f"Site Deep Dive — {site_sel3}",
                            template="plotly_white", showlegend=True)
            st.plotly_chart(fig, use_container_width=True)

    # ==================================================================
    # TAB 7: HEALTH
    # ==================================================================
    with tab7:
        st.markdown('<div class="section-header">🏥 Data Health & Coverage</div>', unsafe_allow_html=True)

        h1, h2, h3, h4 = st.columns(4)
        with h1:
            st.metric("📊 Total Records", f"{len(filtered):,}")
        with h2:
            st.metric("🏢 Active Sites", f"{filtered['Site'].nunique()}")
        with h3:
            days = (filtered['Date'].max() - filtered['Date'].min()).days
            st.metric("📅 Days of Data", f"{days}")
        with h4:
            completeness = (1 - filtered.isnull().sum().sum() / (filtered.shape[0] * filtered.shape[1])) * 100
            st.metric("✅ Completeness", f"{completeness:.1f}%")

        st.markdown("---")

        # Data completeness by column
        col_completeness = (1 - filtered.isnull().mean()) * 100
        fig = go.Figure(go.Bar(
            x=col_completeness.values,
            y=col_completeness.index,
            orientation='h',
            marker_color=COLORS['solar']
        ))
        fig.update_layout(**get_layout(title="Data Completeness by Column (%)", height=500, xaxis_title="%"))
        st.plotly_chart(fig, use_container_width=True)

        # Sites data summary
        agg_dict = {
            'Records': ('Date', 'count'),
            'First_Date': ('Date', 'min'),
            'Last_Date': ('Date', 'max'),
        }
        if 'EUI_Actual' in filtered.columns:
            agg_dict['Avg_EUI'] = ('EUI_Actual', 'mean')
        site_summary = filtered.groupby('Site').agg(**agg_dict).reset_index()

        st.dataframe(site_summary, use_container_width=True)

else:
    st.warning("⚠️ No data available for the selected filters. Try adjusting the site selection or date range.")

