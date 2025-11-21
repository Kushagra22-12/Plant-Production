import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Production Analytics Dashboard", layout="wide")

@st.cache_data
def load_data(path):
    raw = pd.read_excel(path, sheet_name=0, header=None, engine='openpyxl')
    # Detect header row dynamically
    required_cols = {'Plant', 'Line', 'Grade'}
    header_idx = None
    for i in range(len(raw)):
        row_vals = set(str(x).strip() for x in raw.iloc[i].tolist())
        if required_cols.issubset(row_vals):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError('Could not find header row with columns Plant, Line, Grade')
    columns = [str(x).strip() for x in raw.iloc[header_idx].tolist()]
    df = raw.iloc[header_idx+1:].copy().reset_index(drop=True)
    if df.shape[1] > len(columns):
        df = df.iloc[:, :len(columns)]
    df.columns = columns
    df = df.dropna(how='all').copy()

    def classify_row(grade_val):
        if pd.isna(grade_val):
            return 'Unknown'
        s = str(grade_val).strip().lower()
        if s == 'line total':
            return 'Line Total'
        if s == 'plant total':
            return 'Plant Total'
        if s == 'grand total':
            return 'Grand Total'
        return 'Detail'

    df['RowType'] = df['Grade'].map(classify_row)

    rename_map = {
        'Today Qty (No)': 'TodayQty',
        'Today Qty %': 'TodayQtyPct',
        'Today KW': 'TodayKW',
        'MTD Qty (No)': 'MTDQty',
        'MTD Qty %': 'MTDQtyPct',
        'MTD KW': 'MTDKW',
    }
    for col in list(df.columns):
        c_norm = str(col).strip()
        if c_norm in rename_map:
            df.rename(columns={col: rename_map[c_norm]}, inplace=True)

    num_cols = ['TodayQty', 'TodayQtyPct', 'TodayKW', 'MTDQty', 'MTDQtyPct', 'MTDKW']
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    for c in ['Plant', 'Line', 'Grade']:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    if 'TodayKW' in df.columns and 'TodayQty' in df.columns:
        df['TodayKW_per_Unit'] = np.where(df['TodayQty']>0, df['TodayKW']/df['TodayQty'], np.nan)
    if 'MTDKW' in df.columns and 'MTDQty' in df.columns:
        df['MTDKW_per_Unit'] = np.where(df['MTDQty']>0, df['MTDKW']/df['MTDQty'], np.nan)

    return df

@st.cache_data
def detect_date_range(path):
    raw = pd.read_excel(path, sheet_name=0, header=None, engine='openpyxl')
    required_cols = {'Plant', 'Line', 'Grade'}
    header_idx = None
    for i in range(len(raw)):
        row_vals = set(str(x).strip() for x in raw.iloc[i].tolist())
        if required_cols.issubset(row_vals):
            header_idx = i
            break
    pre_header_text = '\n'.join('\t'.join(map(str, raw.iloc[i].tolist())) for i in range(header_idx)) if header_idx is not None else ''
    import re
    m = re.search(r'(\d{1,2} [A-Za-z]+ 20\d{2})\s+to\s+(\d{1,2} [A-Za-z]+ 20\d{2})', pre_header_text)
    if m:
        return m.group(1), m.group(2)
    return None, None

# Sidebar — file selector and filters
st.sidebar.title("Controls")
excel_file = st.sidebar.file_uploader("Upload Production Summary Excel", type=["xlsx"]) 

if excel_file is None:
    st.info("Please upload the Excel file to begin.")
    st.stop()

# Load data
try:
    df = load_data(excel_file)
except Exception as e:
    st.error(f"Failed to parse the uploaded file: {e}")
    st.stop()

start_date, end_date = detect_date_range(excel_file)

# Filters
plants = sorted([p for p in df['Plant'].dropna().unique().tolist() if p and p.lower()!='nan']) if 'Plant' in df.columns else []
lines = sorted([l for l in df['Line'].dropna().unique().tolist() if l and l.lower()!='nan']) if 'Line' in df.columns else []
grades = sorted([g for g in df['Grade'].dropna().unique().tolist() if g and g.lower()!='nan']) if 'Grade' in df.columns else []

sel_plants = st.sidebar.multiselect('Plant', plants, default=plants)
sel_lines = st.sidebar.multiselect('Line', lines, default=lines)
sel_grades = st.sidebar.multiselect('Grade', grades, default=grades)
include_totals = st.sidebar.checkbox('Include Total Rows (Line/Plant/Grand)', value=False)

# Apply filters
fdf = df.copy()
if sel_plants:
    fdf = fdf[fdf['Plant'].isin(sel_plants)]
if sel_lines:
    fdf = fdf[fdf['Line'].isin(sel_lines)]
if sel_grades:
    fdf = fdf[fdf['Grade'].isin(sel_grades)]
if not include_totals:
    fdf = fdf[fdf['RowType'] == 'Detail']

# Title
title = "Production Analytics Dashboard"
if start_date and end_date:
    title += f" — {start_date} to {end_date}"
st.title(title)

# KPIs
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("Today Qty (No)", value=f"{np.nansum(fdf.get('TodayQty', np.nan)):.0f}")
with c2:
    st.metric("Today KW", value=f"{np.nansum(fdf.get('TodayKW', np.nan)):.2f}")
with c3:
    st.metric("MTD Qty (No)", value=f"{np.nansum(fdf.get('MTDQty', np.nan)):.0f}")
with c4:
    st.metric("MTD KW", value=f"{np.nansum(fdf.get('MTDKW', np.nan)):.2f}")

st.markdown("---")

# Distribution by Plant
if 'Plant' in fdf.columns and 'MTDQty' in fdf.columns:
    st.subheader("MTD Qty by Plant")
    plant_summary = fdf.groupby('Plant', dropna=False)['MTDQty'].sum().sort_values(ascending=False)
    st.bar_chart(plant_summary)

# Distribution by Line (within Plant)
if {'Plant','Line','MTDQty'}.issubset(fdf.columns):
    st.subheader("MTD Qty by Line")
    line_summary = fdf.groupby(['Plant','Line'])['MTDQty'].sum().reset_index().sort_values(['Plant','MTDQty'], ascending=[True, False])
    # Pivot for chart to show lines per plant
    pivot = line_summary.pivot(index='Line', columns='Plant', values='MTDQty').fillna(0)
    st.bar_chart(pivot)

# Grade mix (% of MTD Qty)
if 'Grade' in fdf.columns and 'MTDQty' in fdf.columns:
    st.subheader("Grade Mix (by MTD Qty)")
    gmix = fdf.groupby('Grade')['MTDQty'].sum().sort_values(ascending=False)
    gmix_pct = (gmix / gmix.sum() * 100).round(2)
    st.dataframe(pd.DataFrame({'MTDQty': gmix, 'Share %': gmix_pct}))

# Efficiency
if {'Plant','Line','MTDKW_per_Unit'}.issubset(fdf.columns):
    st.subheader("MTD KW per Unit by Line")
    eff = fdf.groupby(['Plant','Line'])['MTDKW_per_Unit'].mean().reset_index()
    eff_pivot = eff.pivot(index='Line', columns='Plant', values='MTDKW_per_Unit').round(4)
    st.bar_chart(eff_pivot)

st.markdown("---")

# Detailed table
st.subheader("Detailed Rows (after filters)")
show_cols = [c for c in [
    'Plant','Line','Grade','TodayQty','TodayQtyPct','TodayKW','MTDQty','MTDQtyPct','MTDKW','TodayKW_per_Unit','MTDKW_per_Unit'
] if c in fdf.columns]

st.dataframe(fdf[show_cols])

# Download filtered data
csv = fdf.to_csv(index=False).encode('utf-8')
st.download_button("Download filtered data as CSV", data=csv, file_name="filtered_production_data.csv", mime="text/csv")

# Simple note
st.caption("Upload a new Excel in the sidebar to refresh the dashboard.")
