import streamlit as st
import pandas as pd
import numpy as np

# Try Plotly; if unavailable, fall back to Streamlit native charts
PLOTLY_AVAILABLE = True
try:
    import plotly.express as px
except Exception:
    PLOTLY_AVAILABLE = False

st.set_page_config(page_title="Production Analytics — Resilient", layout="wide")

@st.cache_data
def load_data(upload):
    raw = pd.read_excel(upload, sheet_name=0, header=None, engine='openpyxl')
    required_cols = {'Plant', 'Line', 'Grade'}
    header_idx = None
    for i in range(len(raw)):
        row_vals = set(str(x).strip() for x in raw.iloc[i].tolist())
        if required_cols.issubset(row_vals):
            header_idx = i
            break
    columns = [str(x).strip() for x in raw.iloc[header_idx].tolist()]
    df = raw.iloc[header_idx+1:].copy().reset_index(drop=True)
    if df.shape[1] > len(columns):
        df = df.iloc[:, :len(columns)]
    df.columns = columns

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

st.sidebar.title("Controls")
file = st.sidebar.file_uploader("Upload Production Summary Excel", type=["xlsx"]) 
if file is None:
    st.info("Upload the Excel to begin.")
    st.stop()

try:
    df = load_data(file)
except Exception as e:
    st.error(f"Parsing failed: {e}")
    st.stop()

plants = sorted([p for p in df['Plant'].dropna().unique().tolist() if p and p.lower()!='nan'])
lines = sorted([l for l in df['Line'].dropna().unique().tolist() if l and l.lower()!='nan'])
grades = sorted([g for g in df['Grade'].dropna().unique().tolist() if g and g.lower()!='nan'])

sel_plants = st.sidebar.multiselect('Plant', plants, default=plants)
sel_lines = st.sidebar.multiselect('Line', lines, default=lines)
sel_grades = st.sidebar.multiselect('Grade', grades, default=grades)
include_totals = st.sidebar.checkbox('Include Total Rows (Line/Plant/Grand)', value=False)
metric_basis = st.sidebar.radio("Metric basis", options=["MTD", "Today"], index=0)
normalize = st.sidebar.checkbox('Normalize to % share (grade/plant/line)', value=False)

fdf = df.copy()
if sel_plants:
    fdf = fdf[fdf['Plant'].isin(sel_plants)]
if sel_lines:
    fdf = fdf[fdf['Line'].isin(sel_lines)]
if sel_grades:
    fdf = fdf[fdf['Grade'].isin(sel_grades)]
if not include_totals:
    fdf = fdf[fdf['RowType'] == 'Detail']

qty_col = 'MTDQty' if metric_basis == 'MTD' else 'TodayQty'
kw_col = 'MTDKW' if metric_basis == 'MTD' else 'TodayKW'
kw_per_unit_col = 'MTDKW_per_Unit' if metric_basis == 'MTD' else 'TodayKW_per_Unit'

st.title("Production Analytics — Resilient")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric(f"{metric_basis} Qty (No)", value=f"{np.nansum(fdf.get(qty_col, np.nan)):.0f}")
with c2:
    st.metric(f"{metric_basis} KW", value=f"{np.nansum(fdf.get(kw_col, np.nan)):.2f}")
with c3:
    st.metric("Distinct Plants", value=len(fdf['Plant'].unique()))
with c4:
    st.metric("Distinct Lines", value=len(fdf['Line'].unique()))

st.markdown("---")

# Plant summary
plant_sum = fdf.groupby('Plant')[qty_col].sum().sort_values(ascending=False)
if PLOTLY_AVAILABLE:
    import plotly.express as px
    fig1 = px.bar(plant_sum, x=plant_sum.index, y=plant_sum.values, title=f"{metric_basis} Qty by Plant")
    fig1.update_layout(xaxis_title="Plant", yaxis_title=f"{metric_basis} Qty")
    st.plotly_chart(fig1, use_container_width=True)
else:
    st.warning("Plotly not found. Falling back to Streamlit bar_chart. Add 'plotly' to requirements to enable interactive charts.")
    st.bar_chart(plant_sum)

# Grade mix
gsum = fdf.groupby('Grade')[qty_col].sum().sort_values(ascending=False)
if normalize and gsum.sum()>0:
    gvalues = (gsum/gsum.sum()*100).round(2)
else:
    gvalues = gsum

if PLOTLY_AVAILABLE:
    fig_g = px.bar(gvalues, x=gvalues.index, y=gvalues.values, title="Grade Mix")
    st.plotly_chart(fig_g, use_container_width=True)
else:
    st.bar_chart(gvalues)

# Efficiency by Line
eff_line = fdf.groupby(['Plant','Line'])[kw_per_unit_col].mean().reset_index()
if PLOTLY_AVAILABLE:
    fig_e = px.bar(eff_line, x='Line', y=kw_per_unit_col, color='Plant', title=f"{metric_basis} KW/Unit by Line")
    st.plotly_chart(fig_e, use_container_width=True)
else:
    pivot = eff_line.pivot(index='Line', columns='Plant', values=kw_per_unit_col).fillna(0)
    st.bar_chart(pivot)

st.caption("Install 'plotly' and 'kaleido' for full interactive charts and downloads.")
