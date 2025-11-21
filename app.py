
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Production Analytics — Advanced", layout="wide")

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

@st.cache_data
def detect_date_range(upload):
    raw = pd.read_excel(upload, sheet_name=0, header=None, engine='openpyxl')
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

# Sidebar: file + filters
st.sidebar.title("Controls")
file = st.sidebar.file_uploader("Upload Production Summary Excel", type=["xlsx"]) 
if file is None:
    st.info("Upload the Excel to begin.")
    st.stop()

# Load
try:
    df = load_data(file)
except Exception as e:
    st.error(f"Parsing failed: {e}")
    st.stop()

start_date, end_date = detect_date_range(file)

# Filters
plants = sorted([p for p in df['Plant'].dropna().unique().tolist() if p and p.lower()!='nan'])
lines = sorted([l for l in df['Line'].dropna().unique().tolist() if l and l.lower()!='nan'])
grades = sorted([g for g in df['Grade'].dropna().unique().tolist() if g and g.lower()!='nan'])

sel_plants = st.sidebar.multiselect('Plant', plants, default=plants)
sel_lines = st.sidebar.multiselect('Line', lines, default=lines)
sel_grades = st.sidebar.multiselect('Grade', grades, default=grades)
include_totals = st.sidebar.checkbox('Include Total Rows (Line/Plant/Grand)', value=False)

metric_basis = st.sidebar.radio("Metric basis", options=["MTD", "Today"], index=0)
normalize = st.sidebar.checkbox('Normalize to % share (grade/plant/line)', value=False)

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

qty_col = 'MTDQty' if metric_basis == 'MTD' else 'TodayQty'
kw_col = 'MTDKW' if metric_basis == 'MTD' else 'TodayKW'
kw_per_unit_col = 'MTDKW_per_Unit' if metric_basis == 'MTD' else 'TodayKW_per_Unit'

# Title
title = f"Production Analytics — {metric_basis}"
if start_date and end_date:
    title += f" — {start_date} to {end_date}"
st.title(title)

# Tabs
tab_overview, tab_grade, tab_eff, tab_lines, tab_pivot, tab_export = st.tabs([
    "Overview", "Grade Analytics", "Efficiency", "Lines & Trends", "Pivot & Heatmaps", "Export"
])

# === Overview ===
with tab_overview:
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric(f"{metric_basis} Qty (No)", value=f"{np.nansum(fdf.get(qty_col, np.nan)):.0f}")
    with c2:
        st.metric(f"{metric_basis} KW", value=f"{np.nansum(fdf.get(kw_col, np.nan)):.2f}")
    with c3:
        st.metric("Distinct Plants", value=len(fdf['Plant'].unique()))
    with c4:
        st.metric("Distinct Lines", value=len(fdf['Line'].unique()))

    chart_type = st.selectbox("Chart type", options=["Bar", "Line", "Area"], index=0)

    # Plant summary
    if 'Plant' in fdf.columns and qty_col in fdf.columns:
        plant_sum = fdf.groupby('Plant')[qty_col].sum().sort_values(ascending=False)
        fig1 = None
        if chart_type == 'Bar':
            fig1 = px.bar(plant_sum, x=plant_sum.index, y=plant_sum.values, title=f"{metric_basis} Qty by Plant")
        elif chart_type == 'Line':
            fig1 = px.line(plant_sum, x=plant_sum.index, y=plant_sum.values, title=f"{metric_basis} Qty by Plant")
        else:
            fig1 = px.area(plant_sum, x=plant_sum.index, y=plant_sum.values, title=f"{metric_basis} Qty by Plant")
        fig1.update_layout(xaxis_title="Plant", yaxis_title=f"{metric_basis} Qty")
        st.plotly_chart(fig1, use_container_width=True)

    # Line summary by Plant
    if {'Plant','Line',qty_col}.issubset(fdf.columns):
        line_sum = fdf.groupby(['Plant','Line'])[qty_col].sum().reset_index()
        fig2 = px.bar(line_sum, x='Line', y=qty_col, color='Plant', barmode='group', title=f"{metric_basis} Qty by Line (grouped by Plant)")
        fig2.update_layout(xaxis_title="Line", yaxis_title=f"{metric_basis} Qty")
        st.plotly_chart(fig2, use_container_width=True)

# === Grade Analytics ===
with tab_grade:
    st.subheader("Grade mix and drilldowns")
    topn = st.slider("Show top N grades (by Qty)", min_value=1, max_value=max(1, len(grades)), value=min(5, len(grades)))

    gsum = fdf.groupby('Grade')[qty_col].sum().sort_values(ascending=False)
    if normalize and gsum.sum() > 0:
        gvalues = (gsum / gsum.sum() * 100).round(2)
        y_lbl = "Share %"
    else:
        gvalues = gsum
        y_lbl = f"{metric_basis} Qty"

    gtop = gvalues.head(topn)
    fig_g = px.bar(gtop, x=gtop.index, y=gtop.values, title=f"Top {topn} Grades — {y_lbl}")
    fig_g.update_layout(xaxis_title="Grade", yaxis_title=y_lbl)
    st.plotly_chart(fig_g, use_container_width=True)

    # Grade by Plant
    gp = fdf.groupby(['Plant','Grade'])[qty_col].sum().reset_index()
    if normalize:
        gp['Share'] = gp.groupby('Plant')[qty_col].apply(lambda s: s/s.sum()*100)
        y = 'Share'
        y_title = 'Share %'
    else:
        y = qty_col
        y_title = f"{metric_basis} Qty"
    fig_gp = px.bar(gp, x='Plant', y=y, color='Grade', title=f"Grade mix by Plant ({y_title})", barmode='stack')
    fig_gp.update_layout(yaxis_title=y_title)
    st.plotly_chart(fig_gp, use_container_width=True)

    # Grade by Line (faceted by Plant)
    gl = fdf.groupby(['Plant','Line','Grade'])[qty_col].sum().reset_index()
    fig_gl = px.bar(gl, x='Line', y=qty_col, color='Grade', facet_col='Plant', facet_col_wrap=2, title=f"Grade by Line (faceted by Plant)", barmode='stack')
    fig_gl.update_layout(yaxis_title=f"{metric_basis} Qty")
    st.plotly_chart(fig_gl, use_container_width=True)

# === Efficiency ===
with tab_eff:
    st.subheader("KW per Unit by Line and Grade")
    # Line-level efficiency
    eff_line = fdf.groupby(['Plant','Line'])[kw_per_unit_col].mean().reset_index()
    fig_e1 = px.bar(eff_line, x='Line', y=kw_per_unit_col, color='Plant', title=f"{metric_basis} KW per Unit — by Line")
    fig_e1.update_layout(yaxis_title=f"{metric_basis} KW/Unit")
    st.plotly_chart(fig_e1, use_container_width=True)

    # Grade-level efficiency
    eff_grade = fdf.groupby('Grade')[kw_per_unit_col].mean().reset_index()
    fig_e2 = px.bar(eff_grade, x='Grade', y=kw_per_unit_col, title=f"{metric_basis} KW per Unit — by Grade")
    fig_e2.update_layout(yaxis_title=f"{metric_basis} KW/Unit")
    st.plotly_chart(fig_e2, use_container_width=True)

    # Distribution boxplot by grade
    fig_e3 = px.box(fdf.dropna(subset=[kw_per_unit_col]), x='Grade', y=kw_per_unit_col, points='outliers', title=f"Distribution of {metric_basis} KW per Unit by Grade")
    fig_e3.update_layout(yaxis_title=f"{metric_basis} KW/Unit")
    st.plotly_chart(fig_e3, use_container_width=True)

# === Lines & Trends ===
with tab_lines:
    st.subheader("Line plots and comparisons")
    # A line plot of Qty across Lines grouped by Plant (no time series available, so categorical line)
    line_sum = fdf.groupby(['Plant','Line'])[qty_col].sum().reset_index()
    fig_l1 = px.line(line_sum, x='Line', y=qty_col, color='Plant', markers=True, title=f"{metric_basis} Qty across Lines (by Plant)")
    fig_l1.update_layout(yaxis_title=f"{metric_basis} Qty")
    st.plotly_chart(fig_l1, use_container_width=True)

    # Area plot of Grade contributions per Line
    gl = fdf.groupby(['Line','Grade'])[qty_col].sum().reset_index()
    fig_l2 = px.area(gl, x='Line', y=qty_col, color='Grade', title=f"Grade contributions per Line")
    fig_l2.update_layout(yaxis_title=f"{metric_basis} Qty")
    st.plotly_chart(fig_l2, use_container_width=True)

# === Pivot & Heatmaps ===
with tab_pivot:
    st.subheader("Interactive pivot and heatmaps")
    # Pivot Line x Grade
    pivot_lg = fdf.pivot_table(index='Line', columns='Grade', values=qty_col, aggfunc='sum', fill_value=0)
    st.dataframe(pivot_lg)
    fig_h1 = px.imshow(pivot_lg, aspect='auto', color_continuous_scale='Blues', title=f"Heatmap — {metric_basis} Qty (Line × Grade)")
    st.plotly_chart(fig_h1, use_container_width=True)

    # Pivot Plant x Grade
    pivot_pg = fdf.pivot_table(index='Plant', columns='Grade', values=qty_col, aggfunc='sum', fill_value=0)
    st.dataframe(pivot_pg)
    fig_h2 = px.imshow(pivot_pg, aspect='auto', color_continuous_scale='Greens', title=f"Heatmap — {metric_basis} Qty (Plant × Grade)")
    st.plotly_chart(fig_h2, use_container_width=True)

# === Export ===
with tab_export:
    st.subheader("Download filtered data & figures")
    csv = fdf.to_csv(index=False).encode('utf-8')
    st.download_button("Download filtered data as CSV", data=csv, file_name="filtered_data.csv", mime="text/csv")

    # Optional: export key charts as PNG using plotly + kaleido
    try:
        fig_names = {
            'overview_plant.png': px.bar(fdf.groupby('Plant')[qty_col].sum().sort_values(ascending=False), 
                                         x=fdf.groupby('Plant')[qty_col].sum().sort_values(ascending=False).index,
                                         y=fdf.groupby('Plant')[qty_col].sum().sort_values(ascending=False).values,
                                         title=f"{metric_basis} Qty by Plant"),
            'grade_top.png': px.bar(fdf.groupby('Grade')[qty_col].sum().sort_values(ascending=False).head(5), 
                                    x=fdf.groupby('Grade')[qty_col].sum().sort_values(ascending=False).head(5).index,
                                    y=fdf.groupby('Grade')[qty_col].sum().sort_values(ascending=False).head(5).values,
                                    title=f"Top 5 Grades by {metric_basis} Qty"),
        }
        for fname, fig in fig_names.items():
            import io
            buf = fig.to_image(format='png')
            st.download_button(label=f"Download {fname}", data=buf, file_name=fname, mime='image/png')
    except Exception as e:
        st.caption("(Install 'kaleido' to enable figure downloads)")
