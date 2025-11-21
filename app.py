import streamlit as st
import pandas as pd
import numpy as np
import sys

# --- Detect Plotly availability and version ---
PLOTLY_AVAILABLE = True
KALEIDO_AVAILABLE = False
plotly_version = None
try:
    import plotly
    import plotly.express as px
    plotly_version = plotly.__version__
    # Check kaleido (for PNG downloads)
    try:
        import kaleido  # noqa: F401
        KALEIDO_AVAILABLE = True
    except Exception:
        KALEIDO_AVAILABLE = False
except Exception:
    PLOTLY_AVAILABLE = False

st.set_page_config(page_title="Production Analytics — Main", layout="wide")

# Sidebar diagnostics
with st.sidebar.expander("Environment Diagnostics", expanded=False):
    st.write("**Python**:", sys.executable)
    if PLOTLY_AVAILABLE:
        st.success(f"Plotly available — v{plotly_version}")
        st.write("Kaleido for PNG:", "✅" if KALEIDO_AVAILABLE else "❌")
    else:
        st.warning("Plotly not available — charts will use Streamlit natives.")

@st.cache_data
def load_data(upload):
    # Load and clean the Production Summary Excel with dynamic header detection.
    raw = pd.read_excel(upload, sheet_name=0, header=None, engine='openpyxl')
    required_cols = {'Plant', 'Line', 'Grade'}
    header_idx = None
    for i in range(len(raw)):
        row_vals = set(str(x).strip() for x in raw.iloc[i].tolist())
        if required_cols.issubset(row_vals):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not find header row (expecting columns: Plant, Line, Grade)")

    columns = [str(x).strip() for x in raw.iloc[header_idx].tolist()]
    df = raw.iloc[header_idx+1:].copy().reset_index(drop=True)
    # Trim any extra columns beyond header length
    if df.shape[1] > len(columns):
        df = df.iloc[:, :len(columns)]
    df.columns = columns

    # Tag totals vs details
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

    # Normalize columns
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

    # Efficiency metrics
    if {'TodayKW','TodayQty'}.issubset(df.columns):
        df['TodayKW_per_Unit'] = np.where(df['TodayQty']>0, df['TodayKW']/df['TodayQty'], np.nan)
    if {'MTDKW','MTDQty'}.issubset(df.columns):
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

# --- Sidebar controls ---
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

# --- Helper plotting functions ---
def plot_series_bar(series, title, x_label, y_label):
    if PLOTLY_AVAILABLE:
        fig = px.bar(series, x=series.index, y=series.values, title=title)
        fig.update_layout(xaxis_title=x_label, yaxis_title=y_label)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.subheader(title)
        st.bar_chart(series)

def plot_series_line(series, title, x_label, y_label):
    if PLOTLY_AVAILABLE:
        fig = px.line(series, x=series.index, y=series.values, title=title)
        fig.update_layout(xaxis_title=x_label, yaxis_title=y_label)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.subheader(title)
        st.line_chart(series)

def plot_series_area(series, title, x_label, y_label):
    if PLOTLY_AVAILABLE:
        fig = px.area(series, x=series.index, y=series.values, title=title)
        fig.update_layout(xaxis_title=x_label, yaxis_title=y_label)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.subheader(title)
        st.area_chart(series)

# For DataFrames
def plot_df_bar(df_plot, x, y, color=None, title="", barmode='group'): 
    if PLOTLY_AVAILABLE:
        fig = px.bar(df_plot, x=x, y=y, color=color, title=title, barmode=barmode)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.subheader(title)
        if color and color in df_plot.columns:
            pivot = df_plot.pivot(index=x, columns=color, values=y).fillna(0)
            st.bar_chart(pivot)
        else:
            st.bar_chart(df_plot.set_index(x)[y])

def plot_df_line(df_plot, x, y, color=None, title=""):
    if PLOTLY_AVAILABLE:
        fig = px.line(df_plot, x=x, y=y, color=color, title=title, markers=True)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.subheader(title)
        if color and color in df_plot.columns:
            pivot = df_plot.pivot(index=x, columns=color, values=y).fillna(0)
            st.line_chart(pivot)
        else:
            st.line_chart(df_plot.set_index(x)[y])

# Title
page_title = f"Production Analytics — {metric_basis}"
if start_date and end_date:
    page_title += f" — {start_date} to {end_date}"
st.title(page_title)

# Tabs
(tab_overview, tab_grade, tab_eff, tab_lines, tab_pivot, tab_export) = st.tabs([
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
        if chart_type == 'Bar':
            plot_series_bar(plant_sum, f"{metric_basis} Qty by Plant", "Plant", f"{metric_basis} Qty")
        elif chart_type == 'Line':
            plot_series_line(plant_sum, f"{metric_basis} Qty by Plant", "Plant", f"{metric_basis} Qty")
        else:
            plot_series_area(plant_sum, f"{metric_basis} Qty by Plant", "Plant", f"{metric_basis} Qty")

    # Line summary by Plant
    if {'Plant','Line',qty_col}.issubset(fdf.columns):
        line_sum = fdf.groupby(['Plant','Line'])[qty_col].sum().reset_index()
        plot_df_bar(line_sum, x='Line', y=qty_col, color='Plant', title=f"{metric_basis} Qty by Line (grouped by Plant)", barmode='group')

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
    plot_series_bar(gtop, f"Top {topn} Grades — {y_lbl}", "Grade", y_lbl)

    # Grade by Plant
    gp = fdf.groupby(['Plant','Grade'])[qty_col].sum().reset_index()
    if normalize:
        gp['Share'] = gp.groupby('Plant')[qty_col].apply(lambda s: s/s.sum()*100)
        y = 'Share'; y_title = 'Share %'
    else:
        y = qty_col; y_title = f"{metric_basis} Qty"
    plot_df_bar(gp, x='Plant', y=y, color='Grade', title=f"Grade mix by Plant ({y_title})", barmode='stack')

    # Grade by Line (faceted by Plant) — Plotly only for facets; fallback with grouped bars per plant
    gl = fdf.groupby(['Plant','Line','Grade'])[qty_col].sum().reset_index()
    if PLOTLY_AVAILABLE:
        fig = px.bar(gl, x='Line', y=qty_col, color='Grade', facet_col='Plant', facet_col_wrap=2, title=f"Grade by Line (faceted by Plant)", barmode='stack')
        fig.update_layout(yaxis_title=f"{metric_basis} Qty")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.subheader("Grade by Line (grouped by Plant)")
        for p in sorted(gl['Plant'].unique()):
            sub = gl[gl['Plant']==p]
            st.caption(f"Plant {p}")
            pivot = sub.pivot(index='Line', columns='Grade', values=qty_col).fillna(0)
            st.bar_chart(pivot)

# === Efficiency ===
with tab_eff:
    st.subheader("KW per Unit by Line and Grade")
    eff_line = fdf.groupby(['Plant','Line'])[kw_per_unit_col].mean().reset_index()
    plot_df_bar(eff_line, x='Line', y=kw_per_unit_col, color='Plant', title=f"{metric_basis} KW per Unit — by Line")

    eff_grade = fdf.groupby('Grade')[kw_per_unit_col].mean().reset_index()
    if PLOTLY_AVAILABLE:
        fig_e2 = px.bar(eff_grade, x='Grade', y=kw_per_unit_col, title=f"{metric_basis} KW per Unit — by Grade")
        fig_e2.update_layout(yaxis_title=f"{metric_basis} KW/Unit")
        st.plotly_chart(fig_e2, use_container_width=True)

        # Distribution boxplot by grade
        df_box = fdf.dropna(subset=[kw_per_unit_col])
        fig_box = px.box(df_box, x='Grade', y=kw_per_unit_col, points='outliers', title=f"Distribution of {metric_basis} KW per Unit by Grade")
        fig_box.update_layout(yaxis_title=f"{metric_basis} KW/Unit")
        st.plotly_chart(fig_box, use_container_width=True)
    else:
        st.subheader(f"{metric_basis} KW per Unit — by Grade")
        st.bar_chart(eff_grade.set_index('Grade')[kw_per_unit_col])
        st.caption("Boxplot requires Plotly; showing summary table instead.")
        st.dataframe(df.groupby('Grade')[kw_per_unit_col].describe().round(4))

# === Lines & Trends ===
with tab_lines:
    st.subheader("Line plots and comparisons")
    line_sum = fdf.groupby(['Plant','Line'])[qty_col].sum().reset_index()
    plot_df_line(line_sum, x='Line', y=qty_col, color='Plant', title=f"{metric_basis} Qty across Lines (by Plant)")

    # Area plot of Grade contributions per Line (Plotly preferred)
    gl = fdf.groupby(['Line','Grade'])[qty_col].sum().reset_index()
    if PLOTLY_AVAILABLE:
        fig_l2 = px.area(gl, x='Line', y=qty_col, color='Grade', title=f"Grade contributions per Line")
        fig_l2.update_layout(yaxis_title=f"{metric_basis} Qty")
        st.plotly_chart(fig_l2, use_container_width=True)
    else:
        st.subheader("Grade contributions per Line")
        pivot = gl.pivot(index='Line', columns='Grade', values=qty_col).fillna(0)
        st.area_chart(pivot)

# === Pivot & Heatmaps ===
with tab_pivot:
    st.subheader("Interactive pivot and heatmaps")
    pivot_lg = fdf.pivot_table(index='Line', columns='Grade', values=qty_col, aggfunc='sum', fill_value=0)
    st.dataframe(pivot_lg)
    if PLOTLY_AVAILABLE:
        fig_h1 = px.imshow(pivot_lg, aspect='auto', color_continuous_scale='Blues', title=f"Heatmap — {metric_basis} Qty (Line × Grade)")
        st.plotly_chart(fig_h1, use_container_width=True)
    else:
        st.caption("Heatmap requires Plotly; showing pivot above.")

    pivot_pg = fdf.pivot_table(index='Plant', columns='Grade', values=qty_col, aggfunc='sum', fill_value=0)
    st.dataframe(pivot_pg)
    if PLOTLY_AVAILABLE:
        fig_h2 = px.imshow(pivot_pg, aspect='auto', color_continuous_scale='Greens', title=f"Heatmap — {metric_basis} Qty (Plant × Grade)")
        st.plotly_chart(fig_h2, use_container_width=True)
    else:
        st.caption("Heatmap requires Plotly; showing pivot above.")

# === Export ===
with tab_export:
    st.subheader("Download filtered data & figures")
    csv = fdf.to_csv(index=False).encode('utf-8')
    st.download_button("Download filtered data as CSV", data=csv, file_name="filtered_data.csv", mime="text/csv")

    if PLOTLY_AVAILABLE and KALEIDO_AVAILABLE:
        st.caption("Download key charts as PNG")
        # Example: Plant summary
        plant_sum = fdf.groupby('Plant')[qty_col].sum().sort_values(ascending=False)
        fig_overview = px.bar(plant_sum, x=plant_sum.index, y=plant_sum.values, title=f"{metric_basis} Qty by Plant")
        png_bytes = fig_overview.to_image(format='png')
        st.download_button(label="Download overview_plant.png", data=png_bytes, file_name="overview_plant.png", mime="image/png")
    elif PLOTLY_AVAILABLE and not KALEIDO_AVAILABLE:
        st.caption("(Install 'kaleido' to enable figure downloads)")
    else:
        st.caption("Plotly not available — figure downloads are disabled.")

st.caption("Tip: Switch metric basis (MTD/Today) and toggle % normalization in the sidebar to change all views.")
