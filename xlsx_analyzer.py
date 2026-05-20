import streamlit as st
import pandas as pd
import plotly.express as px

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Excel Analyzer",
    page_icon="📗",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

    html, body, [class*="css"]  { font-family: 'IBM Plex Sans', sans-serif; }
    h1, h2, h3                  { font-family: 'IBM Plex Mono', monospace; }

    /* dark canvas */
    .main { background-color: #0b0f0e; color: #e8ede9; }
    section[data-testid="stSidebar"] {
        background-color: #0f1412;
        border-right: 1px solid #1e2b26;
    }

    /* metric cards */
    .metric-card {
        background: #121a17;
        border: 1px solid #1e2b26;
        border-left: 3px solid #3ddc84;
        border-radius: 8px;
        padding: 0.9rem 1.2rem;
        text-align: center;
    }
    .metric-label { font-size: 0.7rem; color: #7a9e8a; letter-spacing: 0.12em; text-transform: uppercase; }
    .metric-value { font-size: 1.55rem; font-family: 'IBM Plex Mono', monospace; color: #3ddc84; margin-top: 2px; }

    /* section titles */
    .section-title {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.78rem;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: #7a9e8a;
        border-bottom: 1px solid #1e2b26;
        padding-bottom: 0.45rem;
        margin-bottom: 0.9rem;
    }

    /* tab strip */
    .stTabs [data-baseweb="tab-list"]  { gap: 6px; border-bottom: 1px solid #1e2b26; }
    .stTabs [data-baseweb="tab"]       { font-family: 'IBM Plex Mono', monospace; font-size: 0.8rem; padding: 6px 18px; border-radius: 4px 4px 0 0; color: #7a9e8a; background: transparent; border: none; }
    .stTabs [aria-selected="true"]     { color: #3ddc84 !important; border-bottom: 2px solid #3ddc84 !important; }

    div[data-testid="stSelectbox"] label,
    div[data-testid="stMultiSelect"] label { font-size: 0.78rem; color: #7a9e8a; letter-spacing: 0.05em; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("# 📗 Excel Analyzer")
st.markdown("Upload an Excel workbook to explore sheets, data, and auto-generated charts.")
st.divider()

# ── File Upload ────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Drop your Excel file here",
    type=["xlsx", "xlsm", "xls"],
    help="Supports .xlsx, .xlsm, and .xls files.",
)

if uploaded_file is None:
    st.info("👆 Upload an Excel file (.xlsx / .xlsm / .xls) to get started.")
    st.stop()

# ── Load workbook (all sheets) ─────────────────────────────────────────────────
@st.cache_data
def load_workbook(file):
    return pd.read_excel(file, sheet_name=None, engine="openpyxl")

all_sheets = load_workbook(uploaded_file)
sheet_names = list(all_sheets.keys())

# Sheet selector
selected_sheet = st.selectbox("📄 Select sheet", sheet_names, key="sheet_select")
df = all_sheets[selected_sheet].copy()

# ── Column classification ──────────────────────────────────────────────────────
num_cols  = df.select_dtypes(include="number").columns.tolist()
cat_cols  = df.select_dtypes(exclude="number").columns.tolist()

# ── Summary metrics ────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5 = st.columns(5)
metrics = [
    (c1, "Sheets",       str(len(sheet_names))),
    (c2, "Rows",         f"{len(df):,}"),
    (c3, "Columns",      str(df.shape[1])),
    (c4, "Numeric cols", str(len(num_cols))),
    (c5, "Text cols",    str(len(cat_cols))),
]
for col, label, value in metrics:
    with col:
        st.markdown(
            f'<div class="metric-card">'
            f'<div class="metric-label">{label}</div>'
            f'<div class="metric-value">{value}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
st.markdown("")

# ── Shared chart theme ─────────────────────────────────────────────────────────
CHART_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font_family="IBM Plex Sans",
    font_color="#e8ede9",
    margin=dict(l=10, r=10, t=44, b=10),
)
COLOR_SEQ = ["#3ddc84","#00b4d8","#f4a261","#e76f51","#a8dadc","#457b9d","#f1c40f","#9b59b6"]

def style(fig):
    fig.update_layout(**CHART_LAYOUT)
    fig.update_xaxes(gridcolor="#1e2b26", zeroline=False)
    fig.update_yaxes(gridcolor="#1e2b26", zeroline=False)
    return fig

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_data, tab_bar, tab_line, tab_stats = st.tabs(["🗂 Data", "📊 Bar Chart", "📈 Line Chart", "🔢 Statistics"])

# ── Tab 1 · Data table ─────────────────────────────────────────────────────────
with tab_data:
    st.markdown('<div class="section-title">Sheet Data</div>', unsafe_allow_html=True)

    # Optional column filter
    show_cols = st.multiselect(
        "Visible columns (leave empty to show all)",
        df.columns.tolist(),
        default=[],
        key="col_filter",
    )
    display_df = df[show_cols] if show_cols else df
    st.dataframe(display_df, use_container_width=True, height=400)

    # Download filtered data as CSV
    csv_bytes = display_df.to_csv(index=False).encode()
    st.download_button(
        "⬇ Download as CSV",
        data=csv_bytes,
        file_name=f"{selected_sheet}.csv",
        mime="text/csv",
    )

# ── Tab 2 · Bar chart ──────────────────────────────────────────────────────────
with tab_bar:
    st.markdown('<div class="section-title">Bar Chart</div>', unsafe_allow_html=True)

    if not num_cols:
        st.warning("No numeric columns found in this sheet.")
    else:
        bco1, bco2, bco3 = st.columns(3)
        with bco1:
            bar_y = st.selectbox("Value (Y axis)", num_cols, key="bar_y")
        with bco2:
            x_choices = cat_cols if cat_cols else df.columns.tolist()
            bar_x = st.selectbox("Category (X axis)", x_choices, key="bar_x")
        with bco3:
            bar_agg = st.selectbox("Aggregation", ["sum","mean","count","max","min"], key="bar_agg")

        bar_color_opt = st.selectbox(
            "Color by (optional)",
            ["— none —"] + (cat_cols if cat_cols else []),
            key="bar_color",
        )
        color_col = None if bar_color_opt == "— none —" else bar_color_opt

        agg_df = (
            df.groupby(bar_x)[bar_y]
            .agg(bar_agg)
            .reset_index()
            .sort_values(by=bar_y, ascending=False)
        )

        fig_bar = px.bar(
            agg_df,
            x=bar_x, y=bar_y,
            color=bar_x if color_col is None else None,
            color_discrete_sequence=COLOR_SEQ,
            title=f"{bar_agg.capitalize()} of {bar_y} by {bar_x}",
        )
        st.plotly_chart(style(fig_bar), use_container_width=True)

# ── Tab 3 · Line chart ─────────────────────────────────────────────────────────
with tab_line:
    st.markdown('<div class="section-title">Line Chart</div>', unsafe_allow_html=True)

    if not num_cols:
        st.warning("No numeric columns found in this sheet.")
    else:
        lco1, lco2 = st.columns([1, 2])
        with lco1:
            line_x = st.selectbox("X axis", df.columns.tolist(), key="line_x")
        with lco2:
            line_ys = st.multiselect(
                "Y axis — numeric columns",
                num_cols,
                default=num_cols[:min(3, len(num_cols))],
                key="line_ys",
            )

        if not line_ys:
            st.info("Select at least one numeric column for the Y axis.")
        else:
            plot_df = df[[line_x] + line_ys].copy()

            # Try datetime parsing for nicer axis
            try:
                plot_df[line_x] = pd.to_datetime(plot_df[line_x])
            except Exception:
                pass

            plot_df[line_x] = plot_df[line_x].astype(str)

            fig_line = px.line(
                plot_df,
                x=line_x, y=line_ys,
                color_discrete_sequence=COLOR_SEQ,
                title=f"{', '.join(line_ys)} over {line_x}",
                markers=True,
            )
            st.plotly_chart(style(fig_line), use_container_width=True)

# ── Tab 4 · Statistics ────────────────────────────────────────────────────────
with tab_stats:
    st.markdown('<div class="section-title">Descriptive Statistics</div>', unsafe_allow_html=True)

    if not num_cols:
        st.info("No numeric columns to summarise in this sheet.")
    else:
        stats_df = df[num_cols].describe().T
        st.dataframe(stats_df.style.format("{:.3f}"), use_container_width=True)

    if cat_cols:
        st.markdown('<div class="section-title" style="margin-top:1.5rem">Categorical Columns — Unique Value Counts</div>', unsafe_allow_html=True)
        vc_col = st.selectbox("Column", cat_cols, key="vc_col")
        vc_df  = df[vc_col].value_counts().reset_index()
        vc_df.columns = [vc_col, "Count"]

        fig_vc = px.bar(
            vc_df.head(30),
            x=vc_col, y="Count",
            color=vc_col,
            color_discrete_sequence=COLOR_SEQ,
            title=f"Value counts — {vc_col}",
        )
        st.plotly_chart(style(fig_vc), use_container_width=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.divider()
st.markdown(
    "<center style='color:#2a4a3a;font-size:0.72rem;'>Built with Streamlit · Plotly · Pandas · OpenPyXL</center>",
    unsafe_allow_html=True,
)
