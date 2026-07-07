"""
Nium Playbook — Global Payout Capability Explorer
UI entry point. Scraping & data logic lives in scraper.py.

Run:
  streamlit run nium_capability_explorer.py
"""

import os
import time
from datetime import datetime

import pandas as pd
import streamlit as st

from config import CACHE_TTL_SECONDS
from scraper import (
    scrape_dataset, transform_raw_to_wide, save_scraped_data,
    log_scrape_failures, create_formatted_excel, get_last_updated,
    COUNTRIES, FI_PATH, NON_FI_PATH,
)

# ═══════════════════════════════════════════════════════════════════════════════
# DATA LAYER  (cache wrapper lives here — needs @st.cache_data)
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=CACHE_TTL_SECONDS)
def load_data_cached():
    """Load FI and Non-FI Excel files. Falls back to any matching file in cwd."""
    fi_path  = str(FI_PATH)  if FI_PATH.exists()  else None
    nfi_path = str(NON_FI_PATH) if NON_FI_PATH.exists() else None

    if not fi_path or not nfi_path:
        for f in os.listdir(os.getcwd()):
            if f.endswith('.xlsx'):
                if 'FI' in f and 'Non' not in f and fi_path is None:
                    fi_path = f
                elif 'Non' in f and 'FI' in f and nfi_path is None:
                    nfi_path = f

    if not fi_path or not nfi_path:
        return None, None

    fi  = pd.read_excel(fi_path)
    nfi = pd.read_excel(nfi_path)
    fi['_t']  = 'FI'
    nfi['_t'] = 'Non-FI'
    return fi, nfi


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG & STYLING
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="Nium Playbook", page_icon="⚡", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
*, .main, .block-container, .stApp { font-family: 'Segoe UI', system-ui, -apple-system, BlinkMacSystemFont, 'Helvetica Neue', Arial, sans-serif !important; }
.block-container { padding: 1rem 2rem 2rem 2rem; max-width: 100%; }
#MainMenu, footer, header { visibility: hidden; }
hr { margin: 0.5rem 0; opacity: 0.08; }

.n-hdr {
    display: flex; align-items: center; justify-content: space-between;
    padding: 1rem 1.5rem; background: #0A0A0A; border-radius: 14px;
    margin-bottom: 1rem; position: relative; overflow: hidden;
    border: 1px solid rgba(0,212,170,0.15);
}
.n-hdr::before { content:''; position:absolute; top:-60%; right:-8%; width:280px; height:280px; background:radial-gradient(circle,rgba(0,212,170,0.12) 0%,transparent 65%); pointer-events:none; }
.n-hdr::after { content:''; position:absolute; bottom:-40%; left:20%; width:200px; height:200px; background:radial-gradient(circle,rgba(0,229,191,0.06) 0%,transparent 65%); pointer-events:none; }
.n-logo { display:flex; align-items:center; gap:0.7rem; }
.n-icon { width:36px; height:36px; background:linear-gradient(135deg,#00D4AA,#00E5BF); border-radius:9px; display:flex; align-items:center; justify-content:center; font-size:1.1rem; font-weight:800; color:#0A0A0A; box-shadow:0 0 20px rgba(0,212,170,0.3); }
.n-hdr h1 { color:#fff; font-size:1.3rem; font-weight:700; margin:0; padding:0; border:none; letter-spacing:-0.02em; }
.n-hdr .n-sub { color:#00D4AA; font-size:0.68rem; font-weight:600; letter-spacing:0.08em; text-transform:uppercase; }
.n-hdr .n-updated { color:#64748b; font-size:0.65rem; font-weight:400; }

.stRadio > div { flex-direction:row !important; gap:0; }
.stRadio > div > label { background:#f7faf9 !important; padding:0.35rem 1.2rem !important; border:1px solid #d1e8e0 !important; font-weight:600 !important; font-size:0.78rem !important; color:#2d6a5a !important; margin:0 !important; }
.stRadio > div > label:first-child { border-radius:8px 0 0 8px !important; border-right:none !important; }
.stRadio > div > label:last-child { border-radius:0 8px 8px 0 !important; }
.stRadio > div > label[data-checked="true"] { background:#0A0A0A !important; color:#00D4AA !important; border-color:#0A0A0A !important; }

.n-stats { display:flex; gap:0.5rem; }
.n-sc { flex:1; background:#f7faf9; border:1px solid #d1e8e0; border-radius:10px; padding:0.55rem 0.9rem; border-left:3px solid #00D4AA; }
.n-sc .v { font-size:1.4rem; font-weight:800; color:#0A0A0A; line-height:1.2; letter-spacing:-0.03em; }
.n-sc .l { font-size:0.6rem; color:#00997A; font-weight:700; text-transform:uppercase; letter-spacing:0.07em; }

.n-fc { background:#fff; border:1px solid #e0efe8; border-radius:12px; padding:0.8rem 1rem; margin-bottom:0.6rem; border-top:3px solid #00D4AA; }
.n-fc-t { font-size:0.65rem; font-weight:700; color:#00997A; text-transform:uppercase; letter-spacing:0.08em; margin-bottom:0.3rem; }

.stSelectbox label, .stMultiSelect label { font-size:0.72rem !important; font-weight:600 !important; color:#2d6a5a !important; text-transform:uppercase !important; letter-spacing:0.04em !important; }
.stCheckbox label span { font-size:0.78rem !important; font-weight:500 !important; color:#1a1a1a !important; }
.stCheckbox { margin-bottom:-0.6rem; }
.stMultiSelect [data-baseweb="tag"] { background:#00D4AA !important; color:#0A0A0A !important; border-radius:6px !important; font-size:0.7rem !important; font-weight:700 !important; }
[data-testid="stMetric"] { display:none; }

.stDownloadButton > button { background:#0A0A0A !important; color:#00D4AA !important; border:1px solid rgba(0,212,170,0.3) !important; border-radius:8px !important; font-weight:700 !important; font-size:0.78rem !important; padding:0.45rem 1rem !important; }
.stDownloadButton > button:hover { background:#111 !important; border-color:#00D4AA !important; box-shadow:0 4px 15px rgba(0,212,170,0.2); transform:translateY(-1px); }
.stButton > button { border-radius:8px; font-weight:600; font-size:0.78rem; border:1px solid #d1e8e0; background:#0A0A0A !important; color:#00D4AA !important; border-color:rgba(0,212,170,0.3) !important; }
.stButton > button:hover { background:#111 !important; border-color:#00D4AA !important; box-shadow:0 3px 10px rgba(0,212,170,0.18); transform:translateY(-1px); }

[data-testid="stDataFrame"] { border:1px solid #e0efe8; border-radius:10px; overflow:hidden; }
.streamlit-expanderHeader { font-size:0.8rem !important; font-weight:600 !important; background:#f7faf9 !important; border-radius:8px !important; color:#2d6a5a !important; }
.n-bdg { display:inline-block; background:rgba(0,212,170,0.12); color:#00997A; padding:0.18rem 0.6rem; border-radius:20px; font-size:0.7rem; font-weight:700; border:1px solid rgba(0,212,170,0.25); }
.stAlert { border-radius:10px; font-size:0.82rem; }
.n-ft { text-align:center; color:#94a3b8; font-size:0.68rem; padding:1.2rem 0 0.5rem 0; margin-top:1.5rem; border-top:1px solid #e0efe8; }
.n-ft span { color:#00D4AA; font-weight:600; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def get_cascaded_options(data, col, active_filters):
    """Return sorted unique values for col after applying all other active filters."""
    d = data.copy()
    for c, v in active_filters.items():
        if c != col and v:
            d = d[d[c].isin(v)]
    return sorted(d[col].dropna().unique())


def format_cell_value(v):
    if pd.isna(v) or str(v) in ('', 'nan', 'None'):
        return "—"
    return str(v)


# ═══════════════════════════════════════════════════════════════════════════════
# UI — HEADER
# ═══════════════════════════════════════════════════════════════════════════════

last_updated = get_last_updated()
updated_html = (
    f'<div class="n-updated">Last updated: {last_updated}</div>'
    if last_updated else
    '<div class="n-updated">No data yet — click 🔄 Refresh</div>'
)

st.markdown(f"""
<div class="n-hdr">
    <div class="n-logo">
        <div class="n-icon">N</div>
        <div>
            <h1>Nium Playbook</h1>
            <div class="n-sub">Global Payout Capability Explorer</div>
            {updated_html}
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# UI — REFRESH BUTTON
# ═══════════════════════════════════════════════════════════════════════════════

_IS_CLOUD = os.environ.get("STREAMLIT_SHARING_MODE") or os.environ.get("IS_STREAMLIT_CLOUD") or "/mount/src" in os.path.abspath(__file__)

if 'scrape_confirm' not in st.session_state:
    st.session_state.scrape_confirm = False

if _IS_CLOUD:
    st.info("🔄 **Data updates automatically every Monday at 2:00 AM UTC (7:30 AM IST / 10:00 AM SGT)** — no action needed. The latest Nium Playbook data is always pulled fresh and reflected here within the hour.", icon="✅")
elif not st.session_state.scrape_confirm:
    if st.button("🔄 Refresh Data (Scrape from Nium Playbook)", use_container_width=True, key="refresh"):
        st.session_state.scrape_confirm = True
        st.rerun()

if not _IS_CLOUD and st.session_state.scrape_confirm:
    st.markdown("""
    <div style="background:#fff8e1;border:1px solid #f59e0b;border-radius:10px;padding:0.9rem 1.2rem;margin-bottom:0.8rem;">
        <strong style="color:#92400e;">⚠️ Confirm Full Scrape</strong><br>
        <span style="font-size:0.85rem;color:#78350f;">
        This will scrape <strong>{len(COUNTRIES)} countries × 2 datasets</strong> from Nium Playbook.
        It takes approximately <strong>25 minutes</strong> — do not close the browser tab.
        Existing data in <code>data/</code> will be overwritten.
        </span>
    </div>
    """, unsafe_allow_html=True)

    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        start_scrape = st.button("▶ Yes, Start Scrape", use_container_width=True, key="confirm_scrape")
    with btn_col2:
        if st.button("✕ Cancel", use_container_width=True, key="cancel_scrape"):
            st.session_state.scrape_confirm = False
            st.rerun()

    if start_scrape:
        st.session_state.scrape_confirm = False
        st.markdown("---")

        st.markdown("**Scraping FI dataset** (Financial Institutions)")
        fi_bar    = st.progress(0, text="Waiting...")
        fi_status = st.empty()

        st.markdown("**Scraping Non-FI dataset**")
        nfi_bar    = st.progress(0, text="Waiting...")
        nfi_status = st.empty()

        overall_status = st.empty()

        try:
            all_results = {}
            bars = {"FI": (fi_bar, fi_status), "Non-FI": (nfi_bar, nfi_status)}

            for ds_type in ["FI", "Non-FI"]:
                bar, status = bars[ds_type]
                status.caption(f"🚀 Starting {ds_type} scrape...")
                raw_rows, failed = scrape_dataset(ds_type, bar, status)

                status.caption(f"🔄 Transforming {ds_type} data ({len(raw_rows):,} raw rows)...")
                df_scraped = transform_raw_to_wide(raw_rows)

                if not df_scraped.empty:
                    app_path, _ = save_scraped_data(df_scraped, ds_type)
                    all_results[ds_type] = {"rows": len(df_scraped), "failed_countries": failed}
                    bar.progress(1.0, text=f"✅ {ds_type} done — {len(df_scraped):,} corridors")
                    status.caption(f"✅ Saved to {app_path}")
                else:
                    all_results[ds_type] = {"rows": 0, "failed_countries": failed}
                    bar.progress(1.0, text=f"⚠️ {ds_type} — no data returned")

            failures_by_dataset = {ds: info['failed_countries'] for ds, info in all_results.items()}
            log_path = log_scrape_failures(failures_by_dataset)
            if log_path:
                st.session_state['_scrape_failures'] = failures_by_dataset
                st.session_state['_scrape_log_path'] = log_path

            summary_lines = []
            for ds_type, info in all_results.items():
                summary_lines.append(f"**{ds_type}:** {info['rows']:,} corridors")
                if info.get('failed_countries'):
                    summary_lines.append(f"&nbsp;&nbsp;⚠️ {len(info['failed_countries'])} countries skipped — logged to `logs/scrape_failures.log`")

            overall_status.success("✅ Scrape complete!\n\n" + "\n\n".join(summary_lines))
            st.cache_data.clear()
            time.sleep(2)
            st.rerun()

        except Exception as e:
            overall_status.error(f"❌ Scrape failed: {str(e)}")
            st.info("Make sure Chrome is installed and you have internet access.")

# ═══════════════════════════════════════════════════════════════════════════════
# UI — LOAD & DISPLAY DATA
# ═══════════════════════════════════════════════════════════════════════════════

with st.spinner("Loading capability data..."):
    fi_data, non_fi_data = load_data_cached()

if fi_data is None or non_fi_data is None:
    st.info("📭 No data files found. Click **🔄 Refresh Data** above to scrape from Nium Playbook, or place Excel files in the `data/` folder.")
    st.stop()

REQUIRED_COLS = {"Country", "Payment Mode", "Currency", "TAT", "Supported Modes"}
for _name, _df in [("FI", fi_data), ("Non-FI", non_fi_data)]:
    _missing = REQUIRED_COLS - set(_df.columns)
    if _missing:
        st.error(
            f"❌ **{_name} dataset** is missing required columns: "
            f"`{'`, `'.join(sorted(_missing))}`\n\n"
            "Check that your Excel file matches the expected structure, then reload."
        )
        st.stop()

# ── Post-scrape failure banner (persists across the rerun) ──
if '_scrape_failures' in st.session_state:
    _failures = st.session_state.pop('_scrape_failures')
    _log_path = st.session_state.pop('_scrape_log_path', None)
    _total    = sum(len(v) for v in _failures.values())
    with st.expander(f"⚠️ {_total} countr{'ies' if _total != 1 else 'y'} were skipped in the last scrape", expanded=True):
        for ds, countries in _failures.items():
            if countries:
                st.markdown(f"**{ds}** — {len(countries)} skipped:")
                st.markdown("\n".join(f"- {c}" for c in sorted(countries)))
        if _log_path:
            st.caption(f"Full log: `{_log_path}`")

# ── Toggle + Stats ──
r1, r2 = st.columns([1, 5])
with r1:
    ds = st.radio("_", ["FI", "Non-FI"], horizontal=True, label_visibility="collapsed")
data = (fi_data if ds == "FI" else non_fi_data).copy().drop('_t', axis=1)
with r2:
    st.markdown(f"""
    <div class="n-stats">
        <div class="n-sc"><div class="l">Corridors</div><div class="v">{len(data):,}</div></div>
        <div class="n-sc"><div class="l">Countries</div><div class="v">{data['Country'].nunique()}</div></div>
        <div class="n-sc"><div class="l">Pay Modes</div><div class="v">{data['Payment Mode'].nunique()}</div></div>
        <div class="n-sc"><div class="l">Currencies</div><div class="v">{data['Currency'].nunique()}</div></div>
    </div>
    """, unsafe_allow_html=True)

# ── Filters ──
st.markdown('<div class="n-fc"><div class="n-fc-t">⚡ Filters — cascading</div>', unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns(4)
active_filters = {}

with f1:
    countries = st.multiselect("Country", sorted(data['Country'].dropna().unique()), key="c", placeholder="All")
if countries:
    active_filters['Country'] = countries

with f2:
    modes = st.multiselect("Payment Mode", get_cascaded_options(data, 'Payment Mode', active_filters), key="m", placeholder="All")
if modes:
    active_filters['Payment Mode'] = modes

with f3:
    currencies = st.multiselect("Currency", get_cascaded_options(data, 'Currency', active_filters), key="cr", placeholder="All")
if currencies:
    active_filters['Currency'] = currencies

with f4:
    tat_opts = sorted(
        get_cascaded_options(data, 'TAT', active_filters),
        key=lambda x: {'Realtime': 0, 'T0': 1, 'T1': 2, 'T2': 3}.get(x, 9),
    )
    tats = st.multiselect("TAT", tat_opts, key="t", placeholder="All")

txn = []

st.markdown('</div>', unsafe_allow_html=True)

# ── Apply Filters ──
df = data.copy()
if countries:              df = df[df['Country'].isin(countries)]
if modes:                  df = df[df['Payment Mode'].isin(modes)]
if currencies:             df = df[df['Currency'].isin(currencies)]
if tats:                   df = df[df['TAT'].isin(tats)]
if txn:                    df = df[df.apply(lambda r: all(t in str(r['Supported Modes']).split(', ') for t in txn), axis=1)]

# ── Results Bar ──
active_count = sum([bool(countries), bool(modes), bool(currencies), bool(tats)])
rc1, rc2 = st.columns([5, 1.5])
with rc1:
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:0.6rem;padding:0.3rem 0;">
        <span style="font-size:1.05rem;font-weight:700;color:#0A0A0A;">{len(df):,} results</span>
        <span class="n-bdg">{active_count} filter{"s" if active_count != 1 else ""}</span>
    </div>
    """, unsafe_allow_html=True)
with rc2:
    if len(df) > 0:
        excel_buf = create_formatted_excel(df, df.columns.tolist(), ds)
        st.download_button(
            "⬇ Download Excel", excel_buf,
            f"Nium_Capability_Matrix_{ds}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ── Table ──
if len(df) > 0:
    st.dataframe(df, use_container_width=True, hide_index=True)
else:
    st.info("No corridors match. Adjust filters above.")

# ── Footer ──
st.markdown(f'<div class="n-ft">Powered by <span>Nium</span> · Playbook · {datetime.now().strftime("%B %Y")}</div>', unsafe_allow_html=True)
