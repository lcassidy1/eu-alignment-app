"""
EU Alignment Dashboard — Streamlit app
Hosted on Streamlit Community Cloud.
To update data: run scrapers locally, commit new CSVs to GitHub.
"""

from pathlib import Path

import pandas as pd
import streamlit as st

CACHE_UN80 = Path(__file__).parent / "scraped_results.csv"
CACHE_GA79 = Path(__file__).parent / "scraped_results_ga79.csv"

# Session date boundaries
SESSIONS = {
    "UN80 (Sep 2025 – present)":      ("2025-09-08", "2099-12-31"),
    "GA79 (Sep 2024 – Sep 2025)":     ("2024-09-10", "2025-09-09"),
    "All sessions combined":           ("2000-01-01", "2099-12-31"),
    "Custom date range":               None,
}

COUNTRIES = [
    "Montenegro", "Moldova", "Bosnia and Herzegovina", "North Macedonia",
    "Albania", "Ukraine", "Georgia", "San Marino", "Monaco", "Andorra",
    "Serbia", "Armenia", "Iceland", "Turkey", "Norway", "Liechtenstein", "UK",
    "Azerbaijan",
]

st.set_page_config(page_title="EU Alignment Tracker", layout="wide")
st.title("EU Delegation UN New York — Alignment Tracker")


@st.cache_data
def load_data():
    dfs = []
    for cache in [CACHE_UN80, CACHE_GA79]:
        if cache.exists():
            df = pd.read_csv(cache)
            df["_date"] = pd.to_datetime(df["Date"], errors="coerce")
            dfs.append(df)
    combined = pd.concat(dfs, ignore_index=True)
    # Drop exact duplicates (same URL appearing in both caches)
    combined = combined.drop_duplicates(subset=["URL"]).copy()
    combined = combined.dropna(subset=["_date"]).sort_values("_date", ascending=False)
    return combined


df_all = load_data()
abs_min = df_all["_date"].min().date()
abs_max = df_all["_date"].max().date()

# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.header("Filters")

session_choice = st.sidebar.selectbox("Session / time period", list(SESSIONS.keys()))

if session_choice == "Custom date range":
    start_date = st.sidebar.date_input("From", value=abs_min, min_value=abs_min, max_value=abs_max)
    end_date   = st.sidebar.date_input("To",   value=abs_max, min_value=abs_min, max_value=abs_max)
else:
    bounds = SESSIONS[session_choice]
    start_date = pd.Timestamp(bounds[0]).date()
    end_date   = pd.Timestamp(bounds[1]).date()
    st.sidebar.caption(f"{start_date.strftime('%d/%m/%Y')} → {min(end_date, abs_max).strftime('%d/%m/%Y')}")

show_only_aligned = st.sidebar.checkbox("Only statements with alignment clause", value=False)

selected_country = st.sidebar.selectbox(
    "Filter by country",
    options=["— all —"] + COUNTRIES,
)

st.sidebar.divider()
st.sidebar.caption("Data last updated by EU Delegation NY team. To refresh, run scrapers locally and push updated CSVs to GitHub.")

# ── Filter ────────────────────────────────────────────────────────────────────
mask = (df_all["_date"].dt.date >= start_date) & (df_all["_date"].dt.date <= end_date)
df = df_all[mask].copy()

if show_only_aligned:
    df = df[df["Has Alignment"] == True]

if selected_country != "— all —" and selected_country in df.columns:
    df = df[df[selected_country] == True]

# ── Metrics ───────────────────────────────────────────────────────────────────
total   = len(df)
aligned = int(df["Has Alignment"].sum()) if "Has Alignment" in df.columns else 0

c1, c2, c3 = st.columns(3)
c1.metric("Total statements", total)
c2.metric("With alignment clause", aligned)
c3.metric("Coverage", f"{min(start_date, abs_max).strftime('%d/%m/%Y')} – {min(end_date, abs_max).strftime('%d/%m/%Y')}")

st.divider()

# ── Summary table ─────────────────────────────────────────────────────────────
st.subheader("Alignment % by Country")

df_aligned = df[df["Has Alignment"] == True]
n = len(df_aligned)

rows = []
for c in COUNTRIES:
    if c not in df_aligned.columns:
        continue
    cnt = int(df_aligned[c].fillna(False).astype(bool).sum())
    rows.append({
        "Country": c,
        "Times Aligned": cnt,
        "Out of (alignment stmts)": n,
        "Alignment %": f"{cnt/n*100:.1f}%" if n else "—",
    })

df_summary = pd.DataFrame(rows)
df_summary["_sort"] = df_summary["Times Aligned"] / n if n else 0
df_summary = df_summary.sort_values("_sort", ascending=False).drop(columns="_sort").reset_index(drop=True)
df_summary.index += 1

st.dataframe(df_summary, use_container_width=True, height=min(50 + 35 * len(df_summary), 700))

st.divider()

# ── Statement list ────────────────────────────────────────────────────────────
label = f"Statements where {selected_country} aligned" if selected_country != "— all —" else "All statements in range"
st.subheader(label)

df_view = df.copy()
df_view["Date"] = df_view["_date"].dt.strftime("%d/%m/%Y")
df_view["Aligned?"] = df_view["Has Alignment"].map({True: "✅", False: "—"})

display_cols = [c for c in ["Date", "Aligned?", "Title", "URL"] if c in df_view.columns]

st.dataframe(
    df_view[display_cols].reset_index(drop=True),
    use_container_width=True,
    column_config={
        "URL": st.column_config.LinkColumn("URL"),
        "Title": st.column_config.TextColumn("Title", width="large"),
    },
    height=450,
)
