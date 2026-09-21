"""
ST Dashboard — Supertails Automation Website
Professional UI that replicates the Excel workbook reports exactly.
Now includes: Manual Rider Count, Breach Order IDs + Reasons.
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
import tempfile
import os

from process_orders import run_full_pipeline

# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ST Dashboard | Supertails",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1a1a2e;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #666;
        margin-bottom: 1.5rem;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.6rem;
    }
    .breach-row {
        background-color: #fff5f5;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────
with st.sidebar:
    st.title("ST Dashboard")
    st.caption("Supertails × Zippee")
    
    st.markdown("---")
    st.subheader("Settings")
    
    mode = st.radio(
        "Processing Mode",
        options=["all", "daily"],
        format_func=lambda x: "All Data (Full History)" if x == "all" else "Daily Mode (Today + Open + Delivered Today)",
        index=1,  # Default to daily
        help="'Daily' matches your handwritten SOP + delivered-today orders."
    )
    
    st.markdown("---")
    st.subheader("Data Source")
    
    st.markdown("**Upload the 3 store CSVs** (exported from Zippee / Pidge):")
    kalyan_file = st.file_uploader("1. Kalyan Nagar (Hyperlocal)", type=["csv"], key="k")
    varthur_file = st.file_uploader("2. Varthur", type=["csv"], key="v")
    indranagar_file = st.file_uploader("3. Indranagar", type=["csv"], key="i")
    
    st.markdown("---")
    st.subheader("Manual Rider Count")
    st.caption("Enter rider count for each store (used for Avg Order/Rider)")
    
    rider_kalyan = st.number_input("Kalyan Nagar Riders", min_value=0, value=5, step=1, key="rider_k")
    rider_indra = st.number_input("Indranagar Riders", min_value=0, value=4, step=1, key="rider_i")
    rider_varthur = st.number_input("Varthur Riders", min_value=0, value=3, step=1, key="rider_v")
    
    rider_map = {
        "Kalyan Nagar": rider_kalyan,
        "Indranagar": rider_indra,
        "Varthur": rider_varthur
    }
    
    st.markdown("---")
    st.markdown("**Channel Mapping**")
    st.code("Kalyan  → Supertail_Hyperlocal\nIndranagar → Supertails-Indranagar\nVarthur → Supertails-Varthur", language=None)
    
    st.markdown("---")
    st.caption("Logic = 100% identical to Excel SOP")

# ─────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────
st.markdown('<p class="main-header">ST Dashboard</p>', unsafe_allow_html=True)
st.markdown(
    f'<p class="sub-header">Supertails Delivery Performance · {datetime.now().strftime("%d %b %Y %H:%M")}</p>',
    unsafe_allow_html=True
)

# ─────────────────────────────────────────────
# Run button
# ─────────────────────────────────────────────
col_btn1, col_btn2, _ = st.columns([1, 1, 3])
with col_btn1:
    run_clicked = st.button("Generate Reports", type="primary", use_container_width=True)
with col_btn2:
    if "result" in st.session_state:
        st.success("Reports ready")

if run_clicked:
    with st.spinner("Processing orders & applying Excel SOP logic..."):
        try:
            if not (kalyan_file and varthur_file and indranagar_file):
                st.error("Please upload all three CSV files (Kalyan Nagar, Varthur, Indranagar).")
                st.stop()
            
            with tempfile.TemporaryDirectory() as tmp:
                k_path = os.path.join(tmp, "k.csv")
                v_path = os.path.join(tmp, "v.csv")
                i_path = os.path.join(tmp, "i.csv")
                open(k_path, "wb").write(kalyan_file.getbuffer())
                open(v_path, "wb").write(varthur_file.getbuffer())
                open(i_path, "wb").write(indranagar_file.getbuffer())
                result = run_full_pipeline(k_path, v_path, i_path, mode=mode)
            
            st.session_state["result"] = result
            st.session_state["mode"] = mode
            st.rerun()
        except Exception as e:
            st.error(f"Processing failed: {e}")
            st.exception(e)
            st.stop()

# ─────────────────────────────────────────────
# Results
# ─────────────────────────────────────────────
if "result" not in st.session_state:
    st.info("Click **Generate Reports** to process the order data.")
    st.stop()

result = st.session_state["result"]
df = result["processed_data"]
pendency = result["pendency"].copy()
breach = result["breach"].copy()
eod = result["eod"].copy()

# ── Inject Rider Count & Avg Order/Rider ──
def add_rider_metrics(report_df):
    report_df = report_df.copy()
    report_df["Rider Count"] = report_df["Store"].map(rider_map).fillna(0).astype(int)
    
    # Prefer Delivered → Total Orders → Total Operational Pending → 0
    def calc_avg(r):
        if r["Rider Count"] <= 0:
            return 0.0
        if "Delivered" in report_df.columns and pd.notna(r.get("Delivered")):
            return round(float(r["Delivered"]) / r["Rider Count"], 1)
        if "Total Orders" in report_df.columns and pd.notna(r.get("Total Orders")):
            return round(float(r["Total Orders"]) / r["Rider Count"], 1)
        if "Total Operational Pending" in report_df.columns and pd.notna(r.get("Total Operational Pending")):
            return round(float(r["Total Operational Pending"]) / r["Rider Count"], 1)
        if "Total Orders (EOD)" in report_df.columns and pd.notna(r.get("Total Orders (EOD)")):
            return round(float(r["Total Orders (EOD)"]) / r["Rider Count"], 1)
        return 0.0
    
    report_df["Avg Order/Rider"] = report_df.apply(calc_avg, axis=1)
    return report_df

breach = add_rider_metrics(breach)
pendency = add_rider_metrics(pendency)
eod = add_rider_metrics(eod)

# ── Top KPI row ──
st.markdown("### Key Performance Indicators")
k1, k2, k3, k4, k5, k6 = st.columns(6)

total_orders = int(eod["Total Orders (EOD)"].sum()) if "Total Orders (EOD)" in eod.columns else int(breach["Total Orders"].sum())
total_delivered = int(eod["Delivered"].sum()) if "Delivered" in eod.columns else int(breach["Delivered"].sum())
total_breach = int(eod["Total Breach"].sum()) if "Total Breach" in eod.columns else int(breach["Total Breach"].sum())
avg_breach = eod["Breach %"].mean() if "Breach %" in eod.columns else breach["Breach %"].mean()
on_time = total_delivered - total_breach
total_riders = sum(rider_map.values())

k1.metric("Total Orders", f"{total_orders:,}")
k2.metric("Delivered", f"{total_delivered:,}")
k3.metric("On-Time", f"{on_time:,}", delta=f"{(on_time/total_delivered*100):.1f}%" if total_delivered else "0%")
k4.metric("Total Breach", f"{total_breach:,}", delta=f"-{avg_breach:.1f}%", delta_color="inverse")
k5.metric("Avg Breach %", f"{avg_breach:.1f}%")
k6.metric("Total Riders", f"{total_riders}")

st.markdown("---")

# ── Tabs ──
tab_eod, tab_pend, tab_breach, tab_breach_detail, tab_detail, tab_raw = st.tabs([
    "EOD Report",
    "Pendency Report",
    "Breach Report",
    "Breach Order IDs + Reason",
    "Order Detail",
    "Downloads"
])

# ═══════════════════════════════════════════════
# TAB 1: EOD
# ═══════════════════════════════════════════════
with tab_eod:
    st.subheader("EOD Consolidated Report")
    st.caption("Same structure as the Excel EOD sheet + Rider metrics")
    
    # Store-wise cards
    stores = eod["Store"].tolist()
    cols = st.columns(len(stores))
    
    for i, store in enumerate(stores):
        row = eod[eod["Store"] == store].iloc[0]
        with cols[i]:
            st.markdown(f"**{store}**")
            st.metric("Orders", int(row.get("Total Orders (EOD)", row.get("Total Orders", 0))))
            st.metric("Delivered", int(row["Delivered"]))
            st.metric("Breach", int(row["Total Breach"]), 
                      delta=f"{row['Breach %']:.1f}%", delta_color="inverse")
            st.metric("Riders", int(row["Rider Count"]))
            st.metric("Avg/Rider", row["Avg Order/Rider"])
    
    st.markdown("#### Full EOD Table")
    display_cols = [c for c in [
        "Store", "Total Orders (EOD)", "Delivered", "On Time Delivered",
        "Total Breach", "Breach %", "Rider Count", "Avg Order/Rider",
        "Attempted Delivery", "Not Dispatched",
        "0 Attempt", "Picked Up", "Intransit", "Cancelled",
        "15 Min Orders", "30 Min Orders", "Other Orders"
    ] if c in eod.columns]
    
    st.dataframe(eod[display_cols], use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════
# TAB 2: PENDENCY
# ═══════════════════════════════════════════════
with tab_pend:
    st.subheader("Pendency Report")
    st.caption("Operational pending states by store (Excel PENDENCY REPORT logic) + Rider metrics. This shows only open / pending orders — not full daily volume.")
    
    # Reorder columns — Store always first
    preferred = ["Store", "Rider Count", "Avg Order/Rider", "Total Operational Pending",
                 "Attempted Delivery", "Not Dispatched", "0 Attempt", "Picked Up", "Intransit", "Cancelled", "0-Attempt %"]
    pend_cols = [c for c in preferred if c in pendency.columns]
    # add any remaining columns
    pend_cols += [c for c in pendency.columns if c not in pend_cols]
    st.dataframe(pendency[pend_cols], use_container_width=True, hide_index=True)
    
    st.markdown("#### Pending Breakdown")
    chart_cols = [c for c in ["Attempted Delivery", "Not Dispatched", "Picked Up", "Intransit"] if c in pendency.columns]
    if chart_cols:
        chart_df = pendency.set_index("Store")[chart_cols]
        st.bar_chart(chart_df, use_container_width=True)

# ═══════════════════════════════════════════════
# TAB 3: BREACH SUMMARY
# ═══════════════════════════════════════════════
with tab_breach:
    st.subheader("Breach Report")
    st.caption("Delivered vs Breach by model (Excel BREACH REPORT logic) + Rider Count.  'Total Orders' = full daily volume (created today + delivered today + still open). This is different from Pendency Total Orders.")
    
    # Put Rider Count and Breach % in nice order
    breach_display_cols = [c for c in [
        "Store", "Total Orders", "Delivered", 
        "15 Min Orders", "30 Min Orders", "Other Orders",
        "Total Breach", "15 Min Breach", "30 Min Breach", "Other Breach",
        "Breach %", "Rider Count", "Avg Order/Rider"
    ] if c in breach.columns]
    
    if breach_display_cols:
        st.dataframe(breach[breach_display_cols], use_container_width=True, hide_index=True)
    else:
        st.dataframe(breach, use_container_width=True, hide_index=True)
    
    st.info("Note: Breach Report 'Total Orders' is the full daily set. Excel Pendency 'Total Orders' is only the current open/pending orders. Both are correct for their purpose — numbers will differ.")
    
    c1, c2 = st.columns(2)
    
    with c1:
        st.markdown("#### Model Mix (Delivered Orders)")
        model_df = breach.set_index("Store")[["15 Min Orders", "30 Min Orders", "Other Orders"]]
        st.bar_chart(model_df, use_container_width=True)
    
    with c2:
        st.markdown("#### Breach by Model")
        breach_model = breach.set_index("Store")[["15 Min Breach", "30 Min Breach", "Other Breach"]]
        st.bar_chart(breach_model, use_container_width=True)
    
    st.markdown("#### Breach % by Store")
    st.bar_chart(breach.set_index("Store")[["Breach %"]], use_container_width=True)

# ═══════════════════════════════════════════════
# TAB 4: BREACH ORDER IDs + REASON (NEW)
# ═══════════════════════════════════════════════
with tab_breach_detail:
    st.subheader("Breach Order IDs + Reason")
    st.caption("All orders where Final Verdict = Yes (SLA Breach). You can also see Rider and Last Failed Remark.")
    
    breach_orders = df[df["Final_Verdict"] == "Yes"].copy()
    
    if len(breach_orders) == 0:
        st.success("No breach orders found for the current filter.")
    else:
        st.markdown(f"**{len(breach_orders)} breach order(s) found**")
        
        # Prepare a clean view
        detail_cols = []
        preferred = [
            "CDR ID", "Reference ID", "Channel", "Customer Name",
            "Model", "New_Model",
            "Order_datetime", "drop_time", "ADT(Actual Delivery Time)",
            "Breach_delay", "Breach_bucket", "SLA_STATUS",
            "Rider Name", "Rider ID",
            "Last Failed Remark", "Failed Attempt Count",
            "Fulfillment Status", "Status",
            "Pickup to Drop Distance(Kms)"
        ]
        for c in preferred:
            if c in breach_orders.columns:
                detail_cols.append(c)
        
        # Sort by Breach_delay descending
        if "Breach_delay" in breach_orders.columns:
            breach_orders = breach_orders.sort_values("Breach_delay", ascending=False)
        
        st.dataframe(
            breach_orders[detail_cols],
            use_container_width=True,
            hide_index=True,
            height=400
        )
        
        # Quick summary by store
        st.markdown("#### Breach Summary by Store")
        summary = breach_orders.groupby("Channel").agg(
            Breach_Count=("CDR ID", "count"),
            Avg_Delay_Mins=("Breach_delay", "mean"),
            Max_Delay_Mins=("Breach_delay", "max")
        ).reset_index()
        summary["Avg_Delay_Mins"] = summary["Avg_Delay_Mins"].round(1)
        st.dataframe(summary, use_container_width=True, hide_index=True)
        
        # Download just the breach orders
        st.download_button(
            "Download Breach Orders Detail (CSV)",
            data=breach_orders[detail_cols].to_csv(index=False).encode("utf-8"),
            file_name=f"ST_Breach_Orders_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )

# ═══════════════════════════════════════════════
# TAB 5: ORDER DETAIL
# ═══════════════════════════════════════════════
with tab_detail:
    st.subheader("Processed Order Detail")
    st.caption("All PIDGE calculated columns applied — filter as needed")
    
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        ch_opts = sorted(df["Channel"].dropna().unique())
        ch_sel = st.multiselect("Channel", ch_opts, default=ch_opts)
    with f2:
        model_opts = sorted(df["Model"].dropna().unique())
        model_sel = st.multiselect("Model", model_opts, default=model_opts)
    with f3:
        sla_opts = sorted(df["SLA_STATUS"].dropna().unique())
        sla_sel = st.multiselect("SLA Status", sla_opts, default=sla_opts)
    with f4:
        verdict_opts = sorted(df["Final_Verdict"].dropna().unique())
        verdict_sel = st.multiselect("Final Verdict", verdict_opts, default=verdict_opts)
    
    filtered = df[
        df["Channel"].isin(ch_sel) &
        df["Model"].isin(model_sel) &
        df["SLA_STATUS"].isin(sla_sel) &
        df["Final_Verdict"].isin(verdict_sel)
    ]
    
    st.markdown(f"**Showing {len(filtered):,} of {len(df):,} orders**")
    
    show_cols = [c for c in [
        "CDR ID", "Reference ID", "Channel", "Customer Name", "Status", "Fulfillment Status",
        "Model", "New_Model", "Order_datetime", "drop_time", "Promise_time",
        "TAT", "ADT(Actual Delivery Time)", "SLA_STATUS",
        "Breach_delay", "Breach_bucket", "Final_Verdict",
        "Rider Name", "Last Failed Remark",
        "Warehouse_pincode", "Destination_pincode", "Pickup to Drop Distance(Kms)"
    ] if c in filtered.columns]
    
    st.dataframe(filtered[show_cols], use_container_width=True, height=450)

# ═══════════════════════════════════════════════
# TAB 6: DOWNLOADS
# ═══════════════════════════════════════════════
with tab_raw:
    st.subheader("Download Reports")
    
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    
    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.markdown("**EOD Report**")
        st.download_button(
            "Download EOD CSV",
            data=eod.to_csv(index=False).encode("utf-8"),
            file_name=f"ST_EOD_{ts}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with c2:
        st.markdown("**Pendency Report**")
        st.download_button(
            "Download Pendency CSV",
            data=pendency.to_csv(index=False).encode("utf-8"),
            file_name=f"ST_Pendency_{ts}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with c3:
        st.markdown("**Breach Report**")
        st.download_button(
            "Download Breach CSV",
            data=breach.to_csv(index=False).encode("utf-8"),
            file_name=f"ST_Breach_{ts}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    st.markdown("---")
    st.markdown("**Full Processed Data** (all PIDGE columns)")
    st.download_button(
        "Download Full Processed Data (CSV)",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name=f"ST_Processed_Full_{ts}.csv",
        mime="text/csv",
        use_container_width=True
    )
    
    st.markdown("---")
    st.info(
        "All calculations follow the Excel SOP formulas exactly "
        "(Model, TAT, SLA STATUS, Breach Delay, Final Verdict, etc.). "
        "Rider Count is entered manually in the sidebar."
    )

# Footer
st.markdown("---")
st.caption(
    f"ST Dashboard · Mode: **{st.session_state.get('mode', 'daily')}** · "
    f"Processed {len(df):,} orders · Rider Counts: Kalyan={rider_kalyan}, Indra={rider_indra}, Varthur={rider_varthur}"
)
