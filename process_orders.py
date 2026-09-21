"""
ST Dashboard Automation Engine
Replicates the Excel SOP logic (PIDGE calculations + reports)
+ Handwritten daily cleaning rules
"""

import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import re
from pathlib import Path

# ============================================================
# 1. LOAD & BASIC CLEANING (from handwritten SOP)
# ============================================================

def load_and_clean_csv(filepath: str, store_label: str = None) -> pd.DataFrame:
    """Load one store CSV and apply handwritten cleaning steps."""
    df = pd.read_csv(filepath, dtype=str)
    
    # STEP: Delete Movement Type (column E)
    if "Movement Type" in df.columns:
        df = df.drop(columns=["Movement Type"])
    
    # STEP: Delete Cancellation Source → Cancellation Reference ID (AJ-AM equivalent)
    cancel_cols = [
        "Cancellation Source",
        "Cancellation Reason",
        "Cancellation At Status",
        "Cancellation Reference ID"
    ]
    for c in cancel_cols:
        if c in df.columns:
            df = df.drop(columns=[c])
    
    # Standardize column names (keep original for mapping)
    df.columns = [c.strip() for c in df.columns]
    
    # Parse key datetime columns
    datetime_cols = [
        "Creation Date", "Manifested Time", "Out For Pickup Time",
        "Reached Pickup Time", "Pickup Time", "Out For Delivery Time",
        "Reached Delivery Time", "Terminal Time",
        "EDT(Expected Delivery Time)", "ADT(Actual Delivery Time)",
        "Cancellation Date", "First Allocation Request", "Fulfillment Request Time"
    ]
    for col in datetime_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    
    # Numeric
    for col in ["Bill Amount", "COD Amount", "Total Weight", "Total Volume",
                "Pickup to Drop Distance(Kms)", "Failed Attempt Count",
                "Failed Partner Allocations", "Owner ID"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    
    # Add store label if needed
    if store_label:
        df["Store_Label"] = store_label
    
    return df


def apply_date_filters(df: pd.DataFrame, report_date: datetime = None, mode: str = "all") -> pd.DataFrame:
    """
    Apply the handwritten date filtering logic.
    
    mode="all"     → keep full history (useful for sample / historical analysis)
    mode="daily"   → keep only orders whose promised drop_time falls on report_date
                     (excludes next-day / future SDD and previous-day leftovers from Total Orders & Pendency)
    """
    df = df.copy()
    
    # Always remove RTO
    if "Last Failed Remark" in df.columns:
        rto_mask = df["Last Failed Remark"].fillna("").str.contains("RTO", case=False, na=False)
        df = df[~rto_mask]
    if "Fulfillment Status" in df.columns:
        df = df[df["Fulfillment Status"].fillna("") != "RTO_DELIVERED"]
    
    if mode == "all":
        return df.reset_index(drop=True)
    
    # daily mode
    if report_date is None:
        report_date = df["Creation Date"].max()
    
    # Ensure today is a normalized Timestamp
    today = pd.Timestamp(report_date).normalize()
    
    # Extract drop_time from Notes early so we can filter strictly by promised day
    if "Notes" in df.columns:
        parsed = df["Notes"].apply(parse_notes_timestamps)
        df["_drop_time_tmp"] = [p[1] for p in parsed]
    else:
        df["_drop_time_tmp"] = pd.NaT
    
    # Keep ONLY orders whose promised drop_time is on the report date.
    # This excludes:
    #   - future SDD / scheduled orders (drop_time next day or later)
    #   - previous-day undelivered leftovers
    #   - historical completed deliveries whose Notes still carry a future drop_time
    mask_drop_today = df["_drop_time_tmp"].dt.normalize() == today
    # Exclude already-delivered orders whose ADT is NOT on the report date
    has_adt = df["ADT(Actual Delivery Time)"].notna()
    adt_not_today = has_adt & (df["ADT(Actual Delivery Time)"].dt.normalize() != today)
    mask_drop_today = mask_drop_today & ~adt_not_today
    # Fallback for rows missing Notes/drop_time: still keep created-today or delivered-today
    mask_fallback = (
        df["_drop_time_tmp"].isna() &
        (
            (df["Creation Date"].dt.normalize() == today) |
            (df["ADT(Actual Delivery Time)"].dt.normalize() == today)
        )
    )
    df = df[mask_drop_today | mask_fallback].copy()
    
    # Cancelled previous date
    if "Cancellation Date" in df.columns and "Status" in df.columns:
        cancelled_prev = (
            (df["Status"].str.lower() == "cancelled") &
            (df["Cancellation Date"].notna()) &
            (df["Cancellation Date"].dt.normalize() < today)
        )
        df = df[~cancelled_prev]
    
    # Clean temporary column
    df = df.drop(columns=["_drop_time_tmp"], errors="ignore")
    
    return df.reset_index(drop=True)


# ============================================================
# 2. PIDGE CALCULATED COLUMNS (exact logic from SOP)
# ============================================================

def parse_notes_timestamps(notes: str):
    """Extract Order Create Time Stamp and drop_time from Notes field."""
    if pd.isna(notes) or not isinstance(notes, str):
        return None, None
    create_ts = None
    drop_ts = None
    # Order Create Time Stamp: 2026-09-18 10:26:18
    m1 = re.search(r"Order Create Time Stamp:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})", notes)
    if m1:
        try:
            create_ts = pd.to_datetime(m1.group(1))
        except:
            pass
    # drop_time: 2026-09-18 10:58:31
    m2 = re.search(r"drop_time:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})", notes)
    if m2:
        try:
            drop_ts = pd.to_datetime(m2.group(1))
        except:
            pass
    return create_ts, drop_ts


def classify_model(row) -> str:
    """AV2 formula — Model classification."""
    at = str(row.get("Warehouse_pincode", ""))[-6:] if pd.notna(row.get("Warehouse_pincode")) else ""
    au = str(row.get("Destination_pincode", ""))[-6:] if pd.notna(row.get("Destination_pincode")) else ""
    channel = str(row.get("Channel", ""))
    g = str(row.get("Sender Name", ""))
    az = row.get("Order_datetime")
    bh = row.get("Promise_time")  # duration
    
    if at == "562162":
        return "Warehouse Order"
    
    valid_channels = ["Supertails-Varthur", "Supertail_Hyperlocal", "Supertails-Indranagar"]
    valid_senders = ["DKS_Blr_varthur_002", "DKS_Blr_kalyannagar_002", "DKS_Blr_kaggdaspura_002"]
    
    if channel not in valid_channels or au == "560049" or g not in valid_senders:
        return "SDD Delivery"
    
    if pd.isna(az):
        return "SDD Delivery"
    
    hour_az = az.hour if hasattr(az, "hour") else 0
    hour_bh = 0
    minute_bh = 0
    if pd.notna(bh) and hasattr(bh, "total_seconds"):
        total_sec = bh.total_seconds()
        hour_bh = int(total_sec // 3600)
        minute_bh = int((total_sec % 3600) // 60)
    elif pd.notna(bh):
        try:
            hour_bh = int(bh)
        except:
            pass
    
    # 15 Min Delivery
    if (7 <= hour_az < 22 and hour_bh < 1 and minute_bh <= 15):
        return "15 Min Delivery"
    # 30 Min Delivery
    if (7 <= hour_az < 22 and hour_bh < 1):
        return "30 Min Delivery"
    # 2 Hours Delivery
    if (8 <= hour_az < 20 and hour_bh == 2):
        return "2 Hours Delivery"
    # Overnight
    if (hour_bh < 1 and (hour_az >= 22 or hour_az < 7)):
        return "Overnight order"
    
    return "SDD Delivery"


def calculate_tat(row) -> pd.Timestamp:
    """AW2 formula — TAT / SLA deadline."""
    channel = str(row.get("Channel", ""))
    av = str(row.get("Model", ""))
    az = row.get("Order_datetime")
    
    if pd.isna(az):
        return pd.NaT
    
    # Gurgaon special (not used in current 3 stores but kept)
    if channel == "Supertails-Gurgaon":
        if az.hour < 14:
            return az.normalize() + pd.Timedelta(hours=23, minutes=59, seconds=59)
        else:
            return (az.normalize() + pd.Timedelta(days=1)) + pd.Timedelta(hours=23, minutes=59, seconds=59)
    
    if av == "SDD Delivery":
        if channel in ["Supertails-Varthur", "Supertail_Hyperlocal"]:
            return (az.normalize() + pd.Timedelta(days=1)) + pd.Timedelta(hours=13)
        if channel == "Supertails-Indranagar":
            return (az.normalize() + pd.Timedelta(days=1)) + pd.Timedelta(hours=11)
    
    if av == "Overnight order":
        if channel in ["Supertail_Hyperlocal", "Supertails-Varthur", "Supertails-Indranagar"]:
            if az.hour >= 22:
                return (az.normalize() + pd.Timedelta(days=1)) + pd.Timedelta(hours=9, minutes=1)
            else:
                return az.normalize() + pd.Timedelta(hours=9, minutes=1)
    
    if av == "2 Hours Delivery":
        return az + pd.Timedelta(hours=2)
    
    if av == "30 Min Delivery":
        return az + pd.Timedelta(minutes=31)
    
    if av == "15 Min Delivery":
        if channel in ["Supertail_Hyperlocal", "Supertails-Varthur"]:
            return az + pd.Timedelta(minutes=15)
        # Indranagar 15 min not explicitly in formula for AW, falls through
    
    if av == "Warehouse Order":
        if az.hour < 11:
            return az.normalize() + pd.Timedelta(hours=23, minutes=59, seconds=59)
        else:
            return (az.normalize() + pd.Timedelta(days=1)) + pd.Timedelta(hours=23, minutes=59, seconds=59)
    
    return pd.NaT


def add_pidge_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Add all calculated columns AT → BV following the SOP formulas."""
    df = df.copy()
    
    # AT Warehouse pincode = RIGHT(Sender Address, 6)
    df["Warehouse_pincode"] = df["Sender Address"].fillna("").astype(str).str[-6:]
    
    # AU Destination pincode
    df["Destination_pincode"] = df["Customer Address"].fillna("").astype(str).str[-6:]
    
    # Parse Notes → AZ (Order_datetime), BA (drop_time)
    parsed = df["Notes"].apply(parse_notes_timestamps)
    df["Order_datetime"] = [p[0] for p in parsed]
    df["drop_time"] = [p[1] for p in parsed]
    
    # Fallback: if parse fails, use Creation Date
    df["Order_datetime"] = df["Order_datetime"].fillna(df["Creation Date"])
    
    # BH Promise time = drop_time - Order_datetime
    df["Promise_time"] = df["drop_time"] - df["Order_datetime"]
    
    # AV Model
    df["Model"] = df.apply(classify_model, axis=1)
    
    # AW TAT
    df["TAT"] = df.apply(calculate_tat, axis=1)
    
    # AX SLA STATUS
    def sla_status(row):
        adt = row.get("ADT(Actual Delivery Time)")
        tat = row.get("TAT")
        if pd.isna(adt):
            return "undelivered"
        if pd.isna(tat):
            return "undelivered"
        if adt > tat:
            return "SLA BREACH"
        return "SLA MET"
    df["SLA_STATUS"] = df.apply(sla_status, axis=1)
    
    # AY Time portion
    df["Time_portion"] = df["Order_datetime"].apply(lambda x: x - x.normalize() if pd.notna(x) else pd.NaT)
    
    # BB order_date
    df["order_date"] = df["Order_datetime"].dt.strftime("%d-%b")
    
    # BC Packing_Time = Creation Date - Order_datetime
    df["Packing_Time"] = df["Creation Date"] - df["Order_datetime"]
    
    # BD Packing_delay
    def packing_delay(td):
        if pd.isna(td):
            return ""
        mins = td.total_seconds() / 60
        if mins <= 15:
            return "15 min delay"
        if mins <= 30:
            return "15-30 min delay"
        return "More than 30 min delay"
    df["Packing_delay"] = df["Packing_Time"].apply(packing_delay)
    
    # BE Delay_bucket (finer)
    def delay_bucket(td):
        if pd.isna(td):
            return ""
        mins = td.total_seconds() / 60
        if mins <= 5:
            return "0–5 min delay"
        if mins <= 10:
            return "5–10 min delay"
        if mins <= 15:
            return "10–15 min delay"
        return "More than 15 min delay"
    df["Delay_bucket"] = df["Packing_Time"].apply(delay_bucket)
    
    # BF New Model (normalized)
    def new_model(row):
        av = row["Model"]
        bh = row["Promise_time"]
        if av in ["SDD Delivery", "Overnight order"]:
            return av
        if pd.isna(bh):
            return av
        total_min = bh.total_seconds() / 60 if hasattr(bh, "total_seconds") else 0
        if total_min <= 15:
            return "15 Min Delivery"
        if total_min <= 60:
            return "30 Min Delivery"
        return av
    df["New_Model"] = df.apply(new_model, axis=1)
    
    # BG = AD (distance)
    df["Distance_km"] = df["Pickup to Drop Distance(Kms)"]
    
    # BI / BJ / BP store-specific models
    def kalyan_model(row):
        if row["Channel"] != "Supertail_Hyperlocal":
            return row["New_Model"]
        bh = row["Promise_time"]
        if pd.isna(bh):
            return row["New_Model"]
        mins = bh.total_seconds() / 60
        if mins == 30:
            return "30 Min Order"
        if 50 <= mins <= 70:  # ~1 hour
            return "60 Min Order"
        if 110 <= mins <= 130:
            return "2 hour order"
        return row["New_Model"]
    df["Kalyan_Model"] = df.apply(kalyan_model, axis=1)
    
    def varthur_model(row):
        if row["Channel"] != "Supertails-Varthur":
            return row["New_Model"]
        bh = row["Promise_time"]
        if pd.isna(bh):
            return row["New_Model"]
        mins = bh.total_seconds() / 60
        if mins == 30:
            return "30 Min Order"
        if 50 <= mins <= 70:
            return "60 Min Order"
        if 110 <= mins <= 130:
            return "2 hour order"
        return row["New_Model"]
    df["Varthur_Model"] = df.apply(varthur_model, axis=1)
    
    def indranagar_model(row):
        if row["Channel"] != "Supertails-Indranagar":
            return row["New_Model"]
        bh = row["Promise_time"]
        if pd.isna(bh):
            return row["New_Model"]
        mins = bh.total_seconds() / 60
        if mins == 30:
            return "30 Min Order"
        if 50 <= mins <= 70:
            return "60 Min Order"
        return row["New_Model"]
    df["Indranagar_Model"] = df.apply(indranagar_model, axis=1)
    
    # BM = ADT if present else Status
    df["BM_ADT"] = df["ADT(Actual Delivery Time)"].fillna(df["Status"])
    
    # BL = drop_time (BA)
    df["BL_drop"] = df["drop_time"]
    
    # BR Breach_delay (minutes) = ROUNDUP((BM - BL)*24*60)
    def breach_delay(row):
        bm = row["ADT(Actual Delivery Time)"]
        bl = row["drop_time"]
        if pd.isna(bm) or pd.isna(bl):
            return np.nan
        delta = (bm - bl).total_seconds() / 60
        return np.ceil(delta) if delta > 0 else 0
    df["Breach_delay"] = df.apply(breach_delay, axis=1)
    
    # BS Breach_bucket
    def breach_bucket(val):
        if pd.isna(val):
            return "undelivered"
        if val <= 0:
            return "on time"
        if val <= 5:
            return "0-5 mins delay"
        if val <= 10:
            return "5-10 mins delay"
        if val <= 15:
            return "10-15 mins delay"
        if val <= 20:
            return "15-20 mins delay"
        if val <= 25:
            return "20-25 mins delay"
        if val <= 30:
            return "25-30 mins delay"
        return "more than 30 mins delay"
    df["Breach_bucket"] = df["Breach_delay"].apply(breach_bucket)
    
    # BT Breach acc. Delay
    def bt_flag(val):
        if pd.isna(val):
            return "NA"
        if val < 2:
            return "no"
        return "yes"
    df["BT_flag"] = df["Breach_delay"].apply(bt_flag)
    
    # BU Breach acc. Bucket
    def bu_flag(bucket):
        if bucket == "undelivered":
            return "NA"
        if bucket == "on time":
            return "no"
        return "yes"
    df["BU_flag"] = df["Breach_bucket"].apply(bu_flag)
    
    # BV Final Verdict
    df["Final_Verdict"] = np.where(
        (df["BT_flag"] == "yes") & (df["BU_flag"] == "yes") & (df["SLA_STATUS"] == "SLA BREACH"),
        "Yes", "No"
    )
    
    return df


# ============================================================
# 3. REPORT GENERATORS
# ============================================================

def generate_pendency_report(df: pd.DataFrame) -> pd.DataFrame:
    """PENDENCY REPORT logic."""
    stores = {
        "Indranagar": "Supertails-Indranagar",
        "Kalyan Nagar": "Supertail_Hyperlocal",
        "Varthur": "Supertails-Varthur"
    }
    
    rows = []
    for label, channel in stores.items():
        sub = df[df["Channel"] == channel]
        
        attempted = len(sub[sub["Fulfillment Status"] == "UNDELIVERED"])
        not_dispatched = len(sub[
            (sub["Fulfillment Status"] == "CREATED") |
            (sub["Status"] == "pending")
        ])
        zero_attempt = len(sub[
            ((sub["Fulfillment Status"] == "CREATED") | (sub["Status"] == "pending")) &
            (sub["Failed Attempt Count"].fillna(0) == 0)
        ])
        picked_up = len(sub[sub["Fulfillment Status"].isin(["PICKED_UP", "REACHED_DELIVERY", "REACHED_PICKUP"])])
        intransit = len(sub[sub["Fulfillment Status"].isin(["OUT_FOR_DELIVERY", "IN_TRANSIT"])])
        cancelled = len(sub[sub["Status"].str.lower() == "cancelled"])
        
        total_ops = attempted + not_dispatched + picked_up + intransit
        
        rows.append({
            "Store": label,
            "Attempted Delivery": attempted,
            "Not Dispatched": not_dispatched,
            "0 Attempt": zero_attempt,
            "Picked Up": picked_up,
            "Intransit": intransit,
            "Cancelled": cancelled,
            "Total Operational Pending": total_ops,
            "0-Attempt %": round(zero_attempt / total_ops * 100, 1) if total_ops else 0
        })
    
    return pd.DataFrame(rows)


def generate_breach_report(df: pd.DataFrame) -> pd.DataFrame:
    """BREACH REPORT logic."""
    stores = {
        "Kalyan Nagar": "Supertail_Hyperlocal",
        "Indranagar": "Supertails-Indranagar",
        "Varthur": "Supertails-Varthur"
    }
    
    rows = []
    for label, channel in stores.items():
        sub = df[df["Channel"] == channel]
        # Exclude RTO
        sub = sub[sub["Fulfillment Status"].fillna("") != "RTO_DELIVERED"]
        
        total_orders = len(sub)
        delivered = sub[sub["ADT(Actual Delivery Time)"].notna()]
        delivered_count = len(delivered)
        
        min15 = len(delivered[delivered["Model"] == "15 Min Delivery"])
        min30 = len(delivered[delivered["Model"] == "30 Min Delivery"])
        other = delivered_count - min15 - min30
        
        breach = delivered[delivered["Final_Verdict"] == "Yes"]
        breach_count = len(breach)
        b15 = len(breach[breach["Model"] == "15 Min Delivery"])
        b30 = len(breach[breach["Model"] == "30 Min Delivery"])
        b_other = breach_count - b15 - b30
        
        breach_pct = round(breach_count / delivered_count * 100, 1) if delivered_count else 0
        
        rows.append({
            "Store": label,
            "Total Orders": total_orders,
            "Delivered": delivered_count,
            "15 Min Orders": min15,
            "30 Min Orders": min30,
            "Other Orders": other,
            "Total Breach": breach_count,
            "15 Min Breach": b15,
            "30 Min Breach": b30,
            "Other Breach": b_other,
            "Breach %": breach_pct
        })
    
    return pd.DataFrame(rows)


def generate_eod_summary(pendency: pd.DataFrame, breach: pd.DataFrame) -> pd.DataFrame:
    """Simple EOD consolidated view."""
    # Merge on Store
    eod = pendency.merge(breach, on="Store", how="outer")
    eod["On Time Delivered"] = eod["Delivered"] - eod["Total Breach"]
    eod["Total Orders (EOD)"] = (
        eod["Attempted Delivery"] + eod["Not Dispatched"] +
        eod["Picked Up"] + eod["Intransit"] + eod["Cancelled"] + eod["Delivered"]
    )
    return eod


# ============================================================
# 4. MAIN PIPELINE
# ============================================================

def run_full_pipeline(
    kalyan_path: str,
    varthur_path: str,
    indranagar_path: str,
    report_date: datetime = None,
    mode: str = "all"
):
    """Load → Clean → Calculate → Reports.
    
    mode: "all" (full history) or "daily" (report_date only + open orders)
    """
    
    # Load
    df_k = load_and_clean_csv(kalyan_path, "Kalyan Nagar")
    df_v = load_and_clean_csv(varthur_path, "Varthur")
    df_i = load_and_clean_csv(indranagar_path, "Indranagar")
    
    # Combine
    df = pd.concat([df_k, df_v, df_i], ignore_index=True)
    
    # Date filters (handwritten SOP)
    df = apply_date_filters(df, report_date, mode=mode)
    
    # PIDGE calculations
    df = add_pidge_columns(df)
    
    # Reports
    pendency = generate_pendency_report(df)
    breach = generate_breach_report(df)
    eod = generate_eod_summary(pendency, breach)
    
    return {
        "processed_data": df,
        "pendency": pendency,
        "breach": breach,
        "eod": eod
    }


if __name__ == "__main__":
    base = Path("/home/workdir/attachments")
    result = run_full_pipeline(
        str(base / "orders-default,performance-25071.csv"),
        str(base / "orders-default,performance-25072.csv"),
        str(base / "orders-default,performance-25073.csv"),
        mode="all"
    )
    
    print("=== PENDENCY REPORT ===")
    print(result["pendency"].to_string(index=False))
    print("\n=== BREACH REPORT ===")
    print(result["breach"].to_string(index=False))
    print("\n=== EOD SUMMARY ===")
    print(result["eod"][["Store", "Total Orders (EOD)", "Delivered", "On Time Delivered", "Total Breach", "Breach %"]].to_string(index=False))
    print(f"\nProcessed rows: {len(result['processed_data'])}")
    print("\nModel distribution:")
    print(result["processed_data"]["Model"].value_counts())
    print("\nSLA_STATUS distribution:")
    print(result["processed_data"]["SLA_STATUS"].value_counts())
    print("\nFinal_Verdict distribution:")
    print(result["processed_data"]["Final_Verdict"].value_counts())