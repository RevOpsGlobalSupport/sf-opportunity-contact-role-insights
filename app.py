import streamlit as st
import pandas as pd
from datetime import datetime

# -----------------------
# Robust CSV loading
# -----------------------
def load_csv(file):
    """
    Load a CSV robustly:
    - Tries multiple encodings typical of Salesforce exports
    - Rewinds file between attempts
    - Final fallback replaces invalid bytes so app never crashes
    """
    for enc in ["utf-8-sig", "utf-8", "latin1", "cp1252"]:
        try:
            file.seek(0)
            return pd.read_csv(file, encoding=enc)
        except UnicodeDecodeError:
            continue
        except Exception:
            continue

    file.seek(0)
    return pd.read_csv(file, encoding="latin1", errors="replace")


# -----------------------
# Column normalization
# -----------------------
def normalize_and_standardize_columns(df, is_roles=False):
    """
    Normalize column names (trim, lower) and map common Salesforce variants
    into a standard schema used by the app.
    """
    original_cols = list(df.columns)
    col_map = {}

    for col in original_cols:
        norm = col.strip().lower()

        # base opportunity fields
        if norm == "opportunity id":
            col_map[col] = "Opportunity ID"
        elif norm == "opportunity name":
            col_map[col] = "Opportunity Name"
        elif norm == "account id":
            col_map[col] = "Account ID"
        elif norm == "amount":
            col_map[col] = "Amount"
        elif norm == "type":
            col_map[col] = "Type"
        elif norm == "stage":
            col_map[col] = "Stage"
        elif norm in ("created date", "opportunity created date", "opportunity created"):
            col_map[col] = "Created Date"
        elif norm in ("closed date", "close date", "opportunity close date", "opportunity closed date"):
            col_map[col] = "Close Date"
        elif norm in ("opportunity owner", "owner", "owner name", "opportunity owner name"):
            col_map[col] = "Opportunity Owner"

        # contact role fields
        elif is_roles and "contact id" in norm:
            col_map[col] = "Contact ID"
        elif is_roles and norm == "title":
            col_map[col] = "Title"
        elif is_roles and ("contact role" in norm or norm == "role"):
            col_map[col] = "Contact Role"
        elif is_roles and ("primary" in norm or "is primary" in norm):
            col_map[col] = "Primary"

        else:
            col_map[col] = col

    return df.rename(columns=col_map)


def parse_date(x):
    try:
        return pd.to_datetime(x)
    except Exception:
        return pd.NaT


# -----------------------
# Streamlit App
# -----------------------
st.set_page_config(page_title="SFDC Opportunity Contact Role Insights", layout="wide")
st.title("Salesforce Opportunity Contact Role Insights")

st.markdown(
    """
Upload two Salesforce exports in CSV format:

1. **Opportunities** report  
2. **Opportunities with Contact Roles** report  

The app normalizes common Salesforce header variations  
(e.g., **Close Date vs Closed Date**, **Opportunity Created Date vs Created Date**).
"""
)

opps_file = st.file_uploader("Upload Opportunities CSV", type=["csv"], key="opps")
roles_file = st.file_uploader("Upload Opportunities with Contact Roles CSV", type=["csv"], key="roles")

if opps_file and roles_file:
    try:
        raw_opps = load_csv(opps_file)
        raw_roles = load_csv(roles_file)
    except Exception as e:
        st.error(f"Error reading CSV files. Please confirm they are valid CSV exports. Details: {e}")
        st.stop()

    # normalize headers
    opps = normalize_and_standardize_columns(raw_opps, is_roles=False)
    roles = normalize_and_standardize_columns(raw_roles, is_roles=True)

    # Required columns for calculations
    required_opps = [
        "Opportunity ID", "Opportunity Name", "Account ID", "Amount",
        "Type", "Stage", "Created Date", "Close Date", "Opportunity Owner"
    ]
    required_roles = ["Opportunity ID", "Contact ID"]

    missing_opps = [c for c in required_opps if c not in opps.columns]
    missing_roles = [c for c in required_roles if c not in roles.columns]

    if missing_opps:
        st.error(
            "The Opportunities file is missing required columns after normalization: "
            + ", ".join(missing_opps)
            + ". Please re-export using the Salesforce template."
        )
        st.stop()

    if missing_roles:
        st.error(
            "The Opportunities with Contact Roles file is missing required columns after normalization: "
            + ", ".join(missing_roles)
            + ". Please re-export using the Salesforce template."
        )
        st.stop()

    # -----------------------
    # Cleaning & preparation
    # -----------------------
    opps["Amount"] = pd.to_numeric(opps["Amount"], errors="coerce").fillna(0)
    if "Amount" in roles.columns:
        roles["Amount"] = pd.to_numeric(roles["Amount"], errors="coerce").fillna(0)

    opps["Created Date"] = opps["Created Date"].apply(parse_date)
    opps["Close Date"] = opps["Close Date"].apply(parse_date)

    stage = opps["Stage"].fillna("")
    won_mask = stage.str.contains("Won", case=False, na=False)
    lost_mask = stage.str.contains("Lost", case=False, na=False)
    open_mask = ~won_mask & ~lost_mask

    # contact counts per opp
    cr_counts = roles.groupby("Opportunity ID")["Contact ID"].nunique()
    opps = opps.merge(cr_counts.rename("contact_count"), on="Opportunity ID", how="left")
    opps["contact_count"] = opps["contact_count"].fillna(0)

    # -----------------------
    # Core KPIs
    # -----------------------
    total_opps = opps["Opportunity ID"].nunique()
    total_pipeline = opps["Amount"].sum()

    opps_with_cr_ids = roles["Opportunity ID"].dropna().unique()
    opps_with_cr = opps[opps["Opportunity ID"].isin(opps_with_cr_ids)]["Opportunity ID"].nunique()
    opps_without_cr = total_opps - opps_with_cr

    pipeline_with_cr = opps[opps["Opportunity ID"].isin(opps_with_cr_ids)]["Amount"].sum()
    pipeline_without_cr = total_pipeline - pipeline_with_cr

    one_cr_ids = cr_counts[cr_counts == 1].index
    opps_one_cr = opps[opps["Opportunity ID"].isin(one_cr_ids)]["Opportunity ID"].nunique()
    pipeline_one_cr = opps[opps["Opportunity ID"].isin(one_cr_ids)]["Amount"].sum()

    open_opps = opps[open_mask].copy()
    won_opps = opps[won_mask].copy()
    lost_opps = opps[lost_mask].copy()

    won_count = won_opps["Opportunity ID"].nunique()
    lost_count = lost_opps["Opportunity ID"].nunique()
    win_rate = won_count / (won_count + lost_count) if (won_count + lost_count) > 0 else 0

    # -----------------------
    # Contact role averages
    # -----------------------
    avg_cr_lost = lost_opps["contact_count"].mean() if not lost_opps.empty else 0
    avg_cr_won = won_opps["contact_count"].mean() if not won_opps.empty else 0
    avg_cr_open = open_opps["contact_count"].mean() if not open_opps.empty else 0

    # -----------------------
    # Time metrics
    # -----------------------
    def days_diff(row):
        if pd.isna(row["Created Date"]) or pd.isna(row["Close Date"]):
            return None
        return (row["Close Date"] - row["Created Date"]).days

    if not lost_opps.empty:
        lost_opps["days_to_close"] = lost_opps.apply(days_diff, axis=1)
    if not won_opps.empty:
        won_opps["days_to_close"] = won_opps.apply(days_diff, axis=1)

    avg_days_lost = (
        lost_opps["days_to_close"].dropna().mean()
        if "days_to_close" in lost_opps and not lost_opps["days_to_close"].dropna().empty
        else None
    )
    avg_days_won = (
        won_opps["days_to_close"].dropna().mean()
        if "days_to_close" in won_opps and not won_opps["days_to_close"].dropna().empty
        else None
    )

    today = pd.Timestamp.today().normalize()
    if not open_opps.empty:
        open_opps["age_days"] = (today - open_opps["Created Date"]).dt.days
        avg_age_open = (
            open_opps["age_days"].dropna().mean()
            if not open_opps["age_days"].dropna().empty
            else None
        )
    else:
        avg_age_open = None

    # -----------------------
    # Simple uplift model
    # -----------------------
    industry_cr_open = 2.0

    if avg_cr_lost not in (0, None) and avg_cr_won not in (0, None):
        contact_influence_ratio = avg_cr_won / avg_cr_lost
    else:
        contact_influence_ratio = 1

    enhanced_win_rate = win_rate * contact_influence_ratio
    enhanced_win_rate = min(max(enhanced_win_rate, win_rate), 0.95)

    open_pipeline = open_opps["Amount"].sum() if not open_opps.empty else 0
    incremental_won_pipeline = max(0, (enhanced_win_rate - win_rate) * open_pipeline)

    # -----------------------
    # Layout + descriptions
    # -----------------------
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Core Metrics")

        st.metric("Total Opportunities", f"{total_opps:,}")
        st.caption("Unique Opportunity IDs in the Opportunities export.")

        st.metric("Total Pipeline", f"${total_pipeline:,.0f}")
        st.caption("Sum of Amount for all opportunities.")

        st.metric("Current Win Rate", f"{win_rate:.1%}")
        st.caption("Closed Won ÷ (Closed Won + Closed Lost), based on Stage containing Won/Lost.")

        st.subheader("Contact Role Coverage")

        st.write(f"Opportunities **with** Contact Roles: **{opps_with_cr:,}**")
        st.caption("Unique opportunities that appear in the Contact Roles export.")

        st.write(f"Opportunities **without** Contact Roles: **{opps_without_cr:,}**")
        st.caption("Total Opportunities minus those with Contact Roles.")

        st.write(f"Pipeline with Contact Roles: **${pipeline_with_cr:,.0f}**")
        st.caption("Sum of Amount for opportunities that have ≥1 Contact Role.")

        st.write(f"Pipeline without Contact Roles: **${pipeline_without_cr:,.0f}**")
        st.caption("Sum of Amount for opportunities with 0 Contact Roles.")

        st.write(f"Opps with **only 1** Contact Role: **{opps_one_cr:,}**")
        st.caption("Unique opportunities where Contact Role count = 1.")

        st.write(f"Pipeline with only 1 Contact Role: **${pipeline_one_cr:,.0f}**")
        st.caption("Sum of Amount for opportunities with exactly 1 Contact Role.")

    with col2:
        st.subheader("Contact Roles by Outcome")

        st.write(f"Avg Contact Roles – **Won**: **{avg_cr_won:.1f}**")
        st.caption("Average Contact Role count per Opportunity where Stage contains Won.")

        st.write(f"Avg Contact Roles – **Lost**: **{avg_cr_lost:.1f}**")
        st.caption("Average Contact Role count per Opportunity where Stage contains Lost.")

        st.write(f"Avg Contact Roles – **Open**: **{avg_cr_open:.1f}**")
        st.caption("Average Contact Role count per Opportunity where Stage does not contain Won or Lost.")

        st.subheader("Time to Close")

        if avg_days_won is not None:
            st.write(f"Avg days to close – **Won**: **{avg_days_won:.0f}**")
            st.caption("Average (Close Date − Created Date) for Won opportunities.")

        if avg_days_lost is not None:
            st.write(f"Avg days to close – **Lost**: **{avg_days_lost:.0f}**")
            st.caption("Average (Close Date − Created Date) for Lost opportunities.")

        if avg_age_open is not None:
            st.write(f"Avg age of **Open** opps: **{avg_age_open:.0f}**")
            st.caption("Average (Today − Created Date) for open opportunities.")

        st.subheader("Modeled Uplift")

        st.write(f"Contact Influence Ratio (Won vs Lost): **{contact_influence_ratio:.2f}×**")
        st.caption("Avg Contact Roles on Won ÷ Avg Contact Roles on Lost. Indicates how strongly contacts correlate with winning.")

        st.write(f"Enhanced Win Rate (modeled): **{enhanced_win_rate:.1%}**")
        st.caption("Current Win Rate scaled by Contact Influence Ratio, capped at 95% for sanity.")

        st.write(f"Incremental Won Pipeline (modeled): **${incremental_won_pipeline:,.0f}**")
        st.caption("(Enhanced Win Rate − Current Win Rate) × Open Pipeline Amount.")

    # -----------------------
    # Executive Summary
    # -----------------------
    st.subheader("Executive Summary")

    bullets = []

    bullets.append(
        f"Won opportunities average **{avg_cr_won:.1f}** contact roles, compared to "
        f"**{avg_cr_lost:.1f}** on lost opportunities, indicating a strong correlation between "
        f"multi-threading and successful outcomes."
    )

    bullets.append(
        f"Open opportunities currently average **{avg_cr_open:.1f}** contact roles. Increasing this "
        f"towards at least **{industry_cr_open:.1f}** contacts per opportunity would better align open deals "
        f"with the contact patterns of won opportunities."
    )

    if avg_days_won is not None and avg_age_open is not None:
        bullets.append(
            f"Won opportunities close in about **{avg_days_won:.0f} days**, while open opportunities have "
            f"already aged around **{avg_age_open:.0f} days**, suggesting older, under-mapped deals are less likely "
            f"to convert without additional stakeholder engagement."
        )

    bullets.append(
        f"If open opportunities improved Contact Role coverage and behaved more like won opportunities, "
        f"the modeled win rate could increase from **{win_rate:.1%}** to **{enhanced_win_rate:.1%}**, unlocking "
        f"an estimated **${incremental_won_pipeline:,.0f}** in incremental won pipeline."
    )

    for b in bullets:
        st.markdown("- " + b)

else:
    st.info("Upload both CSV files above to generate insights.")
