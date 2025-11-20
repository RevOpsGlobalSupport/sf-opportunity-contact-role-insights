import streamlit as st
import pandas as pd
from datetime import datetime

# Page config
st.set_page_config(page_title="SFDC Opportunity Contact Role Insights", layout="wide")

st.title("Salesforce Opportunity Contact Role Insights")

st.markdown(
    """
Upload two Salesforce exports in CSV format:

1. **Opportunities** report  
2. **Opportunities with Contact Roles** report  

Both must follow the template column structure described in your Excel instructions.
"""
)

# File uploads
opps_file = st.file_uploader("Upload Opportunities CSV", type=["csv"], key="opps")
roles_file = st.file_uploader("Upload Opportunities with Contact Roles CSV", type=["csv"], key="roles")

if opps_file and roles_file:
    # Load CSVs
    opps = pd.read_csv(opps_file)
    roles = pd.read_csv(roles_file)

    # Expected columns based on your schema
    opps_cols = [
        "Opportunity ID","Opportunity Name","Account ID","Amount",
        "Type","Stage","Created Date","Closed Date","Opportunity Owner"
    ]
    roles_cols = opps_cols + [
        "Contact ID","Title","Department","Role","Is Primary"
    ]

    missing_opps = [c for c in opps_cols if c not in opps.columns]
    missing_roles = [c for c in roles_cols if c not in roles.columns]

    if missing_opps:
        st.error("Opportunities file is missing columns: " + ", ".join(missing_opps))
    elif missing_roles:
        st.error("Opportunities with Contact Roles file is missing columns: " + ", ".join(missing_roles))
    else:
        # ----- Basic cleaning -----
        opps["Amount"] = pd.to_numeric(opps["Amount"], errors="coerce").fillna(0)
        roles["Amount"] = pd.to_numeric(roles["Amount"], errors="coerce").fillna(0)

        # Parse dates
        def parse_date(x):
            try:
                return pd.to_datetime(x)
            except Exception:
                return pd.NaT

        opps["Created Date"] = opps["Created Date"].apply(parse_date)
        opps["Closed Date"] = opps["Closed Date"].apply(parse_date)

        # Stage masks
        stage = opps["Stage"].fillna("")
        won_mask = stage.str.contains("Won", case=False, na=False)
        lost_mask = stage.str.contains("Lost", case=False, na=False)
        open_mask = ~won_mask & ~lost_mask

        # Contact counts per Opportunity ID
        cr_counts = roles.groupby("Opportunity ID")["Contact ID"].nunique()
        opps = opps.merge(cr_counts.rename("contact_count"), on="Opportunity ID", how="left")
        opps["contact_count"] = opps["contact_count"].fillna(0)

        # ----- Core KPIs -----
        total_opps = opps["Opportunity ID"].nunique()
        total_pipeline = opps["Amount"].sum()

        # Opps with / without Contact Roles
        opps_with_cr_ids = roles["Opportunity ID"].dropna().unique()
        opps_with_cr = opps[opps["Opportunity ID"].isin(opps_with_cr_ids)]["Opportunity ID"].nunique()
        opps_without_cr = total_opps - opps_with_cr

        pipeline_with_cr = opps[opps["Opportunity ID"].isin(opps_with_cr_ids)]["Amount"].sum()
        pipeline_without_cr = total_pipeline - pipeline_with_cr

        # Opps with only 1 Contact Role
        one_cr_ids = cr_counts[cr_counts == 1].index
        opps_one_cr = opps[opps["Opportunity ID"].isin(one_cr_ids)]["Opportunity ID"].nunique()
        pipeline_one_cr = opps[opps["Opportunity ID"].isin(one_cr_ids)]["Amount"].sum()

        # Segment by status
        open_opps = opps[open_mask]
        won_opps = opps[won_mask]
        lost_opps = opps[lost_mask]

        # Win rate: Won / (Won + Lost)
        won_count = won_opps["Opportunity ID"].nunique()
        lost_count = lost_opps["Opportunity ID"].nunique()
        win_rate = won_count / (won_count + lost_count) if (won_count + lost_count) > 0 else 0

        # ----- Contact Role averages -----
        avg_cr_lost = lost_opps["contact_count"].mean() if not lost_opps.empty else 0
        avg_cr_won = won_opps["contact_count"].mean() if not won_opps.empty else 0
        avg_cr_open = open_opps["contact_count"].mean() if not open_opps.empty else 0

        # ----- Time metrics -----
        # Days from Created to Closed for Won and Lost
        def days_diff(row):
            if pd.isna(row["Created Date"]) or pd.isna(row["Closed Date"]):
                return None
            return (row["Closed Date"] - row["Created Date"]).days

        if not lost_opps.empty:
            lost_opps = lost_opps.copy()
            lost_opps["days_to_close"] = lost_opps.apply(days_diff, axis=1)
        if not won_opps.empty:
            won_opps = won_opps.copy()
            won_opps["days_to_close"] = won_opps.apply(days_diff, axis=1)

        avg_days_lost = lost_opps["days_to_close"].dropna().mean() if "days_to_close" in lost_opps else None
        avg_days_won = won_opps["days_to_close"].dropna().mean() if "days_to_close" in won_opps else None

        # Age of open opps = TODAY - Created Date
        today = pd.Timestamp.today().normalize()
        if not open_opps.empty:
            open_opps = open_opps.copy()
            open_opps["age_days"] = (today - open_opps["Created Date"]).dt.days
        avg_age_open = open_opps["age_days"].dropna().mean() if "age_days" in open_opps else None

        # ----- Simple uplift model (can be tuned) -----
        # Industry benchmarks (you can later expose as inputs)
        industry_cr_open = 2.0  # example target for open opps

        # Contact Influence Ratio: Won vs Lost
        if avg_cr_lost not in (0, None):
            contact_influence_ratio = avg_cr_won / avg_cr_lost if avg_cr_won is not None else 1
        else:
            contact_influence_ratio = 1

        # Enhanced win rate (very simple heuristic - can refine)
        enhanced_win_rate = win_rate * contact_influence_ratio
        enhanced_win_rate = min(enhanced_win_rate, 0.95)  # cap for sanity

        open_pipeline = open_opps["Amount"].sum()
        incremental_won_pipeline = max(0, (enhanced_win_rate - win_rate) * open_pipeline)

        # ----- Layout -----
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Core Metrics")
            st.metric("Total Opportunities", f"{total_opps:,}")
            st.metric("Total Pipeline", f"${total_pipeline:,.0f}")
            st.metric("Current Win Rate", f"{win_rate:.1%}")

            st.subheader("Contact Role Coverage")
            st.write(f"Opportunities **with** Contact Roles: **{opps_with_cr:,}**")
            st.write(f"Opportunities **without** Contact Roles: **{opps_without_cr:,}**")
            st.write(f"Pipeline with Contact Roles: **${pipeline_with_cr:,.0f}**")
            st.write(f"Pipeline without Contact Roles: **${pipeline_without_cr:,.0f}**")
            st.write(f"Opps with **only 1** Contact Role: **{opps_one_cr:,}**")
            st.write(f"Pipeline with only 1 Contact Role: **${pipeline_one_cr:,.0f}**")

        with col2:
            st.subheader("Contact Roles by Outcome")
            st.write(f"Avg Contact Roles – **Won**: **{avg_cr_won:.1f}**")
            st.write(f"Avg Contact Roles – **Lost**: **{avg_cr_lost:.1f}**")
            st.write(f"Avg Contact Roles – **Open**: **{avg_cr_open:.1f}**")

            st.subheader("Time to Close")
            if avg_days_won is not None:
                st.write(f"Avg days to close – **Won**: **{avg_days_won:.0f}**")
            if avg_days_lost is not None:
                st.write(f"Avg days to close – **Lost**: **{avg_days_lost:.0f}**")
            if avg_age_open is not None:
                st.write(f"Avg age of **Open** opps: **{avg_age_open:.0f}**")

            st.subheader("Modeled Uplift")
            st.write(f"Contact Influence Ratio (Won vs Lost): **{contact_influence_ratio:.2f}×**")
            st.write(f"Enhanced Win Rate (modeled): **{enhanced_win_rate:.1%}**")
            st.write(f"Incremental Won Pipeline (modeled): **${incremental_won_pipeline:,.0f}**")

        # ----- Executive Summary -----
        st.subheader("Executive Summary")

        bullets = []

        bullets.append(
            f"Won opportunities average **{avg_cr_won:.1f}** contact roles, "
            f"compared to **{avg_cr_lost:.1f}** on lost opportunities, indicating that "
            f"multi-threading is strongly correlated with successful outcomes."
        )

        bullets.append(
            f"Open opportunities currently average **{avg_cr_open:.1f}** contact roles, "
            f"which is closer to lost than won behavior. Increasing this towards at least "
            f"**{industry_cr_open:.1f}** contacts per opportunity would better align open deals with "
            f"the contact patterns of won opportunities."
        )

        if avg_days_won is not None and avg_age_open is not None:
            bullets.append(
                f"Won opportunities close in about **{avg_days_won:.0f} days**, while open opportunities "
                f"have already aged around **{avg_age_open:.0f} days**, suggesting that older, under-mapped deals "
                f"are less likely to convert without additional stakeholder engagement."
            )

        bullets.append(
            f"If open opportunities improved Contact Role coverage and behaved more like current won opportunities, "
            f"the modeled win rate could increase from **{win_rate:.1%}** to **{enhanced_win_rate:.1%}**, "
            f"unlocking an estimated **${incremental_won_pipeline:,.0f}** in incremental won pipeline."
        )

        for b in bullets:
            st.markdown("- " + b)

else:
    st.info("Upload both CSV files above to generate insights.")
