import streamlit as st
import pandas as pd
from datetime import datetime
import html
import altair as alt

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
# Tooltip helper (CSS hover)
# -----------------------
def label_with_tooltip(label: str, tooltip: str):
    safe_tip = html.escape(tooltip)
    st.markdown(
        f"""
        <div class="tooltip-wrap">
          <span class="kpi-label">{label}</span>
          <span class="tooltip-icon">ℹ️
            <span class="tooltip-text">{safe_tip}</span>
          </span>
        </div>
        """,
        unsafe_allow_html=True
    )

def show_value(value: str):
    st.markdown(f"<div class='kpi-value'>{value}</div>", unsafe_allow_html=True)

# -----------------------
# Section "card" wrappers
# -----------------------
def section_start(title: str):
    st.markdown(f"""
    <div class="section-card">
      <div class="section-title">{html.escape(title)}</div>
      <div class="section-divider"></div>
    """, unsafe_allow_html=True)

def section_end():
    st.markdown("</div>", unsafe_allow_html=True)


# -----------------------
# Streamlit App
# -----------------------
st.set_page_config(page_title="SFDC Opportunity Contact Role Insights", layout="wide")

# --- Logo at top (bigger) ---
logo_url = "https://www.revopsglobal.com/wp-content/uploads/2024/09/Footer_Logo.png"
site_url = "https://www.revopsglobal.com/"

st.markdown(
    f"""
    <div style="margin-top:4px;">
      <a href="{site_url}" target="_blank">
        <img src="{logo_url}" style="height:90px;" />
      </a>
    </div>
    <div style="height:18px;"></div>
    """,
    unsafe_allow_html=True
)

# Title starts below logo
st.markdown(
    """
    <div style="font-size:28px;font-weight:700;line-height:1.15;">
      Salesforce Opportunity Contact Role Insights
    </div>
    <div style="font-size:15px;color:#6b7280;margin-top:6px;margin-bottom:10px;">
      Measure Contact Role coverage and its impact on win rates
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("<hr style='margin: 8px 0 16px 0; border:0; border-top:1px solid #e5e7eb;' />", unsafe_allow_html=True)

# -----------------------
# CSS (tooltips + KPI sizing + section cards)
# -----------------------
st.markdown("""
<style>
/* ====== BRAND-Y CARD COLORS (tweak these 2 values if needed) ====== */
:root{
  --card-bg: #F5F9FF;     /* mild light-blue background */
  --card-border: #E2ECFA; /* soft blue border */
}

/* KPI text sizing */
.kpi-label {
  font-size:16px;
  font-weight:600;
}
.kpi-value {
  font-size:20px;
  font-weight:700;
  margin:4px 0 14px 0;
}

/* Section cards */
.section-card{
  background: var(--card-bg);
  border: 1px solid var(--card-border);
  border-radius: 12px;
  padding: 14px 16px 6px 16px;
  margin: 8px 0 14px 0;
}
.section-title{
  font-size:20px;
  font-weight:700;
  margin-bottom:6px;
}
.section-divider{
  height:1px;
  background: var(--card-border);
  margin: 6px 0 12px 0;
  border-radius:2px;
}

/* Tooltip styles */
.tooltip-wrap {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  margin-top: 4px;
}
.tooltip-icon {
  position: relative;
  display: inline-block;
  cursor: help;
  font-size: 14px;
  opacity: 0.9;
}
.tooltip-icon .tooltip-text {
  visibility: hidden;
  width: 280px;
  background-color: #111827;
  color: #FFFFFF;
  text-align: left;
  border-radius: 6px;
  padding: 10px 12px;
  position: absolute;
  z-index: 9999;
  bottom: 135%;
  left: 50%;
  transform: translateX(-50%);
  font-size: 14px;
  line-height: 1.45;
  box-shadow: 0 4px 12px rgba(0,0,0,0.35);
  white-space: normal;
  opacity: 1;
}
.tooltip-icon .tooltip-text::after {
  content: "";
  position: absolute;
  top: 100%;
  left: 50%;
  margin-left: -6px;
  border-width: 6px;
  border-style: solid;
  border-color: #111827 transparent transparent transparent;
}
.tooltip-icon:hover .tooltip-text {
  visibility: visible;
}
</style>
""", unsafe_allow_html=True)


# -----------------------
# Instructions section (expandable)
# -----------------------
with st.expander("📌 How to export Salesforce data & use this app", expanded=False):
    st.markdown(
        """
**Step 1 — Export Opportunity Data from Salesforce**

Create an Opportunities report in Salesforce. Use **Tabular** format and include the following fields, in this exact order:  
Opportunity ID, Opportunity Name, Account ID, Amount, Type, Stage, Created Date, Closed Date, Opportunity Owner.

- Go to **Reports → New Report → Select Opportunities**
- Switch to **Tabular**
- Add the fields in the exact order listed above
- Set filters to:
  - **Show Me:** All Opportunities  
  - **Date Range:** All Time (or the period you want to analyze)
- Export the report as a **CSV file**

You will upload this into the app as the **Opportunities CSV**.

---

**Step 2 — Export Opportunity Contact Role Data**

Create an **Opportunities with Contact Roles** report. Again use **Tabular** format and include these fields, exactly in this order:  
Opportunity ID, Opportunity Name, Account ID, Amount, Type, Stage, Opportunity Created Date, Opportunity Closed Date, Opportunity Owner, Contact ID, Title, Department, Role, Is Primary.

- Go to **Reports → New Report → Opportunities with Contact Roles**
- Switch to **Tabular**
- Add fields in the exact order above
- Set filters the same way:
  - **Show Me:** All Opportunities  
  - **Date Range:** All Time
- Export as a **CSV file**

You will upload this into the app as the **Opportunities with Contact Roles CSV**.

---

**Step 3 — Upload Data into This App**

After exporting both reports:

- Upload each CSV into the correct upload box in this app.

Do **not**:
- Change column order
- Rename columns
- Add/delete columns
- Insert blank rows or columns
- Paste additional headers
        """
    )

st.markdown(
    """
Upload two Salesforce exports in CSV format:

1. **Opportunities** report  
2. **Opportunities with Contact Roles** report  

The app normalizes common Salesforce header variations  
(e.g., **Close Date vs Closed Date**, **Opportunity Created Date vs Created Date**).
"""
)

# Sample CSV links
sample_opps_url = "https://drive.google.com/file/d/11bNN1lSs6HtPyXVV0k9yYboO6XXdI85H/view?usp=sharing"
sample_roles_url = "https://drive.google.com/file/d/1-w_yFF0_naXGUEX00TKOMDT2bAW7ehPE/view?usp=sharing"

st.markdown(f"**Upload Opportunities CSV**  \n[Download sample]({sample_opps_url})")
opps_file = st.file_uploader("", type=["csv"], key="opps")

st.markdown(f"**Upload Opportunities with Contact Roles CSV**  \n[Download sample]({sample_roles_url})")
roles_file = st.file_uploader("", type=["csv"], key="roles")

if opps_file and roles_file:
    try:
        raw_opps = load_csv(opps_file)
        raw_roles = load_csv(roles_file)
    except Exception as e:
        st.error(f"Error reading CSV files. Please confirm they are valid CSV exports. Details: {e}")
        st.stop()

    opps = normalize_and_standardize_columns(raw_opps, is_roles=False)
    roles = normalize_and_standardize_columns(raw_roles, is_roles=True)

    required_opps = [
        "Opportunity ID", "Opportunity Name", "Account ID", "Amount",
        "Type", "Stage", "Created Date", "Close Date", "Opportunity Owner"
    ]
    required_roles = ["Opportunity ID", "Contact ID"]

    missing_opps = [c for c in required_opps if c not in opps.columns]
    missing_roles = [c for c in required_roles if c not in roles.columns]

    if missing_opps:
        st.error("Opportunities file missing columns: " + ", ".join(missing_opps))
        st.stop()
    if missing_roles:
        st.error("Contact Roles file missing columns: " + ", ".join(missing_roles))
        st.stop()

    # cleaning
    opps["Amount"] = pd.to_numeric(opps["Amount"], errors="coerce").fillna(0)
    roles["Amount"] = pd.to_numeric(roles.get("Amount", 0), errors="coerce").fillna(0)

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

    # core KPIs
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

    # contact role averages
    avg_cr_lost = lost_opps["contact_count"].mean() if not lost_opps.empty else 0
    avg_cr_won = won_opps["contact_count"].mean() if not won_opps.empty else 0
    avg_cr_open = open_opps["contact_count"].mean() if not open_opps.empty else 0

    # time metrics
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
        avg_age_open = open_opps["age_days"].dropna().mean()
    else:
        avg_age_open = None

    # uplift model
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
    # SINGLE COLUMN METRICS (CARD SECTIONS)
    # -----------------------
    section_start("Core Metrics")
    label_with_tooltip("Total Opportunities", "Unique Opportunity IDs in the Opportunities export.")
    show_value(f"{total_opps:,}")
    label_with_tooltip("Total Pipeline", "Sum of Amount for all opportunities.")
    show_value(f"${total_pipeline:,.0f}")
    label_with_tooltip("Current Win Rate", "Closed Won ÷ (Closed Won + Closed Lost).")
    show_value(f"{win_rate:.1%}")
    section_end()

    section_start("Contact Role Coverage")
    label_with_tooltip("Opportunities with Contact Roles", "Unique opportunities that appear in the Contact Roles export.")
    show_value(f"{opps_with_cr:,}")
    label_with_tooltip("Opportunities without Contact Roles", "Total Opportunities minus those with Contact Roles.")
    show_value(f"{opps_without_cr:,}")
    label_with_tooltip("Pipeline with Contact Roles", "Sum of Amount for opportunities with ≥1 Contact Role.")
    show_value(f"${pipeline_with_cr:,.0f}")
    label_with_tooltip("Pipeline without Contact Roles", "Sum of Amount for opportunities with 0 Contact Roles.")
    show_value(f"${pipeline_without_cr:,.0f}")
    label_with_tooltip("Opps with only 1 Contact Role", "Unique opportunities where Contact Role count = 1.")
    show_value(f"{opps_one_cr:,}")
    label_with_tooltip("Pipeline with only 1 Contact Role", "Sum of Amount for opportunities with exactly 1 Contact Role.")
    show_value(f"${pipeline_one_cr:,.0f}")
    section_end()

    section_start("Contact Roles by Outcome")
    label_with_tooltip("Avg Contact Roles – Won", "Average Contact Role count per Opportunity where Stage contains Won.")
    show_value(f"{avg_cr_won:.1f}")
    label_with_tooltip("Avg Contact Roles – Lost", "Average Contact Role count per Opportunity where Stage contains Lost.")
    show_value(f"{avg_cr_lost:.1f}")
    label_with_tooltip("Avg Contact Roles – Open", "Average Contact Role count per Opportunity where Stage does not contain Won or Lost.")
    show_value(f"{avg_cr_open:.1f}")
    section_end()

    section_start("Time to Close")
    label_with_tooltip("Avg days to close – Won", "Average (Close Date − Created Date) for Won opportunities.")
    show_value(f"{avg_days_won:.0f} days" if avg_days_won is not None else "0 days")
    label_with_tooltip("Avg days to close – Lost", "Average (Close Date − Created Date) for Lost opportunities.")
    show_value(f"{avg_days_lost:.0f} days" if avg_days_lost is not None else "0 days")
    label_with_tooltip("Avg age of Open opps", "Average (Today − Created Date) for open opportunities.")
    show_value(f"{avg_age_open:.0f} days" if avg_age_open is not None else "0 days")
    section_end()

    section_start("Modeled Uplift")
    label_with_tooltip("Contact Influence Ratio (Won vs Lost)", "Avg Contact Roles on Won ÷ Avg Contact Roles on Lost.")
    show_value(f"{contact_influence_ratio:.2f}×")
    label_with_tooltip("Enhanced Win Rate (modeled)", "Current Win Rate scaled by Contact Influence Ratio, capped at 95%.")
    show_value(f"{enhanced_win_rate:.1%}")
    label_with_tooltip("Incremental Won Pipeline (modeled)", "(Enhanced Win Rate − Current Win Rate) × Open Pipeline Amount.")
    show_value(f"${incremental_won_pipeline:,.0f}")
    section_end()

    # -----------------------
    # Executive Summary
    # -----------------------
    section_start("Executive Summary")
    bullets = []

    bullets.append(
        f"Won opportunities average **{avg_cr_won:.1f}** contact roles, compared to "
        f"**{avg_cr_lost:.1f}** on lost opportunities, indicating a strong correlation between "
        f"multi-threading and successful outcomes."
    )

    if (avg_cr_open < avg_cr_won) or (avg_cr_open < industry_cr_open):
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
    section_end()

    # -----------------------
    # Insights (4 clean charts)
    # -----------------------
    section_start("Insights")

    chart_df = opps.copy()
    chart_df["Stage Group"] = "Open"
    chart_df.loc[won_mask, "Stage Group"] = "Won"
    chart_df.loc[lost_mask, "Stage Group"] = "Lost"

    chart_df["contact_count"] = pd.to_numeric(chart_df["contact_count"], errors="coerce").fillna(0)
    chart_df["Amount"] = pd.to_numeric(chart_df["Amount"], errors="coerce").fillna(0)

    def contact_bucket(n):
        try:
            n = float(n)
        except Exception:
            n = 0
        if n <= 0:
            return "0"
        if n == 1:
            return "1"
        if n == 2:
            return "2"
        if n == 3:
            return "3"
        return "4+"

    chart_df["Contact Bucket"] = chart_df["contact_count"].apply(contact_bucket)
    bucket_order = ["0", "1", "2", "3", "4+"]

    # 1) Win Rate vs Contact Roles (binned)
    closed_df = chart_df[chart_df["Stage Group"].isin(["Won", "Lost"])].copy()
    winrate_bucket = (
        closed_df.groupby("Contact Bucket")
        .agg(
            won=("Stage Group", lambda s: (s == "Won").sum()),
            lost=("Stage Group", lambda s: (s == "Lost").sum())
        )
        .reset_index()
    )
    winrate_bucket["Win Rate"] = winrate_bucket.apply(
        lambda r: r["won"] / (r["won"] + r["lost"]) if (r["won"] + r["lost"]) > 0 else 0,
        axis=1
    )

    st.caption("Win rate increases as more stakeholders are engaged.")
    chart_winrate = (
        alt.Chart(winrate_bucket)
        .mark_line(point=True)
        .encode(
            x=alt.X("Contact Bucket:N", sort=bucket_order, title="Contact Roles per Opportunity (bucketed)"),
            y=alt.Y("Win Rate:Q", axis=alt.Axis(format="%"), title="Win Rate"),
            tooltip=[
                alt.Tooltip("Contact Bucket:N", title="Bucket"),
                alt.Tooltip("Win Rate:Q", format=".1%", title="Win Rate"),
                alt.Tooltip("won:Q", title="Won Opps"),
                alt.Tooltip("lost:Q", title="Lost Opps")
            ]
        )
        .properties(height=260)
    )
    st.altair_chart(chart_winrate, use_container_width=True)

    # 2) Open Pipeline at Risk
    open_df = chart_df[chart_df["Stage Group"] == "Open"].copy()
    open_df["Open Coverage Bucket"] = open_df["contact_count"].apply(
        lambda n: "0 Contact Roles" if n == 0 else ("1 Contact Role" if n == 1 else "2+ Contact Roles")
    )

    open_pipeline_bucket = (
        open_df.groupby("Open Coverage Bucket")["Amount"]
        .sum()
        .reindex(["0 Contact Roles", "1 Contact Role", "2+ Contact Roles"])
        .fillna(0)
        .reset_index()
        .rename(columns={"Amount": "Open Pipeline"})
    )

    st.caption("How much open pipeline is under-covered today.")
    chart_open_pipeline = (
        alt.Chart(open_pipeline_bucket)
        .mark_bar()
        .encode(
            x=alt.X("Open Coverage Bucket:N", title="Open Opportunity Coverage"),
            y=alt.Y("Open Pipeline:Q", title="Open Pipeline ($)"),
            tooltip=[
                "Open Coverage Bucket",
                alt.Tooltip("Open Pipeline:Q", format=",.0f")
            ]
        )
        .properties(height=260)
    )
    st.altair_chart(chart_open_pipeline, use_container_width=True)

    # 3) Avg Days/Age vs Contact Roles
    time_df = chart_df.copy()

    def safe_days_to_close(row):
        if pd.isna(row["Created Date"]) or pd.isna(row["Close Date"]):
            return None
        return (row["Close Date"] - row["Created Date"]).days

    time_df["days_to_close"] = time_df.apply(safe_days_to_close, axis=1)
    today = pd.Timestamp.today().normalize()
    time_df["open_age_days"] = None
    open_mask_local = (time_df["Stage Group"] == "Open") & time_df["Created Date"].notna()
    time_df.loc[open_mask_local, "open_age_days"] = (today - time_df.loc[open_mask_local, "Created Date"]).dt.days

    agg_rows = []
    for sg in ["Won", "Lost"]:
        tmp = time_df[time_df["Stage Group"] == sg].copy()
        grp = (
            tmp.groupby("Contact Bucket")["days_to_close"]
            .mean()
            .reindex(bucket_order)
            .reset_index()
        )
        grp["Stage Group"] = sg
        grp = grp.rename(columns={"days_to_close": "Avg Days"})
        agg_rows.append(grp)

    tmp_open = time_df[time_df["Stage Group"] == "Open"].copy()
    grp_open = (
        tmp_open.groupby("Contact Bucket")["open_age_days"]
        .mean()
        .reindex(bucket_order)
        .reset_index()
    )
    grp_open["Stage Group"] = "Open"
    grp_open = grp_open.rename(columns={"open_age_days": "Avg Days"})
    agg_rows.append(grp_open)

    avg_days_bucket = pd.concat(agg_rows, ignore_index=True)

    st.caption("More contact roles correlate with faster closes and younger open pipeline.")
    chart_days = (
        alt.Chart(avg_days_bucket)
        .mark_bar()
        .encode(
            x=alt.X("Contact Bucket:N", sort=bucket_order, title="Contact Roles per Opportunity (bucketed)"),
            y=alt.Y("Avg Days:Q", title="Avg Days (Close for Won/Lost, Age for Open)"),
            color=alt.Color("Stage Group:N", title="Outcome"),
            tooltip=[
                "Stage Group",
                "Contact Bucket",
                alt.Tooltip("Avg Days:Q", format=",.0f")
            ]
        )
        .properties(height=300)
    )
    st.altair_chart(chart_days, use_container_width=True)

    # 4) Before vs After Win Rate
    impact_df = pd.DataFrame({
        "Scenario": ["Current Win Rate", "Enhanced Win Rate"],
        "Win Rate": [win_rate, enhanced_win_rate]
    })

    st.caption("Modeled uplift if open opportunities matched won-deal contact patterns.")
    chart_impact = (
        alt.Chart(impact_df)
        .mark_bar()
        .encode(
            x=alt.X("Scenario:N", title=""),
            y=alt.Y("Win Rate:Q", axis=alt.Axis(format="%"), title="Win Rate"),
            tooltip=["Scenario", alt.Tooltip("Win Rate:Q", format=".1%")]
        )
        .properties(height=240)
    )
    st.altair_chart(chart_impact, use_container_width=True)

    st.markdown(
        f"**Modeled incremental won pipeline from improved coverage:** "
        f"${incremental_won_pipeline:,.0f}"
    )

    section_end()

    # -----------------------
    # CTA Section after results
    # -----------------------
    section_start("Buying Group Automation")
    st.markdown(
        """
RevOps Global’s **Buying Group Automation** helps sales teams identify stakeholders, close coverage gaps,  
and multi-thread deals faster — directly improving Contact Role coverage and conversion.

👉 **Learn more here:**  
https://www.revopsglobal.com/buying-group-automation/
        """
    )
    section_end()

else:
    st.info("Upload both CSV files above to generate insights.")


# -----------------------
# Footer / Copyright
# -----------------------
st.markdown(
    f"""
<hr style="margin-top:26px; border:0; border-top:1px solid #e5e7eb;" />
<div style="font-size:12px; color:#6b7280; text-align:center; padding:10px 0 2px 0;">
  © {datetime.now().year} RevOps Global. All rights reserved.
</div>
    """,
    unsafe_allow_html=True
)
