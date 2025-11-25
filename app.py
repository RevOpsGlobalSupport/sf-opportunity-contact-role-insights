import streamlit as st
import pandas as pd
from datetime import datetime
import html

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
          <span style="font-size:16px;font-weight:600;">{label}</span>
          <span class="tooltip-icon">ℹ️
            <span class="tooltip-text">{safe_tip}</span>
          </span>
        </div>
        """,
        unsafe_allow_html=True
    )


# -----------------------
# Streamlit App
# -----------------------
st.set_page_config(page_title="SFDC Opportunity Contact Role Insights", layout="wide")

# Logo + hyperlink header
logo_url = "https://www.revopsglobal.com/wp-content/uploads/2024/09/Footer_Logo.png"
site_url = "https://www.revopsglobal.com/"

st.markdown(
    f"""
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:6px;">
      <a href="{site_url}" target="_blank">
        <img src="{logo_url}" style="height:50px;" />
      </a>
      <h2 style="margin:0;padding:0;">Salesforce Opportunity Contact Role Insights</h2>
    </div>
    """,
    unsafe_allow_html=True
)

# CSS for tooltips (readable)
st.markdown("""
<style>
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

st.markdown(
    f"""
**Upload Opportunities CSV**  
[Download sample]({sample_opps_url})
"""
)
opps_file = st.file_uploader("", type=["csv"], key="opps")

st.markdown(
    f"""
**Upload Opportunities with Contact Roles CSV**  
[Download sample]({sample_roles_url})
"""
)
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

    # cleaning
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
        avg_age_open = (
            open_opps["age_days"].dropna().mean()
            if not open_opps["age_days"].dropna().empty
            else None
        )
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
    # Layout with hover tooltips
    # -----------------------
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Core Metrics")

        label_with_tooltip("Total Opportunities",
                           "Unique Opportunity IDs in the Opportunities export.")
        st.metric("", f"{total_opps:,}")

        label_with_tooltip("Total Pipeline",
                           "Sum of Amount for all opportunities.")
        st.metric("", f"${total_pipeline:,.0f}")

        label_with_tooltip("Current Win Rate",
                           "Closed Won ÷ (Closed Won + Closed Lost), based on Stage containing Won/Lost.")
        st.metric("", f"{win_rate:.1%}")

        st.subheader("Contact Role Coverage")

        label_with_tooltip("Opportunities with Contact Roles",
                           "Unique opportunities that appear in the Contact Roles export.")
        st.write(f"**{opps_with_cr:,}**")

        label_with_tooltip("Opportunities without Contact Roles",
                           "Total Opportunities minus those with Contact Roles.")
        st.write(f"**{opps_without_cr:,}**")

        label_with_tooltip("Pipeline with Contact Roles",
                           "Sum of Amount for opportunities that have ≥1 Contact Role.")
        st.write(f"**${pipeline_with_cr:,.0f}**")

        label_with_tooltip("Pipeline without Contact Roles",
                           "Sum of Amount for opportunities with 0 Contact Roles.")
        st.write(f"**${pipeline_without_cr:,.0f}**")

        label_with_tooltip("Opps with only 1 Contact Role",
                           "Unique opportunities where Contact Role count = 1.")
        st.write(f"**{opps_one_cr:,}**")

        label_with_tooltip("Pipeline with only 1 Contact Role",
                           "Sum of Amount for opportunities with exactly 1 Contact Role.")
        st.write(f"**${pipeline_one_cr:,.0f}**")

    with col2:
        st.subheader("Contact Roles by Outcome")

        label_with_tooltip("Avg Contact Roles – Won",
                           "Average Contact Role count per Opportunity where Stage contains Won.")
        st.write(f"**{avg_cr_won:.1f}**")

        label_with_tooltip("Avg Contact Roles – Lost",
                           "Average Contact Role count per Opportunity where Stage contains Lost.")
        st.write(f"**{avg_cr_lost:.1f}**")

        label_with_tooltip("Avg Contact Roles – Open",
                           "Average Contact Role count per Opportunity where Stage does not contain Won or Lost.")
        st.write(f"**{avg_cr_open:.1f}**")

        st.subheader("Time to Close")

        if avg_days_won is not None:
            label_with_tooltip("Avg days to close – Won",
                               "Average (Close Date − Created Date) for Won opportunities.")
            st.write(f"**{avg_days_won:.0f} days**")

        if avg_days_lost is not None:
            label_with_tooltip("Avg days to close – Lost",
                               "Average (Close Date − Created Date) for Lost opportunities.")
            st.write(f"**{avg_days_lost:.0f} days**")

        if avg_age_open is not None:
            label_with_tooltip("Avg age of Open opps",
                               "Average (Today − Created Date) for open opportunities.")
            st.write(f"**{avg_age_open:.0f} days**")

        st.subheader("Modeled Uplift")

        label_with_tooltip("Contact Influence Ratio (Won vs Lost)",
                           "Avg Contact Roles on Won ÷ Avg Contact Roles on Lost. Indicates how strongly contacts correlate with winning.")
        st.write(f"**{contact_influence_ratio:.2f}×**")

        label_with_tooltip("Enhanced Win Rate (modeled)",
                           "Current Win Rate scaled by Contact Influence Ratio, capped at 95% for sanity.")
        st.write(f"**{enhanced_win_rate:.1%}**")

        label_with_tooltip("Incremental Won Pipeline (modeled)",
                           "(Enhanced Win Rate − Current Win Rate) × Open Pipeline Amount.")
        st.write(f"**${incremental_won_pipeline:,.0f}**")

    # -----------------------
    # Executive Summary (conditional open-coverage bullet)
    # -----------------------
    st.subheader("Executive Summary")

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

    # -----------------------
    # CTA Section after results
    # -----------------------
    st.markdown("---")
    st.markdown(
        """
### Want to drive higher win rates with Buying Groups?

RevOps Global’s **Buying Group Automation** helps sales teams identify stakeholders, close coverage gaps,  
and multi-thread deals faster — directly improving Contact Role coverage and conversion.

👉 **Learn more here:**  
https://www.revopsglobal.com/buying-group-automation/
        """
    )

else:
    st.info("Upload both CSV files above to generate insights.")
