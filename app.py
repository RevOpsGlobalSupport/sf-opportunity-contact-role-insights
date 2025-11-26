import streamlit as st
import pandas as pd
from datetime import datetime
import html
import altair as alt

# PDF deps
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_LEFT

import io
import requests
from PIL import Image as PILImage

# Matplotlib for PDF charts
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter


# -----------------------
# Robust CSV loading
# -----------------------
def load_csv(file):
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
# Tooltip helper
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
# Section wrappers
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
# PDF helpers
# -----------------------
LOGO_URL = "https://www.revopsglobal.com/wp-content/uploads/2024/09/Footer_Logo.png"
SITE_URL = "https://www.revopsglobal.com/"
CTA_URL = "https://www.revopsglobal.com/buying-group-automation/"

def fetch_logo_bytes(url: str):
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return io.BytesIO(r.content)
    except Exception:
        return None

def pdf_watermark_and_footer(c: canvas.Canvas, doc):
    c.saveState()
    c.setFont("Helvetica-Bold", 50)
    c.setFillColor(colors.HexColor("#E6EAF0"))
    c.translate(300, 400)
    c.rotate(30)
    c.drawCentredString(0, 0, "RevOps Global")
    c.restoreState()

    c.saveState()
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.grey)
    footer_text = f"© {datetime.now().year} RevOps Global. All rights reserved."
    c.drawCentredString(letter[0] / 2, 0.5 * inch, footer_text)
    c.restoreState()

def fig_to_png_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf

def build_pdf_report(
    metrics_dict,
    bullets,
    owner_rollup_rows,
    top_opps_rows,
    chart_pngs,
    segment_rows,
    won_zero_rows
):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=0.75*inch,
        rightMargin=0.75*inch,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="H1", fontSize=18, leading=22, spaceAfter=10, alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="H2", fontSize=13, leading=16, spaceBefore=8, spaceAfter=6,
                              textColor=colors.HexColor("#0F172A")))
    styles.add(ParagraphStyle(name="Body", fontSize=10.5, leading=14))
    styles.add(ParagraphStyle(name="Small", fontSize=9.5, leading=12, textColor=colors.grey))

    story = []

    logo_bytes = fetch_logo_bytes(LOGO_URL)
    if logo_bytes:
        try:
            pil_img = PILImage.open(logo_bytes)
            w, h = pil_img.size
            aspect = h / w
            img_width = 2.2 * inch
            img_height = img_width * aspect
            logo_bytes.seek(0)
            story.append(Image(logo_bytes, width=img_width, height=img_height))
            story.append(Spacer(1, 0.15*inch))
        except Exception:
            pass

    story.append(Paragraph("Salesforce Opportunity Contact Role Insights", styles["H1"]))
    story.append(Paragraph("Report generated by RevOps Global", styles["Small"]))
    story.append(Spacer(1, 0.2*inch))

    for section_title, rows in metrics_dict.items():
        story.append(Paragraph(section_title, styles["H2"]))
        table_data = [["Metric", "Value"]] + rows
        t = Table(table_data, colWidths=[3.7*inch, 2.7*inch])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F1F5F9")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#0F172A")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 10.5),
            ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
            ("ALIGN", (1,1), (1,-1), "RIGHT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("FONTSIZE", (0,1), (-1,-1), 10),
            ("BACKGROUND", (0,1), (-1,-1), colors.white),
        ]))
        story.append(t)
        story.append(Spacer(1, 0.12*inch))

    story.append(Paragraph("Executive Summary", styles["H2"]))
    for b in bullets:
        story.append(Paragraph(f"• {html.escape(b)}", styles["Body"]))
    story.append(Spacer(1, 0.12*inch))

    if won_zero_rows:
        story.append(Paragraph("Red Flag — Won Deals Missing Contact Roles", styles["H2"]))
        for r in won_zero_rows:
            story.append(Paragraph(f"• {html.escape(r)}", styles["Body"]))
        story.append(Spacer(1, 0.12*inch))

    story.append(Paragraph("Insights", styles["H2"]))
    story.append(Paragraph(
        "The following charts summarize the relationship between stakeholder coverage, velocity, and win rate.",
        styles["Body"]
    ))
    story.append(Spacer(1, 0.1*inch))

    for i, png_buf in enumerate(chart_pngs, start=1):
        story.append(Image(png_buf, width=6.7*inch, height=3.2*inch))
        story.append(Spacer(1, 0.15*inch))
        if i in (2, 4):
            story.append(PageBreak())

    story.append(Paragraph("Segment Uplift (Deal Size Bands)", styles["H2"]))
    story.append(Paragraph(
        "Bands are derived automatically from your opportunity Amount distribution (33rd / 67th percentiles).",
        styles["Body"]
    ))
    if segment_rows:
        seg_table = [["Segment", "Open Pipeline", "Avg Contacts (Open)", "Incremental Won Pipeline (modeled)"]] + segment_rows
        t2 = Table(seg_table, colWidths=[1.4*inch, 1.8*inch, 1.6*inch, 1.9*inch])
        t2.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F1F5F9")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
            ("ALIGN", (1,1), (-1,-1), "RIGHT"),
            ("FONTSIZE", (0,0), (-1,-1), 9.8),
        ]))
        story.append(t2)
    else:
        story.append(Paragraph("Not enough open opportunities to segment.", styles["Body"]))
    story.append(Spacer(1, 0.12*inch))

    story.append(Paragraph("Owner Coverage Rollup (Coaching View)", styles["H2"]))
    story.append(Paragraph(
        "Coach the owners with the highest under-covered open pipeline first.",
        styles["Body"]
    ))
    if owner_rollup_rows:
        for r in owner_rollup_rows:
            story.append(Paragraph(f"• {html.escape(r)}", styles["Body"]))
    else:
        story.append(Paragraph("No owner rollup available.", styles["Body"]))
    story.append(Spacer(1, 0.12*inch))

    story.append(Paragraph("Top Open Opportunities to Fix First", styles["H2"]))
    story.append(Paragraph(
        "Prioritize multi-threading these deals first to reduce risk and improve conversion.",
        styles["Body"]
    ))
    if top_opps_rows:
        for r in top_opps_rows:
            story.append(Paragraph(f"• {html.escape(r)}", styles["Body"]))
    else:
        story.append(Paragraph("No open opportunities found.", styles["Body"]))
    story.append(Spacer(1, 0.12*inch))

    story.append(Paragraph("Buying Group Automation", styles["H2"]))
    story.append(Paragraph(
        "RevOps Global’s Buying Group Automation helps sales teams identify stakeholders, close coverage gaps, and multi-thread deals faster.",
        styles["Body"]
    ))
    story.append(Paragraph(f"Learn more: {CTA_URL}", styles["Body"]))

    doc.build(story, onFirstPage=pdf_watermark_and_footer, onLaterPages=pdf_watermark_and_footer)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------
# App Setup
# -----------------------
st.set_page_config(page_title="SFDC Opportunity Contact Role Insights", layout="wide")

st.markdown(
    f"""
    <div style="margin-top:4px;">
      <a href="{SITE_URL}" target="_blank">
        <img src="{LOGO_URL}" style="height:90px;" />
      </a>
    </div>
    <div style="height:18px;"></div>
    """,
    unsafe_allow_html=True
)

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

st.markdown("""
<style>
:root{
  --card-bg: #F5F9FF;
  --card-border: #E2ECFA;
}
.kpi-label { font-size:16px; font-weight:600; }
.kpi-value { font-size:20px; font-weight:700; margin:4px 0 14px 0; }

.section-card{
  background: var(--card-bg);
  border: 1px solid var(--card-border);
  border-radius: 12px;
  padding: 14px 16px 6px 16px;
  margin: 8px 0 14px 0;
}
.section-title{ font-size:20px; font-weight:700; margin-bottom:6px; }
.section-divider{
  height:1px; background: var(--card-border);
  margin: 6px 0 12px 0; border-radius:2px;
}

.tooltip-wrap { display:inline-flex; align-items:center; gap:6px; margin-top:4px; }
.tooltip-icon { position:relative; cursor:help; font-size:14px; opacity:0.9; }
.tooltip-icon .tooltip-text{
  visibility:hidden; width:320px; background:#111827; color:#fff;
  border-radius:6px; padding:10px 12px; position:absolute; z-index:9999;
  bottom:135%; left:50%; transform:translateX(-50%);
  font-size:14px; line-height:1.45; box-shadow:0 4px 12px rgba(0,0,0,0.35);
  opacity:1;
}
.tooltip-icon .tooltip-text::after{
  content:""; position:absolute; top:100%; left:50%; margin-left:-6px;
  border-width:6px; border-style:solid; border-color:#111827 transparent transparent transparent;
}
.tooltip-icon:hover .tooltip-text{ visibility:visible; }
</style>
""", unsafe_allow_html=True)

# Instructions
with st.expander("📌 How to export Salesforce data & use this app", expanded=False):
    st.markdown(
        """
**Step 1 — Export Opportunity Data from Salesforce**

Create an Opportunities report in Salesforce. Use **Tabular** format and include the following fields, in this exact order:  
Opportunity ID, Opportunity Name, Account ID, Amount, Type, Stage, Created Date, Closed Date, Opportunity Owner.

- Go to **Reports → New Report → Select Opportunities**
- Switch to **Tabular**
- Add fields in the exact order listed above
- Filters:
  - **Show Me:** All Opportunities
  - **Date Range:** All Time (or your analysis window)
- Export as **CSV**

---

**Step 2 — Export Opportunity Contact Role Data**

Create an **Opportunities with Contact Roles** report in Tabular format with fields in this order:  
Opportunity ID, Opportunity Name, Account ID, Amount, Type, Stage, Opportunity Created Date, Opportunity Closed Date, Opportunity Owner, Contact ID, Title, Department, Role, Is Primary.

- Reports → New Report → Opportunities with Contact Roles
- Tabular → add fields exactly in order
- Filters same as above
- Export as CSV

---

**Step 3 — Upload Data**

Upload both CSVs here. Do NOT:
- change column order
- rename columns
- add/remove columns
- insert blank rows/headers
        """
    )

st.markdown("Upload two Salesforce exports: **Opportunities** and **Opportunities with Contact Roles**.")

sample_opps_url = "https://drive.google.com/file/d/11bNN1lSs6HtPyXVV0k9yYboO6XXdI85H/view?usp=sharing"
sample_roles_url = "https://drive.google.com/file/d/1-w_yFF0_naXGUEX00TKOMDT2bAW7ehPE/view?usp=sharing"

st.markdown(f"**Upload Opportunities CSV**  \n[Download sample]({sample_opps_url})")
opps_file = st.file_uploader("", type=["csv"], key="opps")

st.markdown(f"**Upload Opportunities with Contact Roles CSV**  \n[Download sample]({sample_roles_url})")
roles_file = st.file_uploader("", type=["csv"], key="roles")

if opps_file and roles_file:
    raw_opps = load_csv(opps_file)
    raw_roles = load_csv(roles_file)

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
        st.error("Opportunities file missing columns: " + ", ".join(missing_opps)); st.stop()
    if missing_roles:
        st.error("Contact Roles file missing columns: " + ", ".join(missing_roles)); st.stop()

    opps["Amount"] = pd.to_numeric(opps["Amount"], errors="coerce").fillna(0)
    opps["Created Date"] = opps["Created Date"].apply(parse_date)
    opps["Close Date"] = opps["Close Date"].apply(parse_date)

    stage = opps["Stage"].fillna("").astype(str)

    # Stage Mapping
    section_start("Stage Mapping (Customer-specific)")
    st.caption("Map your Salesforce stages into Early, Mid, Late, Won, and Lost buckets so stage gates apply correctly.")

    stage_values = sorted([s for s in stage.dropna().unique().tolist() if str(s).strip() != ""])

    suggested_won  = [s for s in stage_values if "won"  in str(s).lower()]
    suggested_lost = [s for s in stage_values if "lost" in str(s).lower()]

    col1, col2 = st.columns(2)

    with col1:
        won_stages = st.multiselect("Won stages", options=stage_values, default=suggested_won)
        lost_stages = st.multiselect("Lost stages", options=stage_values, default=suggested_lost)

    used = set(won_stages + lost_stages)
    remaining2 = [s for s in stage_values if s not in used]

    with col2:
        early_stages = st.multiselect("Early stages", options=remaining2, default=[])
        used2 = used.union(set(early_stages))
        remaining3 = [s for s in stage_values if s not in used2]

        mid_stages = st.multiselect("Mid stages", options=remaining3, default=[])
        used3 = used2.union(set(mid_stages))
        remaining4 = [s for s in stage_values if s not in used3]

        late_stages = st.multiselect("Late stages", options=remaining4, default=[])

    user_mapped_any = any([early_stages, mid_stages, late_stages, won_stages, lost_stages])
    section_end()

    if user_mapped_any:
        won_mask  = stage.isin(won_stages)
        lost_mask = stage.isin(lost_stages)
        early_mask = stage.isin(early_stages)
        mid_mask   = stage.isin(mid_stages)
        late_mask  = stage.isin(late_stages)
        open_mask = ~(won_mask | lost_mask)
    else:
        won_mask = stage.str.contains("Won", case=False, na=False)
        lost_mask = stage.str.contains("Lost", case=False, na=False)
        open_mask = ~won_mask & ~lost_mask
        early_mask = open_mask
        mid_mask = open_mask
        late_mask = open_mask

    # Contact counts per opp
    cr_counts = roles.groupby("Opportunity ID")["Contact ID"].nunique()
    opps = opps.merge(cr_counts.rename("contact_count"), on="Opportunity ID", how="left")
    opps["contact_count"] = pd.to_numeric(opps["contact_count"], errors="coerce").fillna(0)

    # Core KPIs
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

    avg_cr_lost = lost_opps["contact_count"].mean() if not lost_opps.empty else 0
    avg_cr_won = won_opps["contact_count"].mean() if not won_opps.empty else 0
    avg_cr_open = open_opps["contact_count"].mean() if not open_opps.empty else 0

    # Red-flag: Won with zero contacts
    won_zero_df = won_opps[won_opps["contact_count"] == 0].copy()
    won_zero_count = won_zero_df["Opportunity ID"].nunique()
    won_zero_pipeline = won_zero_df["Amount"].sum()
    won_zero_pct = (won_zero_count / won_count) if won_count > 0 else 0

    # Time metrics
    def days_diff(row):
        if pd.isna(row["Created Date"]) or pd.isna(row["Close Date"]):
            return None
        return (row["Close Date"] - row["Created Date"]).days

    lost_opps["days_to_close"] = lost_opps.apply(days_diff, axis=1) if not lost_opps.empty else None
    won_opps["days_to_close"] = won_opps.apply(days_diff, axis=1) if not won_opps.empty else None

    avg_days_lost = lost_opps["days_to_close"].dropna().mean() if "days_to_close" in lost_opps else None
    avg_days_won = won_opps["days_to_close"].dropna().mean() if "days_to_close" in won_opps else None

    today = pd.Timestamp.today().normalize()
    open_opps["age_days"] = (today - open_opps["Created Date"]).dt.days if not open_opps.empty else None
    avg_age_open = open_opps["age_days"].dropna().mean() if "age_days" in open_opps else None

    # Stage Gate Targets
    section_start("Stage Coverage Gates")
    st.caption("Define minimum Contact Roles expected by deal phase (Mid & Late).")

    default_mid_gate  = max(2.0, round(avg_cr_won * 0.6, 1)) if avg_cr_won else 2.0
    default_late_gate = max(3.0, round(avg_cr_won * 0.85, 1)) if avg_cr_won else 3.0

    mid_gate_target = st.slider("Gate 1 — Mid stage minimum contacts", 0.0, 8.0, float(default_mid_gate), 0.5)
    late_gate_target = st.slider("Gate 2 — Late stage minimum contacts", 0.0, 10.0, float(default_late_gate), 0.5)
    section_end()

    # Gate findings
    mid_df = opps[mid_mask].copy() if user_mapped_any else open_opps.copy()
    late_df = opps[late_mask].copy() if user_mapped_any else open_opps.copy()

    mid_below_gate = mid_df[mid_df["contact_count"] < mid_gate_target]
    late_below_gate = late_df[late_df["contact_count"] < late_gate_target]

    mid_below_cnt = mid_below_gate["Opportunity ID"].nunique()
    mid_cnt = mid_df["Opportunity ID"].nunique()
    mid_below_pct = (mid_below_cnt / mid_cnt) if mid_cnt > 0 else 0
    mid_below_pipe = mid_below_gate["Amount"].sum()

    late_below_cnt = late_below_gate["Opportunity ID"].nunique()
    late_cnt = late_df["Opportunity ID"].nunique()
    late_below_pct = (late_below_cnt / late_cnt) if late_cnt > 0 else 0
    late_below_pipe = late_below_gate["Amount"].sum()

    section_start("Stage Gate Coverage Findings")
    st.caption("Quantifies Mid/Late deals below buying-group targets and pipeline at risk.")
    label_with_tooltip("Mid-stage opps below Gate 1", "Mid bucket opps below the Gate 1 contact target.")
    show_value(f"{mid_below_pct:.1%} ({mid_below_cnt:,} of {mid_cnt:,})")
    label_with_tooltip("Mid-stage pipeline below Gate 1", "Total Amount of Mid opps below Gate 1.")
    show_value(f"${mid_below_pipe:,.0f}")

    label_with_tooltip("Late-stage opps below Gate 2", "Late bucket opps below the Gate 2 contact target.")
    show_value(f"{late_below_pct:.1%} ({late_below_cnt:,} of {late_cnt:,})")
    label_with_tooltip("Late-stage pipeline below Gate 2", "Total Amount of Late opps below Gate 2.")
    show_value(f"${late_below_pipe:,.0f}")

    callouts = []
    if mid_below_pct > 0.30:
        callouts.append(
            f"Mid-stage coverage risk: **{mid_below_pct:.0%}** of Mid opps are below the multi-threading gate."
        )
    if late_below_pct > 0.20:
        callouts.append(
            f"Late-stage stakeholder gap: **{late_below_pct:.0%}** of Late opps are below the buying-group gate."
        )
    if avg_cr_won and mid_cnt > 0 and (mid_df["contact_count"].mean() < avg_cr_won * 0.7):
        callouts.append(
            "Mid-stage opps have materially fewer contacts than Won patterns — expanding stakeholders here should improve conversion."
        )

    if callouts:
        st.markdown("**Callouts:**")
        for c in callouts:
            st.markdown(f"- {c}")

    section_end()

    # Pipeline Risk (Open)
    open_df = open_opps.copy()
    open_df["contact_count"] = pd.to_numeric(open_df["contact_count"], errors="coerce").fillna(0)
    open_pipeline = open_opps["Amount"].sum() if not open_opps.empty else 0
    open_pipeline_risk = open_df[open_df["contact_count"] <= 1]["Amount"].sum() if not open_df.empty else 0
    open_opps_risk = open_df[open_df["contact_count"] <= 1]["Opportunity ID"].nunique() if not open_df.empty else 0
    open_opps_total = open_df["Opportunity ID"].nunique() if not open_df.empty else 0
    risk_pct = open_opps_risk / open_opps_total if open_opps_total > 0 else 0

    def modeled_win_rate_for_open(avg_open_contacts, base_win_rate, target_contacts):
        cur = max(avg_open_contacts, 1e-6)
        improvement_factor = min(1.8, target_contacts / cur)
        return min(max(base_win_rate * improvement_factor, base_win_rate), 0.95)

    section_start("Pipeline Risk")
    st.caption("Shows how much open pipeline is under-covered today (0–1 contact roles).")
    label_with_tooltip("Open Pipeline at Risk", "Sum of Amount for open opportunities with 0–1 contact roles.")
    show_value(f"${open_pipeline_risk:,.0f}")
    label_with_tooltip("% of Open Opps Under-Covered", "Open opps with 0–1 roles ÷ total open opps.")
    show_value(f"{risk_pct:.1%} ({open_opps_risk:,} of {open_opps_total:,})")
    section_end()

    # Buying Group Coverage Score
    pct_2plus_open = open_df[open_df["contact_count"] >= 2]["Opportunity ID"].nunique() / open_opps_total if open_opps_total > 0 else 0
    pct_zero_open  = open_df[open_df["contact_count"] == 0]["Opportunity ID"].nunique() / open_opps_total if open_opps_total > 0 else 0
    gap_open_vs_won = max(0, avg_cr_won - avg_cr_open) if avg_cr_won and avg_cr_open is not None else 0

    score = (
        (pct_2plus_open * 60) +
        ((1 - pct_zero_open) * 30) +
        (max(0, 1 - (gap_open_vs_won / max(avg_cr_won, 1))) * 10)
    )
    score = round(min(max(score, 0), 100), 0)
    score_label = "High risk" if score < 40 else ("Needs improvement" if score < 70 else "Healthy")

    section_start("Buying Group Coverage Score")
    st.caption("A single health grade combining depth, zero-coverage penalty, and gap vs won patterns.")
    label_with_tooltip("Coverage Health Score", "Composite from open coverage depth, zero-coverage penalty, and gap vs won.")
    show_value(f"{score:.0f} / 100 — {score_label}")
    section_end()

    # Core Metrics
    section_start("Core Metrics")
    st.caption("Baseline deal volume, pipeline, and current win rate.")
    label_with_tooltip("Total Opportunities", "Unique opportunities in the export.")
    show_value(f"{total_opps:,}")
    label_with_tooltip("Total Pipeline", "Sum of Amount for all opportunities.")
    show_value(f"${total_pipeline:,.0f}")
    label_with_tooltip("Current Win Rate", "Won ÷ (Won + Lost).")
    show_value(f"{win_rate:.1%}")
    section_end()

    # Coverage
    section_start("Contact Role Coverage")
    st.caption("Overall stakeholder coverage and where pipeline is missing roles.")
    label_with_tooltip("Opportunities with Contact Roles", "Unique opps appearing in Contact Roles export.")
    show_value(f"{opps_with_cr:,}")
    label_with_tooltip("Opportunities without Contact Roles", "Total opps minus those with roles.")
    show_value(f"{opps_without_cr:,}")
    label_with_tooltip("Pipeline with Contact Roles", "Amount on opps with ≥1 role.")
    show_value(f"${pipeline_with_cr:,.0f}")
    label_with_tooltip("Pipeline without Contact Roles", "Amount on opps with 0 roles.")
    show_value(f"${pipeline_without_cr:,.0f}")
    label_with_tooltip("Opps with only 1 Contact Role", "Opps where role count = 1.")
    show_value(f"{opps_one_cr:,}")
    label_with_tooltip("Pipeline with only 1 Contact Role", "Amount on opps with exactly 1 role.")
    show_value(f"${pipeline_one_cr:,.0f}")
    section_end()

    # Outcome roles
    section_start("Contact Roles by Outcome")
    st.caption("Buying-group depth differs strongly between won, lost, and open deals.")
    label_with_tooltip("Avg Contact Roles – Won", "Average roles per Won opp.")
    show_value(f"{avg_cr_won:.1f}")
    label_with_tooltip("Avg Contact Roles – Lost", "Average roles per Lost opp.")
    show_value(f"{avg_cr_lost:.1f}")
    label_with_tooltip("Avg Contact Roles – Open", "Average roles per Open opp.")
    show_value(f"{avg_cr_open:.1f}")

    if won_zero_count > 0:
        label_with_tooltip(
            "⚠️ Won Opps with 0 Contact Roles (Red Flag)",
            "Won opportunities that have no buying-group contacts. Indicates CRM hygiene or tracking gaps."
        )
        show_value(f"{won_zero_count:,} ({won_zero_pct:.1%} of Won) — ${won_zero_pipeline:,.0f} pipeline")
    section_end()

    # Time to close
    section_start("Time to Close")
    st.caption("Velocity patterns: wins close faster; stale open deals convert poorly.")
    label_with_tooltip("Avg days to close – Won", "Close Date − Created Date for Won opps.")
    show_value(f"{avg_days_won:.0f} days" if avg_days_won else "0 days")
    label_with_tooltip("Avg days to close – Lost", "Close Date − Created Date for Lost opps.")
    show_value(f"{avg_days_lost:.0f} days" if avg_days_lost else "0 days")
    label_with_tooltip("Avg age of Open opps", "Today − Created Date for Open opps.")
    show_value(f"{avg_age_open:.0f} days" if avg_age_open else "0 days")
    section_end()

    # Simulator
    section_start("Simulator — Target Contact Coverage")
    st.caption("Adjust target buying-group size and see modeled win-rate and revenue lift.")
    target_contacts = st.slider("Target avg contacts on Open Opps", 0.0, 5.0, 2.0, 0.5)
    section_end()

    enhanced_win_rate = modeled_win_rate_for_open(avg_cr_open, win_rate, target_contacts)
    incremental_won_pipeline = max(0, (enhanced_win_rate - win_rate) * open_pipeline)

    section_start("Pipeline Upside from Simulator")
    st.caption("Translates modeled win-rate uplift into incremental won pipeline dollars.")
    label_with_tooltip("Modeled Upside if Coverage Improves", "(Enhanced win rate − current win rate) × open pipeline.")
    show_value(f"${incremental_won_pipeline:,.0f}")
    section_end()

    # Coverage-adjusted forecast
    section_start("Coverage-Adjusted Forecast")
    st.caption("A conservative forecast that discounts under-covered open deals.")

    def weight_for_contacts(n):
        if n <= 0: return 0.6
        if n == 1: return 0.8
        return 1.0

    open_df["coverage_weight"] = open_df["contact_count"].apply(weight_for_contacts)

    expected_open_wins_current = win_rate * open_pipeline
    open_df["expected_win_rate_adj"] = win_rate * open_df["coverage_weight"]
    expected_open_wins_adj = (open_df["expected_win_rate_adj"] * open_df["Amount"]).sum()

    coverage_adj_win_rate = expected_open_wins_adj / open_pipeline if open_pipeline > 0 else 0
    forecast_gap = max(0, expected_open_wins_current - expected_open_wins_adj)

    label_with_tooltip(
        "Coverage-Adjusted Win Rate (Open)",
        "Estimated win rate for open deals after accounting for buying-group coverage. Deals with more contact roles are treated as healthier; deals with 0–1 roles are discounted."
    )
    show_value(f"{coverage_adj_win_rate:.1%}")

    label_with_tooltip(
        "Coverage-Adjusted Expected Won Pipeline (Open)",
        "Sum(Amount × adjusted win rate) across open deals."
    )
    show_value(f"${expected_open_wins_adj:,.0f}")

    label_with_tooltip(
        "Coverage Risk Gap",
        "Difference between current expected wins and coverage-adjusted expected wins."
    )
    show_value(f"${forecast_gap:,.0f}")

    section_end()

    # Segment uplift
    section_start("Segment Uplift (Deal Size Bands)")
    st.caption("Shows where coverage uplift is concentrated by deal size.")

    seg_rows = []
    segment_rows_for_pdf = []

    if not open_df.empty:
        q33 = opps["Amount"].quantile(0.33)
        q67 = opps["Amount"].quantile(0.67)

        def size_band(a):
            if a <= q33: return "SMB"
            if a <= q67: return "Mid-Market"
            return "Enterprise"

        open_df["Size Band"] = open_df["Amount"].apply(size_band)

        seg = open_df.groupby("Size Band").agg(
            open_pipeline=("Amount", "sum"),
            avg_contacts=("contact_count", "mean"),
            opps=("Opportunity ID", "nunique")
        ).reset_index()

        for _, r in seg.iterrows():
            band = r["Size Band"]
            band_open_pipeline = r["open_pipeline"]
            band_avg_contacts = r["avg_contacts"]

            band_enhanced_win_rate = modeled_win_rate_for_open(band_avg_contacts, win_rate, target_contacts)
            band_incremental = max(0, (band_enhanced_win_rate - win_rate) * band_open_pipeline)

            seg_rows.append({
                "Segment": band,
                "Open Pipeline": band_open_pipeline,
                "Avg Contacts (Open)": band_avg_contacts,
                "Incremental Won Pipeline": band_incremental
            })

            segment_rows_for_pdf.append([
                band,
                f"${band_open_pipeline:,.0f}",
                f"{band_avg_contacts:.1f}",
                f"${band_incremental:,.0f}"
            ])

        seg_df = pd.DataFrame(seg_rows).sort_values("Open Pipeline", ascending=False)

        st.dataframe(
            seg_df.style.format({
                "Open Pipeline": "${:,.0f}",
                "Avg Contacts (Open)": "{:.1f}",
                "Incremental Won Pipeline": "${:,.0f}"
            }),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.write("Not enough open opportunities to segment.")

    section_end()

    # Executive Summary
    section_start("Executive Summary")
    st.caption("Leadership-ready takeaways summarizing coverage patterns, risks, and modeled uplift.")
    bullets = []

    bullets.append(
        f"Won opportunities average **{avg_cr_won:.1f}** contact roles vs **{avg_cr_lost:.1f}** on lost deals, showing stronger multi-threading in wins."
    )

    if won_zero_count > 0:
        bullets.append(
            f"**Red flag:** {won_zero_count:,} Won opportunities (**{won_zero_pct:.1%} of wins**) show **0 contact roles**, indicating a CRM hygiene or tracking gap."
        )

    if (avg_cr_open < avg_cr_won) or (avg_cr_open < 2.0):
        bullets.append(
            f"Open opportunities average **{avg_cr_open:.1f}** roles. Increasing coverage to at least **{target_contacts:.1f}** roles would align open deals with won patterns."
        )

    if avg_days_won and avg_age_open:
        bullets.append(
            f"Won opportunities close in ~**{avg_days_won:.0f} days**, while open deals are already ~**{avg_age_open:.0f} days** old — older deals without stakeholder depth are less likely to win."
        )

    if user_mapped_any and (mid_cnt > 0 or late_cnt > 0):
        bullets.append(
            f"Stage gates show **{mid_below_pct:.0%}** of Mid-stage and **{late_below_pct:.0%}** of Late-stage pipeline below stakeholder targets, creating preventable late-stage risk."
        )

    bullets.append(
        f"Improving open coverage to **{target_contacts:.1f}** roles could lift win rate from **{win_rate:.1%}** → **{enhanced_win_rate:.1%}**, creating **${incremental_won_pipeline:,.0f}** in incremental modeled won pipeline."
    )

    for b in bullets:
        st.markdown("- " + b)
    section_end()

    # Insights (Charts)
    section_start("Insights")
    st.caption("Visualizes relationships between contact coverage, win rate, pipeline risk, and velocity.")

    chart_df = opps.copy()
    chart_df["Stage Group"] = "Open"
    chart_df.loc[won_mask, "Stage Group
