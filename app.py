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
# PDF builder helpers
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
    # Watermark (light gray)
    c.saveState()
    c.setFont("Helvetica-Bold", 50)
    c.setFillColor(colors.HexColor("#E6EAF0"))
    c.translate(300, 400)
    c.rotate(30)
    c.drawCentredString(0, 0, "RevOps Global")
    c.restoreState()

    # Footer
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
    segment_rows
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

    # Logo
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

    # Metric tables
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

    # Executive summary
    story.append(Paragraph("Executive Summary", styles["H2"]))
    for b in bullets:
        story.append(Paragraph(f"• {html.escape(b)}", styles["Body"]))
    story.append(Spacer(1, 0.12*inch))

    # Insights + embedded charts
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

    # Segment uplift
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

    # Owner rollup
    story.append(Paragraph("Owner Coverage Rollup (Coaching View)", styles["H2"]))
    story.append(Paragraph(
        "Coach the owners with the largest under-covered open pipeline first.",
        styles["Body"]
    ))
    if owner_rollup_rows:
        for r in owner_rollup_rows:
            story.append(Paragraph(f"• {html.escape(r)}", styles["Body"]))
    else:
        story.append(Paragraph("No owner rollup available.", styles["Body"]))
    story.append(Spacer(1, 0.12*inch))

    # Top opps
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

    # CTA
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

# -----------------------
# CSS
# -----------------------
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

/* Tooltip styles */
.tooltip-wrap { display:inline-flex; align-items:center; gap:6px; margin-top:4px; }
.tooltip-icon { position:relative; cursor:help; font-size:14px; opacity:0.9; }
.tooltip-icon .tooltip-text{
  visibility:hidden; width:300px; background:#111827; color:#fff;
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

# -----------------------
# Instructions
# -----------------------
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

    # Clean
    opps["Amount"] = pd.to_numeric(opps["Amount"], errors="coerce").fillna(0)
    opps["Created Date"] = opps["Created Date"].apply(parse_date)
    opps["Close Date"] = opps["Close Date"].apply(parse_date)

    stage = opps["Stage"].fillna("")
    won_mask = stage.str.contains("Won", case=False, na=False)
    lost_mask = stage.str.contains("Lost", case=False, na=False)
    open_mask = ~won_mask & ~lost_mask

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

    # -----------------------
    # Pipeline Risk & Upside
    # -----------------------
    open_df = open_opps.copy()
    open_df["contact_count"] = pd.to_numeric(open_df["contact_count"], errors="coerce").fillna(0)

    open_pipeline = open_opps["Amount"].sum() if not open_opps.empty else 0
    open_pipeline_risk = open_df[open_df["contact_count"] <= 1]["Amount"].sum() if not open_df.empty else 0
    open_opps_risk = open_df[open_df["contact_count"] <= 1]["Opportunity ID"].nunique() if not open_df.empty else 0
    open_opps_total = open_df["Opportunity ID"].nunique() if not open_df.empty else 0
    risk_pct = open_opps_risk / open_opps_total if open_opps_total > 0 else 0

    # -----------------------
    # (2) Base uplift model helpers
    # -----------------------
    contact_influence_ratio = (avg_cr_won / avg_cr_lost) if avg_cr_lost not in (0, None) else 1.0

    def modeled_win_rate_for_open(avg_open_contacts, base_win_rate, target_contacts):
        """
        Simple proportional model:
        improvement_factor = target / current open coverage (capped)
        applied on base win rate (then capped).
        """
        cur = max(avg_open_contacts, 1e-6)
        improvement_factor = min(1.8, target_contacts / cur)
        return min(max(base_win_rate * improvement_factor, base_win_rate), 0.95)

    # -----------------------
    # Pipeline Risk & Upside card (needs enhanced later but ok to show now)
    # -----------------------
    section_start("Pipeline Risk & Upside")
    label_with_tooltip("Open Pipeline at Risk", "Sum of Amount for open opportunities with 0–1 contact roles.")
    show_value(f"${open_pipeline_risk:,.0f}")
    label_with_tooltip("% of Open Opps Under-Covered", "Open opportunities with 0–1 contact roles ÷ total open opportunities.")
    show_value(f"{risk_pct:.1%} ({open_opps_risk:,} of {open_opps_total:,})")
    section_end()

    # -----------------------
    # Buying Group Coverage Score (kept)
    # -----------------------
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
    label_with_tooltip("Coverage Health Score", "Composite score from open coverage depth, zero-coverage penalty, and gap vs won.")
    show_value(f"{score:.0f} / 100 — {score_label}")
    section_end()

    # -----------------------
    # Core Metrics
    # -----------------------
    section_start("Core Metrics")
    label_with_tooltip("Total Opportunities", "Unique opportunities in the export.")
    show_value(f"{total_opps:,}")
    label_with_tooltip("Total Pipeline", "Sum of Amount for all opportunities.")
    show_value(f"${total_pipeline:,.0f}")
    label_with_tooltip("Current Win Rate", "Won ÷ (Won + Lost).")
    show_value(f"{win_rate:.1%}")
    section_end()

    section_start("Contact Role Coverage")
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

    section_start("Contact Roles by Outcome")
    label_with_tooltip("Avg Contact Roles – Won", "Average roles per Won opp.")
    show_value(f"{avg_cr_won:.1f}")
    label_with_tooltip("Avg Contact Roles – Lost", "Average roles per Lost opp.")
    show_value(f"{avg_cr_lost:.1f}")
    label_with_tooltip("Avg Contact Roles – Open", "Average roles per Open opp.")
    show_value(f"{avg_cr_open:.1f}")
    section_end()

    section_start("Time to Close")
    label_with_tooltip("Avg days to close – Won", "Close Date − Created Date for Won opps.")
    show_value(f"{avg_days_won:.0f} days" if avg_days_won else "0 days")
    label_with_tooltip("Avg days to close – Lost", "Close Date − Created Date for Lost opps.")
    show_value(f"{avg_days_lost:.0f} days" if avg_days_lost else "0 days")
    label_with_tooltip("Avg age of Open opps", "Today − Created Date for Open opps.")
    show_value(f"{avg_age_open:.0f} days" if avg_age_open else "0 days")
    section_end()

    # =========================================================
    # MOVED HERE: (5) Simulator target (all open opps)
    # =========================================================
    section_start("Simulator — Target Contact Coverage")
    st.caption("Slide the target avg contacts for open opportunities to see modeled win-rate and revenue impact.")
    target_contacts = st.slider("Target avg contacts on Open Opps", 0.0, 5.0, 2.0, 0.5)
    section_end()

    # Apply simulator target to uplift
    enhanced_win_rate = modeled_win_rate_for_open(avg_cr_open, win_rate, target_contacts)
    incremental_won_pipeline = max(0, (enhanced_win_rate - win_rate) * open_pipeline)

    # Update Pipeline Risk & Upside with upside now that target is known
    section_start("Pipeline Upside from Simulator")
    label_with_tooltip("Modeled Upside if Coverage Improves", "(Enhanced win rate − current win rate) × open pipeline.")
    show_value(f"${incremental_won_pipeline:,.0f}")
    section_end()

    # -----------------------
    # (2) Coverage-Adjusted Forecast (simple weights)
    # -----------------------
    section_start("Coverage-Adjusted Forecast")

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
        "Weighted win rate for open opps using simple coverage factors: 0CR=0.6×, 1CR=0.8×, 2+CR=1.0×."
    )
    show_value(f"{coverage_adj_win_rate:.1%}")

    label_with_tooltip(
        "Coverage-Adjusted Expected Won Pipeline (Open)",
        "Sum(Amount × (current win rate × coverage weight)) for open opps."
    )
    show_value(f"${expected_open_wins_adj:,.0f}")

    label_with_tooltip(
        "Coverage Risk Gap",
        "Current expected won pipeline on open opps minus coverage-adjusted expected won pipeline."
    )
    show_value(f"${forecast_gap:,.0f}")

    section_end()

    # -----------------------
    # (4) Segment uplift by dynamic deal size bands
    # -----------------------
    section_start("Segment Uplift (Deal Size Bands)")
    st.caption("Bands are auto-derived from your Amount distribution (33rd / 67th percentiles).")

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

    # -----------------------
    # Executive Summary
    # -----------------------
    section_start("Executive Summary")
    bullets = []

    bullets.append(
        f"Won opportunities average **{avg_cr_won:.1f}** contact roles vs **{avg_cr_lost:.1f}** on lost deals, "
        f"showing stronger multi-threading in wins."
    )

    if (avg_cr_open < avg_cr_won) or (avg_cr_open < 2.0):
        bullets.append(
            f"Open opportunities average **{avg_cr_open:.1f}** roles. Increasing coverage to at least "
            f"**{target_contacts:.1f}** roles would align open deals with won patterns."
        )

    if avg_days_won and avg_age_open:
        bullets.append(
            f"Won opportunities close in ~**{avg_days_won:.0f} days**, while open deals are already "
            f"~**{avg_age_open:.0f} days** old — older deals without stakeholder depth are less likely to win."
        )

    bullets.append(
        f"Improving open coverage to **{target_contacts:.1f}** roles could lift win rate from "
        f"**{win_rate:.1%}** → **{enhanced_win_rate:.1%}**, creating **${incremental_won_pipeline:,.0f}** "
        f"in incremental modeled won pipeline."
    )

    for b in bullets:
        st.markdown("- " + b)
    section_end()

    # -----------------------
    # Insights (Altair for app)
    # -----------------------
    section_start("Insights")

    chart_df = opps.copy()
    chart_df["Stage Group"] = "Open"
    chart_df.loc[won_mask, "Stage Group"] = "Won"
    chart_df.loc[lost_mask, "Stage Group"] = "Lost"

    chart_df["contact_count"] = pd.to_numeric(chart_df["contact_count"], errors="coerce").fillna(0)
    chart_df["Amount"] = pd.to_numeric(chart_df["Amount"], errors="coerce").fillna(0)

    def contact_bucket(n):
        n = float(n) if pd.notna(n) else 0
        if n <= 0: return "0"
        if n == 1: return "1"
        if n == 2: return "2"
        if n == 3: return "3"
        return "4+"

    chart_df["Contact Bucket"] = chart_df["contact_count"].apply(contact_bucket)
    bucket_order = ["0", "1", "2", "3", "4+"]

    closed_df = chart_df[chart_df["Stage Group"].isin(["Won", "Lost"])]
    winrate_bucket = closed_df.groupby("Contact Bucket").agg(
        won=("Stage Group", lambda s: (s == "Won").sum()),
        lost=("Stage Group", lambda s: (s == "Lost").sum())
    ).reset_index()
    winrate_bucket["Win Rate"] = winrate_bucket.apply(
        lambda r: r["won"] / (r["won"] + r["lost"]) if (r["won"] + r["lost"]) > 0 else 0,
        axis=1
    )

    st.caption("Win rate increases as more stakeholders are engaged.")
    st.altair_chart(
        alt.Chart(winrate_bucket).mark_line(point=True).encode(
            x=alt.X("Contact Bucket:N", sort=bucket_order, title="Contact Roles per Opportunity (bucketed)"),
            y=alt.Y("Win Rate:Q", axis=alt.Axis(format="%"), title="Win Rate"),
            tooltip=["Contact Bucket", alt.Tooltip("Win Rate:Q", format=".1%"), "won", "lost"]
        ).properties(height=260),
        use_container_width=True
    )

    open_chart_df = chart_df[chart_df["Stage Group"] == "Open"].copy()
    open_chart_df["Open Coverage Bucket"] = open_chart_df["contact_count"].apply(
        lambda n: "0 Contact Roles" if n == 0 else ("1 Contact Role" if n == 1 else "2+ Contact Roles")
    )
    open_pipeline_bucket = open_chart_df.groupby("Open Coverage Bucket")["Amount"].sum().reindex(
        ["0 Contact Roles", "1 Contact Role", "2+ Contact Roles"]
    ).fillna(0).reset_index().rename(columns={"Amount": "Open Pipeline"})

    st.caption("How much open pipeline is under-covered today.")
    st.altair_chart(
        alt.Chart(open_pipeline_bucket).mark_bar().encode(
            x=alt.X("Open Coverage Bucket:N", title="Open Opportunity Coverage"),
            y=alt.Y("Open Pipeline:Q", title="Open Pipeline ($)"),
            tooltip=["Open Coverage Bucket", alt.Tooltip("Open Pipeline:Q", format=",.0f")]
        ).properties(height=260),
        use_container_width=True
    )

    time_df = chart_df.copy()
    time_df["days_to_close"] = time_df.apply(days_diff, axis=1)
    time_df["open_age_days"] = None
    open_mask_local = (time_df["Stage Group"] == "Open") & time_df["Created Date"].notna()
    time_df.loc[open_mask_local, "open_age_days"] = (today - time_df.loc[open_mask_local, "Created Date"]).dt.days

    agg_rows = []
    for sg in ["Won", "Lost"]:
        tmp = time_df[time_df["Stage Group"] == sg]
        grp = tmp.groupby("Contact Bucket")["days_to_close"].mean().reindex(bucket_order).reset_index()
        grp["Stage Group"] = sg
        grp = grp.rename(columns={"days_to_close": "Avg Days"})
        agg_rows.append(grp)

    tmp_open = time_df[time_df["Stage Group"] == "Open"]
    grp_open = tmp_open.groupby("Contact Bucket")["open_age_days"].mean().reindex(bucket_order).reset_index()
    grp_open["Stage Group"] = "Open"
    grp_open = grp_open.rename(columns={"open_age_days": "Avg Days"})
    agg_rows.append(grp_open)

    avg_days_bucket = pd.concat(agg_rows, ignore_index=True)

    st.caption("More contact roles correlate with faster closes and younger open pipeline.")
    st.altair_chart(
        alt.Chart(avg_days_bucket).mark_bar().encode(
            x=alt.X("Contact Bucket:N", sort=bucket_order, title="Contact Roles per Opportunity (bucketed)"),
            y=alt.Y("Avg Days:Q", title="Avg Days (Close for Won/Lost, Age for Open)"),
            color=alt.Color("Stage Group:N", title="Outcome"),
            tooltip=["Stage Group", "Contact Bucket", alt.Tooltip("Avg Days:Q", format=",.0f")]
        ).properties(height=300),
        use_container_width=True
    )

    impact_df = pd.DataFrame({
        "Scenario": ["Current Win Rate", "Enhanced Win Rate"],
        "Win Rate": [win_rate, enhanced_win_rate]
    })
    st.caption("Modeled uplift if open opportunities match target coverage.")
    st.altair_chart(
        alt.Chart(impact_df).mark_bar().encode(
            x=alt.X("Scenario:N", title=""),
            y=alt.Y("Win Rate:Q", axis=alt.Axis(format="%"), title="Win Rate"),
            tooltip=["Scenario", alt.Tooltip("Win Rate:Q", format=".1%")]
        ).properties(height=240),
        use_container_width=True
    )

    section_end()

    # -----------------------
    # Owner rollups + guidance
    # -----------------------
    owner_rollup_rows = []
    if "Opportunity Owner" in opps.columns:
        section_start("Owner Coverage Rollup (Coaching View)")
        st.caption("Coach owners with the largest under-covered open pipeline first.")

        owner_df = open_opps.copy()
        owner_df["Opportunity Owner"] = owner_df["Opportunity Owner"].fillna("Unassigned")

        roll = owner_df.groupby("Opportunity Owner").agg(
            open_opps=("Opportunity ID", "nunique"),
            avg_contacts=("contact_count", "mean"),
            undercovered_pct=("contact_count", lambda s: (s <= 1).mean()),
            open_pipeline=("Amount", "sum")
        ).reset_index().sort_values("open_pipeline", ascending=False).head(8)

        for _, r in roll.iterrows():
            line = (
                f"{r['Opportunity Owner']} — "
                f"{int(r['open_opps'])} open opps, "
                f"avg contacts {r['avg_contacts']:.1f}, "
                f"under-covered {r['undercovered_pct']:.0%}, "
                f"open pipeline ${r['open_pipeline']:,.0f}"
            )
            owner_rollup_rows.append(line)
            st.markdown(f"- **{line}**")
        section_end()

    # -----------------------
    # Top Open Opps + guidance
    # -----------------------
    section_start("Top Open Opportunities to Fix First")
    st.caption("Prioritize multi-threading these deals first based on value, age, and low contact coverage.")

    top_opps_rows = []
    if not open_df.empty:
        tmp = open_df.copy()
        tmp["age_days"] = (today - tmp["Created Date"]).dt.days
        tmp["age_days"] = pd.to_numeric(tmp["age_days"], errors="coerce").fillna(0)

        tmp["priority_score"] = (
            (tmp["Amount"] / (tmp["Amount"].max() if tmp["Amount"].max() > 0 else 1)) * 0.6 +
            (tmp["age_days"] / (tmp["age_days"].max() if tmp["age_days"].max() > 0 else 1)) * 0.3 +
            ((1 - (tmp["contact_count"] / (tmp["contact_count"].max() if tmp["contact_count"].max() > 0 else 1))) * 0.1)
        )

        top_fix = tmp.sort_values("priority_score", ascending=False).head(10)
        for _, r in top_fix.iterrows():
            line = (
                f"{r.get('Opportunity Name','')} (Owner: {r.get('Opportunity Owner','')}) — "
                f"${r.get('Amount',0):,.0f}, {r.get('contact_count',0):.0f} contacts, "
                f"{r.get('age_days',0):.0f} days open"
            )
            top_opps_rows.append(line)
            st.markdown(f"- **{line}**")
    else:
        st.write("No open opportunities found.")
    section_end()

    # -----------------------
    # Build matplotlib charts for PDF
    # -----------------------
    pdf_chart_pngs = []

    fig1 = plt.figure(figsize=(7.2, 3.2))
    ax1 = fig1.add_subplot(111)
    ax1.plot(winrate_bucket["Contact Bucket"], winrate_bucket["Win Rate"], marker="o")
    ax1.set_title("Win Rate vs Contact Roles (bucketed)")
    ax1.set_xlabel("Contact Roles per Opportunity")
    ax1.set_ylabel("Win Rate")
    ax1.yaxis.set_major_formatter(PercentFormatter(1.0))
    pdf_chart_pngs.append(fig_to_png_bytes(fig1))

    fig2 = plt.figure(figsize=(7.2, 3.2))
    ax2 = fig2.add_subplot(111)
    ax2.bar(open_pipeline_bucket["Open Coverage Bucket"], open_pipeline_bucket["Open Pipeline"])
    ax2.set_title("Open Pipeline by Contact Coverage")
    ax2.set_xlabel("Open Opportunity Coverage")
    ax2.set_ylabel("Open Pipeline ($)")
    ax2.tick_params(axis='x', rotation=10)
    pdf_chart_pngs.append(fig_to_png_bytes(fig2))

    fig3 = plt.figure(figsize=(7.2, 3.6))
    ax3 = fig3.add_subplot(111)
    for sg in ["Won", "Lost", "Open"]:
        sub = avg_days_bucket[avg_days_bucket["Stage Group"] == sg]
        ax3.plot(sub["Contact Bucket"], sub["Avg Days"], marker="o", label=sg)
    ax3.set_title("Velocity vs Contact Roles (Won/Lost close days, Open age)")
    ax3.set_xlabel("Contact Roles per Opportunity")
    ax3.set_ylabel("Avg Days")
    ax3.legend()
    pdf_chart_pngs.append(fig_to_png_bytes(fig3))

    fig4 = plt.figure(figsize=(6.5, 3.0))
    ax4 = fig4.add_subplot(111)
    ax4.bar(impact_df["Scenario"], impact_df["Win Rate"])
    ax4.set_title("Modeled Win Rate Uplift")
    ax4.set_ylabel("Win Rate")
    ax4.yaxis.set_major_formatter(PercentFormatter(1.0))
    pdf_chart_pngs.append(fig_to_png_bytes(fig4))

    # -----------------------
    # Download Full PDF
    # -----------------------
    section_start("Download Full PDF Report")

    metrics_dict = {
        "Pipeline Risk & Upside": [
            ["Open Pipeline at Risk", f"${open_pipeline_risk:,.0f}"],
            ["% Open Opps Under-Covered", f"{risk_pct:.1%}"],
            ["Incremental Won Pipeline (modeled)", f"${incremental_won_pipeline:,.0f}"],
        ],
        "Core Metrics": [
            ["Total Opportunities", f"{total_opps:,}"],
            ["Total Pipeline", f"${total_pipeline:,.0f}"],
            ["Current Win Rate", f"{win_rate:.1%}"],
        ],
        "Contact Coverage": [
            ["Opps with Contact Roles", f"{opps_with_cr:,}"],
            ["Opps without Contact Roles", f"{opps_without_cr:,}"],
            ["Pipeline with Contact Roles", f"${pipeline_with_cr:,.0f}"],
            ["Pipeline without Contact Roles", f"${pipeline_without_cr:,.0f}"],
            ["Opps with only 1 Contact Role", f"{opps_one_cr:,}"],
            ["Pipeline with only 1 Contact Role", f"${pipeline_one_cr:,.0f}"],
        ],
        "Contact Roles by Outcome": [
            ["Avg Contact Roles – Won", f"{avg_cr_won:.1f}"],
            ["Avg Contact Roles – Lost", f"{avg_cr_lost:.1f}"],
            ["Avg Contact Roles – Open", f"{avg_cr_open:.1f}"],
        ],
        "Time to Close": [
            ["Avg days to close – Won", f"{avg_days_won:.0f} days" if avg_days_won else "0 days"],
            ["Avg days to close – Lost", f"{avg_days_lost:.0f} days" if avg_days_lost else "0 days"],
            ["Avg age of Open opps", f"{avg_age_open:.0f} days" if avg_age_open else "0 days"],
        ],
        "Coverage-Adjusted Forecast": [
            ["Coverage-Adjusted Win Rate (Open)", f"{coverage_adj_win_rate:.1%}"],
            ["Coverage-Adjusted Expected Won Pipeline (Open)", f"${expected_open_wins_adj:,.0f}"],
            ["Coverage Risk Gap", f"${forecast_gap:,.0f}"],
        ],
        "Modeled Uplift (Simulator Target)": [
            ["Target Avg Contacts (Open)", f"{target_contacts:.1f}"],
            ["Enhanced Win Rate (modeled)", f"{enhanced_win_rate:.1%}"],
            ["Incremental Won Pipeline (modeled)", f"${incremental_won_pipeline:,.0f}"],
        ],
    }

    pdf_bytes = build_pdf_report(
        metrics_dict=metrics_dict,
        bullets=bullets,
        owner_rollup_rows=owner_rollup_rows,
        top_opps_rows=top_opps_rows,
        chart_pngs=pdf_chart_pngs,
        segment_rows=segment_rows_for_pdf
    )

    st.download_button(
        "Download PDF Report (with charts)",
        data=pdf_bytes,
        file_name="RevOps_Global_Opportunity_Contact_Role_Insights.pdf",
        mime="application/pdf"
    )
    section_end()

    # CTA
    section_start("Buying Group Automation")
    st.markdown(
        f"""
RevOps Global’s **Buying Group Automation** helps sales teams identify stakeholders, close coverage gaps,  
and multi-thread deals faster — directly improving Contact Role coverage and conversion.

👉 **Learn more here:**  
{CTA_URL}
        """
    )
    section_end()

else:
    st.info("Upload both CSV files above to generate insights.")


# Footer
st.markdown(
    f"""
<hr style="margin-top:26px; border:0; border-top:1px solid #e5e7eb;" />
<div style="font-size:12px; color:#6b7280; text-align:center; padding:10px 0 2px 0;">
  © {datetime.now().year} RevOps Global. All rights reserved.
</div>
    """,
    unsafe_allow_html=True
)
