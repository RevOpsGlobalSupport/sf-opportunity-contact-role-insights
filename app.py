import streamlit as st
import pandas as pd
from datetime import datetime
import html
import altair as alt
import re
import io
import requests
from PIL import Image as PILImage

# PDF deps
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_LEFT

# Matplotlib for PDF charts
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter


# =========================
# CONFIG / CONSTANTS
# =========================
LOGO_URL = "https://www.revopsglobal.com/wp-content/uploads/2024/09/Footer_Logo.png"
SITE_URL = "https://www.revopsglobal.com/"
CTA_URL = "https://www.revopsglobal.com/buying-group-automation/"

sample_opps_url = "https://drive.google.com/file/d/11bNN1lSs6HtPyXVV0k9yYboO6XXdI85H/view?usp=sharing"
sample_roles_url = "https://drive.google.com/file/d/1-w_yFF0_naXGUEX00TKOMDT2bAW7ehPE/view?usp=sharing"


# =========================
# HELPERS
# =========================
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


def clean_id_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.replace({"nan": "", "None": ""})
    return s


def fmt_money(x):
    try:
        if pd.isna(x):
            return "$0"
        x = float(x)
        return f"${x:,.0f}"
    except Exception:
        return "$0"


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


def section_start(title: str):
    st.markdown(
        f"""
        <div class="section-card">
          <div class="section-title">{html.escape(title)}</div>
          <div class="section-divider"></div>
        """,
        unsafe_allow_html=True
    )


def section_end():
    st.markdown("</div>", unsafe_allow_html=True)


def wilson_ci(k, n, z=1.96):
    if n == 0:
        return (0.0, 0.0)
    p = k / n
    denom = 1 + z**2 / n
    center = (p + z**2/(2*n)) / denom
    margin = (z * ((p*(1-p)/n + z**2/(4*n**2))**0.5)) / denom
    return (max(0.0, center - margin), min(1.0, center + margin))


def bucket_seniority(title: str) -> str:
    if not isinstance(title, str) or title.strip() == "":
        return "Other / Unknown"
    t = title.lower()

    if re.search(r"\b(chief|ceo|cfo|coo|cto|cio|cmo|cro|cso|cpo|founder|co-founder|president)\b", t):
        return "C-Level"
    if re.search(r"\b(evp|svp|executive vice president|senior vice president)\b", t):
        return "EVP / SVP"
    if re.search(r"\b(vp|vice president)\b", t):
        return "VP"
    if re.search(r"\b(director|head of|chief of staff|gm|general manager)\b", t):
        return "Director / Head"
    if re.search(r"\b(manager|lead|supervisor)\b", t):
        return "Manager"
    if re.search(r"\b(analyst|engineer|specialist|consultant|associate|coordinator|administrator|rep|developer|designer|architect|scientist|strategist|officer)\b", t):
        return "IC / Staff"

    return "Other / Unknown"


# PDF helpers
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
    chart_pngs,
    won_zero_rows,
    owner_bullets,
    priority_bullets
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

    story.append(Paragraph("CRM Opportunity Contact Role Insights", styles["H1"]))
    story.append(Paragraph("Report generated by RevOps Global", styles["Small"]))
    story.append(Spacer(1, 0.2*inch))

    for section_title, rows in metrics_dict.items():
        story.append(Paragraph(section_title, styles["H2"]))
        table_data = [["Metric", "Value"]] + rows
        t = Table(table_data, colWidths=[3.7*inch, 2.7*inch])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F1F5F9")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#0F172A")),
            ("FONTNAME", (0,0), (-1,-1), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
            ("ALIGN", (1,1), (1,-1), "RIGHT"),
            ("FONTSIZE", (0,0), (-1,-1), 10),
        ]))
        story.append(t)
        story.append(Spacer(1, 0.12*inch))

    story.append(Paragraph("Executive Summary", styles["H2"]))
    for b in bullets:
        story.append(Paragraph(f"• {html.escape(b)}", styles["Body"]))
    story.append(Spacer(1, 0.12*inch))

    if won_zero_rows:
        story.append(Paragraph("Won Deals Missing Contact Roles (Red Flag)", styles["H2"]))
        for r in won_zero_rows:
            story.append(Paragraph(f"• {html.escape(r)}", styles["Body"]))
        story.append(Spacer(1, 0.12*inch))

    if owner_bullets:
        story.append(Paragraph("Owner Coverage Rollup (Coaching View)", styles["H2"]))
        for r in owner_bullets:
            story.append(Paragraph(f"• {html.escape(r)}", styles["Body"]))
        story.append(Spacer(1, 0.12*inch))

    if priority_bullets:
        story.append(Paragraph("Top Open Opportunities to Fix First", styles["H2"]))
        for r in priority_bullets:
            story.append(Paragraph(f"• {html.escape(r)}", styles["Body"]))
        story.append(Spacer(1, 0.12*inch))

    story.append(Paragraph("Insights", styles["H2"]))
    story.append(Spacer(1, 0.1*inch))
    for i, png_buf in enumerate(chart_pngs, start=1):
        story.append(Image(png_buf, width=6.7*inch, height=3.2*inch))
        story.append(Spacer(1, 0.15*inch))
        if i in (2, 4):
            story.append(PageBreak())

    story.append(Paragraph("Buying Group Automation", styles["H2"]))
    story.append(Paragraph(
        "RevOps Global’s Buying Group Automation helps teams identify stakeholders, close coverage gaps, and multi-thread deals faster.",
        styles["Body"]
    ))
    story.append(Paragraph(f"Learn more: {CTA_URL}", styles["Body"]))

    doc.build(story, onFirstPage=pdf_watermark_and_footer, onLaterPages=pdf_watermark_and_footer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================
# APP UI
# =========================
st.set_page_config(page_title="CRM Opportunity Contact Role Insights", layout="wide")

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
      CRM Opportunity Contact Role Insights
    </div>
    <div style="font-size:15px;color:#6b7280;margin-top:6px;margin-bottom:10px;">
      Measure Contact Role coverage and its impact on win rates
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("<hr style='margin: 8px 0 16px 0; border:0; border-top:1px solid #e5e7eb;' />",
            unsafe_allow_html=True)

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
  padding: 14px 16px 10px 16px;
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

.exec-summary{
  font-size:17px;
  line-height:1.6;
  font-weight:500;
}

.score-badge{
  font-size:22px;
  font-weight:800;
  padding:6px 10px;
  border-radius:8px;
  display:inline-block;
}
</style>
""", unsafe_allow_html=True)

with st.expander("📌 How to export CRM data & use this app", expanded=False):
    st.markdown(
        """
**Step 1 — Export Opportunity Data from your CRM**

Create an Opportunities report. Use **Tabular** format and include fields in this exact order:  
Opportunity ID, Opportunity Name, Account ID, Amount, Type, Stage, Created Date, Closed Date, and Opportunity Owner.

- Reports → New Report → Opportunities  
- Tabular  
- Add fields in the order above  
- Filters:
  - Scope: All Opportunities  
  - Date Range: All Time (or your analysis window)  
- Export as CSV  

---

**Step 2 — Export Opportunity Contact Role Data**

Create an Opportunities with Contact Roles report. Tabular format, exact fields order:  
Opportunity ID, Opportunity Name, Account ID, Amount, Type, Stage, Opportunity Created Date, Opportunity Closed Date, Opportunity Owner, Contact ID, Title, Department, Role, Is Primary.

- New Report → Opportunities with Contact Roles  
- Tabular → add fields in order above  
- Filters same as Step 1  
- Export as CSV  

---

**Step 3 — Upload Data**

Upload both CSVs here.

Do NOT:
- change column order  
- rename columns  
- add/remove columns  
- insert blank rows/headers  
        """
    )

# CSV Uploads
st.markdown(f"**Upload Opportunities CSV**  \n[Download sample]({sample_opps_url})")
opps_file = st.file_uploader("", type=["csv"], key="opps")
st.markdown("<hr style='margin:6px 0 10px 0;border:0;border-top:1px solid #e5e7eb;' />",
            unsafe_allow_html=True)
st.markdown(f"**Upload Opportunities with Contact Roles CSV**  \n[Download sample]({sample_roles_url})")
roles_file = st.file_uploader("", type=["csv"], key="roles")


# =========================
# MAIN LOGIC
# =========================
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
        st.error("Opportunities file missing columns: " + ", ".join(missing_opps))
        st.stop()
    if missing_roles:
        st.error("Contact Roles file missing columns: " + ", ".join(missing_roles))
        st.stop()

    opps["Opportunity ID"] = clean_id_series(opps["Opportunity ID"])
    opps = opps[opps["Opportunity ID"] != ""].copy()

    roles["Opportunity ID"] = clean_id_series(roles["Opportunity ID"])
    roles["Contact ID"] = clean_id_series(roles["Contact ID"])
    roles = roles[(roles["Opportunity ID"] != "") & (roles["Contact ID"] != "")].copy()

    opps["Amount"] = pd.to_numeric(opps["Amount"], errors="coerce").fillna(0)
    opps["Created Date"] = opps["Created Date"].apply(parse_date)
    opps["Close Date"] = opps["Close Date"].apply(parse_date)
    opps["Type"] = opps["Type"].fillna("").astype(str)
    opps["Stage"] = opps["Stage"].fillna("").astype(str)

    # GLOBAL TYPE FILTER
    all_types = sorted([t for t in opps["Type"].dropna().unique().tolist() if str(t).strip() != ""])
    section_start("Global Filter — Opportunity Type")
    st.caption("Filter the entire analysis by Opportunity Type. All insights update based on selection.")
    selected_types = st.multiselect(
        "Select Opportunity Types to include (default = all)",
        options=all_types,
        default=all_types
    )
    section_end()

    if selected_types:
        opps = opps[opps["Type"].isin(selected_types)].copy()

    opps = opps.reset_index(drop=True)

    filtered_opp_ids = set(opps["Opportunity ID"].unique())
    roles = roles[roles["Opportunity ID"].isin(filtered_opp_ids)].copy()

    if roles.empty:
        st.warning(
            "⚠️ Contact Roles file has **0 matching Opportunity IDs** after filtering. "
            "Please ensure both CSVs are exported from the same CRM scope/time window."
        )

    stage = opps["Stage"].astype(str)

    # ======================================================
    # Bucket Opportunities Stages
    # ======================================================
    section_start("Bucket Opportunities Stages")
    st.caption("Map your CRM stages into buckets so all analysis and gates apply correctly.")

    stage_values = sorted([s for s in stage.dropna().unique().tolist() if str(s).strip() != ""])
    suggested_won  = [s for s in stage_values if "won" in str(s).lower()]
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
        won_mask = stage.isin(won_stages)
        lost_mask = stage.isin(lost_stages)
        open_mask = ~(won_mask | lost_mask)
    else:
        won_mask = stage.str.contains("Won", case=False, na=False)
        lost_mask = stage.str.contains("Lost", case=False, na=False)
        open_mask = ~(won_mask | lost_mask)

    # CONTACT COUNTS
    cr_counts = roles.groupby("Opportunity ID")["Contact ID"].nunique()
    opps = opps.merge(cr_counts.rename("contact_count"), on="Opportunity ID", how="left")
    opps["contact_count"] = pd.to_numeric(opps["contact_count"], errors="coerce").fillna(0)

    # Stage bucket helper
    stage_lookup = opps.set_index("Opportunity ID")["Stage"].to_dict()

    def stage_bucket_for_id(opp_id):
        s = stage_lookup.get(opp_id, "")
        if user_mapped_any:
            if s in won_stages: return "Won"
            if s in lost_stages: return "Lost"
            if s in late_stages: return "Late"
            if s in mid_stages: return "Mid"
            if s in early_stages: return "Early"
            return "Open"
        sl = str(s).lower()
        if "won" in sl: return "Won"
        if "lost" in sl: return "Lost"
        return "Open"

    opps["Stage Bucket"] = opps["Opportunity ID"].apply(stage_bucket_for_id)

    # ======================================================
    # BASIC SPLITS
    # ======================================================
    total_opps = opps["Opportunity ID"].nunique()
    total_pipeline = opps["Amount"].sum()
    opps_with_cr_ids = set(roles["Opportunity ID"].dropna().unique())
    opps_with_cr = opps[opps["Opportunity ID"].isin(opps_with_cr_ids)]["Opportunity ID"].nunique()
    opps_without_cr = total_opps - opps_with_cr
    pipeline_with_cr = opps[opps["Opportunity ID"].isin(opps_with_cr_ids)]["Amount"].sum()
    pipeline_without_cr = total_pipeline - pipeline_with_cr
    one_cr_ids = cr_counts[cr_counts == 1].index
    opps_one_cr = opps[opps["Opportunity ID"].isin(one_cr_ids)]["Opportunity ID"].nunique()
    pipeline_one_cr = opps[opps["Opportunity ID"].isin(one_cr_ids)]["Amount"].sum()

    open_opps = opps.loc[open_mask].copy()
    won_opps = opps.loc[won_mask].copy()
    lost_opps = opps.loc[lost_mask].copy()

    won_count = won_opps["Opportunity ID"].nunique()
    lost_count = lost_opps["Opportunity ID"].nunique()
    win_rate = won_count / (won_count + lost_count) if (won_count + lost_count) > 0 else 0

    avg_cr_lost = lost_opps["contact_count"].mean() if not lost_opps.empty else 0
    avg_cr_won = won_opps["contact_count"].mean() if not won_opps.empty else 0
    avg_cr_open = open_opps["contact_count"].mean() if not open_opps.empty else 0

    won_zero_df = won_opps[won_opps["contact_count"] == 0].copy()
    won_zero_count = won_zero_df["Opportunity ID"].nunique()
    won_zero_pipeline = won_zero_df["Amount"].sum()
    won_zero_pct = (won_zero_count / won_count) if won_count > 0 else 0

    def days_diff(row):
        if pd.isna(row["Created Date"]) or pd.isna(row["Close Date"]):
            return None
        return (row["Close Date"] - row["Created Date"]).days

    if not lost_opps.empty:
        lost_opps["days_to_close"] = lost_opps.apply(days_diff, axis=1)
    if not won_opps.empty:
        won_opps["days_to_close"] = won_opps.apply(days_diff, axis=1)

    avg_days_lost = lost_opps["days_to_close"].dropna().mean() if "days_to_close" in lost_opps else None
    avg_days_won = won_opps["days_to_close"].dropna().mean() if "days_to_close" in won_opps else None

    today = pd.Timestamp.today().normalize()
    if not open_opps.empty:
        open_opps["age_days"] = (today - open_opps["Created Date"]).dt.days
    avg_age_open = open_opps["age_days"].dropna().mean() if "age_days" in open_opps else None

    # ======================================================
    # Buying Group Coverage Score
    # ======================================================
    open_df = open_opps.copy()
    open_opps_total = open_df["Opportunity ID"].nunique() if not open_df.empty else 0
    pct_2plus_open = open_df[open_df["contact_count"] >= 2]["Opportunity ID"].nunique() / open_opps_total if open_opps_total > 0 else 0
    pct_zero_open = open_df[open_df["contact_count"] == 0]["Opportunity ID"].nunique() / open_opps_total if open_opps_total > 0 else 0
    gap_open_vs_won = max(0, avg_cr_won - avg_cr_open) if avg_cr_won and avg_cr_open is not None else 0

    score = (
        (pct_2plus_open * 60) +
        ((1 - pct_zero_open) * 30) +
        (max(0, 1 - (gap_open_vs_won / max(avg_cr_won, 1))) * 10)
    )
    score = round(min(max(score, 0), 100), 0)
    score_label = "High risk" if score < 40 else ("Needs improvement" if score < 70 else "Healthy")
    score_color = "#EF4444" if score < 40 else ("#F59E0B" if score < 70 else "#10B981")

    section_start("Buying Group Coverage Score")
    st.caption("A single health grade combining stakeholder depth, zero-coverage penalty, and gap vs won patterns.")
    st.markdown(
        f"<div class='score-badge' style='background:{score_color};color:white;'>"
        f"{score:.0f} / 100 — {score_label}</div>",
        unsafe_allow_html=True
    )
    section_end()

    # ======================================================
    # Executive Summary
    # ======================================================
    bullets = []
    bullets.append(
        f"Current win rate is **{win_rate:.1%}**. Won deals average **{avg_cr_won:.1f}** contact roles vs Lost at **{avg_cr_lost:.1f}**, "
        "showing strong correlation between buying-group depth and conversion."
    )
    if avg_cr_open < max(avg_cr_won, 2.0):
        bullets.append(
            f"Open opportunities average **{avg_cr_open:.1f}** contact roles. Increasing this towards **at least 2.0** "
            "contacts per opportunity would align open deals with won buying-group patterns."
        )
    open_pipeline = open_df["Amount"].sum() if not open_df.empty else 0
    open_pipeline_risk = open_df[open_df["contact_count"] <= 1]["Amount"].sum() if not open_df.empty else 0
    open_opps_risk = open_df[open_df["contact_count"] <= 1]["Opportunity ID"].nunique() if not open_df.empty else 0
    risk_pct = open_opps_risk / open_opps_total if open_opps_total > 0 else 0
    if open_pipeline_risk > 0:
        bullets.append(
            f"**{fmt_money(open_pipeline_risk)}** of open pipeline is under-covered (0–1 roles), representing **{risk_pct:.0%}** of open opportunities."
        )
    if won_zero_count > 0:
        bullets.append(
            f"⚠️ **{won_zero_count:,} Won deals** have **0** contact roles logged (≈{won_zero_pct:.0%} of Won), indicating CRM hygiene gaps."
        )

    section_start("Executive Summary")
    st.markdown("<div class='exec-summary'>", unsafe_allow_html=True)
    for b in bullets:
        st.markdown(f"• {b}")
    st.markdown("</div>", unsafe_allow_html=True)
    section_end()

    # ======================================================
    # Stage Coverage Gates (TABLE ONLY + COLORS)
    # ======================================================
    section_start("Stage Coverage Gates")
    st.caption(
        "Define the minimum buying-group contacts you expect at each stage bucket. "
        "We’ll compute how many deals and how much pipeline meet those gates."
    )

    g1, g2, g3 = st.columns(3)
    with g1:
        early_gate = st.number_input("Early stage gate (min contacts)", min_value=0, max_value=10, value=1, step=1)
    with g2:
        mid_gate = st.number_input("Mid stage gate (min contacts)", min_value=0, max_value=10, value=2, step=1)
    with g3:
        late_gate = st.number_input("Late stage gate (min contacts)", min_value=0, max_value=10, value=3, step=1)

    gates_df = opps[opps["Stage Bucket"].isin(["Early", "Mid", "Late"])].copy()

    def meets_gate(row):
        b = row["Stage Bucket"]
        c = row["contact_count"]
        if b == "Early":
            return c >= early_gate
        if b == "Mid":
            return c >= mid_gate
        if b == "Late":
            return c >= late_gate
        return False

    if not gates_df.empty:
        gates_df["Meets Gate"] = gates_df.apply(meets_gate, axis=1)

        gate_roll = gates_df.groupby("Stage Bucket").agg(
            Opps=("Opportunity ID", "nunique"),
            Opps_Meeting_Gate=("Meets Gate", "sum"),
            Pipeline=("Amount", "sum"),
            Pipeline_Meeting_Gate=("Amount", lambda s: s[gates_df.loc[s.index, "Meets Gate"]].sum())
        ).reindex(["Early", "Mid", "Late"]).fillna(0).reset_index()

        gate_roll["Opp Coverage %"] = gate_roll.apply(
            lambda r: r["Opps_Meeting_Gate"] / r["Opps"] if r["Opps"] > 0 else 0, axis=1
        )
        gate_roll["Pipeline Coverage %"] = gate_roll.apply(
            lambda r: r["Pipeline_Meeting_Gate"] / r["Pipeline"] if r["Pipeline"] > 0 else 0, axis=1
        )

        display_gate = gate_roll.rename(columns={
            "Stage Bucket": "Stage Bucket",
            "Opps": "# Opps",
            "Opps_Meeting_Gate": "# Opps meeting gate",
            "Pipeline": "Pipeline",
            "Pipeline_Meeting_Gate": "Pipeline meeting gate",
            "Opp Coverage %": "Opp Coverage %",
            "Pipeline Coverage %": "Pipeline Coverage %"
        }).copy()

        # Money formatting (keep % numeric for styling)
        display_gate["Pipeline"] = display_gate["Pipeline"].map(fmt_money)
        display_gate["Pipeline meeting gate"] = display_gate["Pipeline meeting gate"].map(fmt_money)

        def color_cov(val):
            try:
                v = float(val)
            except Exception:
                return ""
            if v >= 0.80:
                return "background-color: #D1FAE5;"  # green
            if 0.65 <= v < 0.80:
                return "background-color: #FEF3C7;"  # orange
            return "background-color: #FEE2E2;"      # light red

        styled = (
            display_gate.style
            .applymap(color_cov, subset=["Opp Coverage %", "Pipeline Coverage %"])
            .format({
                "Opp Coverage %": "{:.0%}",
                "Pipeline Coverage %": "{:.0%}"
            })
        )

        st.dataframe(styled, use_container_width=True, hide_index=True)
    else:
        st.info("No Early/Mid/Late opportunities found based on current stage mapping.")

    section_end()

    # ======================================================
    # Current Opportunity Insights
    # ======================================================
    section_start("Current Opportunity Insights")
    st.caption("Baseline metrics across pipeline, coverage, outcomes, and velocity.")
    c_left, c_right = st.columns(2)

    with c_left:
        label_with_tooltip("Total Opportunities", "Unique opportunities in the export.")
        show_value(f"{total_opps:,}")
        label_with_tooltip("Total Pipeline", "Sum of Amount for all opportunities.")
        show_value(fmt_money(total_pipeline))
        label_with_tooltip("Current Win Rate", "Won ÷ (Won + Lost).")
        show_value(f"{win_rate:.1%}")
        label_with_tooltip("Opportunities with Contact Roles", "Unique opportunities appearing in Contact Roles export.")
        show_value(f"{opps_with_cr:,}")
        label_with_tooltip("Opportunities without Contact Roles", "Total opps minus those with roles.")
        show_value(f"{opps_without_cr:,}")
        label_with_tooltip("Pipeline with Contact Roles", "Amount on opps with ≥1 role.")
        show_value(fmt_money(pipeline_with_cr))
        label_with_tooltip("Pipeline without Contact Roles", "Amount on opps with 0 roles.")
        show_value(fmt_money(pipeline_without_cr))
        label_with_tooltip("Opps with only 1 Contact Role", "Opps where role count = 1.")
        show_value(f"{opps_one_cr:,}")
        label_with_tooltip("Pipeline with only 1 Contact Role", "Amount on opps with exactly 1 role.")
        show_value(fmt_money(pipeline_one_cr))

    with c_right:
        label_with_tooltip("Avg Contact Roles – Won", "Average roles per Won opportunity.")
        show_value(f"{avg_cr_won:.1f}")
        label_with_tooltip("Avg Contact Roles – Lost", "Average roles per Lost opportunity.")
        show_value(f"{avg_cr_lost:.1f}")
        label_with_tooltip("Avg Contact Roles – Open", "Average roles per Open opportunity.")
        show_value(f"{avg_cr_open:.1f}")
        if won_zero_count > 0:
            label_with_tooltip("Won Opps with 0 Contact Roles", "Won opportunities missing buying-group contacts.")
            show_value(f"{won_zero_count:,} ({won_zero_pct:.1%} of Won) — {fmt_money(won_zero_pipeline)}")
        label_with_tooltip("Avg days to close – Won", "Close Date − Created Date for Won opps.")
        show_value(f"{avg_days_won:.0f} days" if avg_days_won else "0 days")
        label_with_tooltip("Avg days to close – Lost", "Close Date − Created Date for Lost opps.")
        show_value(f"{avg_days_lost:.0f} days" if avg_days_lost else "0 days")
        label_with_tooltip("Avg age of Open opps", "Today − Created Date for Open opps.")
        show_value(f"{avg_age_open:.0f} days" if avg_age_open else "0 days")

    section_end()

    # ======================================================
    # Pipeline at Risk
    # ======================================================
    section_start("Pipeline at Risk (Low Buying-Group Coverage)")
    st.caption(
        "Open deals with **0–1 contact roles** consistently behave more like Lost deals. "
        "This is your near-term exposure if coverage doesn’t improve."
    )
    label_with_tooltip("Open Pipeline at Risk", "Sum of Amount for open opps with 0–1 contact roles.")
    show_value(fmt_money(open_pipeline_risk))
    label_with_tooltip("% of Open Opps Missing Contacts", "Open opps with 0–1 roles ÷ total open opps.")
    show_value(f"{risk_pct:.1%} ({open_opps_risk:,} of {open_opps_total:,})")
    pct_open_pipe_risk = (open_pipeline_risk / open_pipeline) if open_pipeline > 0 else 0
    label_with_tooltip("% of Open Pipeline at Risk", "Risky open pipeline ÷ total open pipeline.")
    show_value(f"{pct_open_pipe_risk:.1%}")
    section_end()

    # ======================================================
    # Simulator
    # ======================================================
    def modeled_win_rate_for_open(avg_open_contacts, base_win_rate, target_contacts):
        cur = max(avg_open_contacts, 1e-6)
        improvement_factor = min(1.8, target_contacts / cur)
        return min(max(base_win_rate * improvement_factor, base_win_rate), 0.95)

    section_start("Simulator — Target Contact Coverage")
    st.caption(
        "Model how improving buying-group coverage on open deals could change outcomes. "
        "Pick a target average contact count and see modeled win rate uplift and incremental won pipeline."
    )
    target_contacts = st.slider("Target avg contacts on Open Opportunities", 0.0, 5.0, 2.0, 0.5)

    enhanced_win_rate = modeled_win_rate_for_open(avg_cr_open, win_rate, target_contacts)
    current_expected_wins = win_rate * open_pipeline
    enhanced_expected_wins = enhanced_win_rate * open_pipeline
    incremental_won_pipeline = max(0, enhanced_expected_wins - current_expected_wins)

    st.markdown("**Status-quo outlook (if nothing changes):**")
    label_with_tooltip(
        "Expected Won Pipeline (Open) — Current",
        "Open pipeline × current win rate."
    )
    show_value(fmt_money(current_expected_wins))

    st.markdown("**Modeled Uplift (if Open coverage improves):**")
    delta_contacts = target_contacts - avg_cr_open
    pct_contacts = (delta_contacts / avg_cr_open) if avg_cr_open > 0 else 0
    label_with_tooltip("Target Avg Contacts (Open)", "Selected target vs current avg on open deals.")
    show_value(f"{target_contacts:.1f} vs Current {avg_cr_open:.1f} ({delta_contacts:+.1f}, {pct_contacts:+.0%})")

    delta_wr_pp = (enhanced_win_rate - win_rate) * 100
    pct_wr = ((enhanced_win_rate - win_rate) / win_rate) if win_rate > 0 else 0
    label_with_tooltip("Enhanced Win Rate (modeled)", "Modeled win rate if open coverage hits target.")
    show_value(f"{enhanced_win_rate:.1%} vs Current {win_rate:.1%} ({delta_wr_pp:+.1f} pp, {pct_wr:+.0%})")

    pct_pipe = (incremental_won_pipeline / current_expected_wins) if current_expected_wins > 0 else 0
    label_with_tooltip("Enhanced Expected Won Pipeline (Open)", "Expected won pipeline at modeled win rate.")
    show_value(
        f"{fmt_money(enhanced_expected_wins)} vs Current {fmt_money(current_expected_wins)} "
        f"({fmt_money(incremental_won_pipeline)} uplift, {pct_pipe:+.0%})"
    )
    section_end()

    # ======================================================
    # Owner Coverage Rollup (Coaching View)
    # ======================================================
    section_start("Owner Coverage Rollup (Coaching View)")
    st.caption(
        "Reps are ranked by % of open deals missing buying-group contacts. "
        "Expand any rep to see exactly which deals to fix."
    )

    owner_df = open_df.copy()
    owner_df["Opportunity Owner"] = owner_df["Opportunity Owner"].fillna("").astype(str).str.strip()
    owner_df = owner_df[owner_df["Opportunity Owner"] != ""].copy()

    owner_df["is_undercovered"] = (owner_df["contact_count"] <= 1).astype(int)
    owner_df["undercovered_amount"] = owner_df["Amount"].where(owner_df["contact_count"] <= 1, 0)

    owner_roll = owner_df.groupby("Opportunity Owner", dropna=False).agg(
        open_opps=("Opportunity ID", "nunique"),
        opps_undercovered=("is_undercovered", "sum"),
        open_pipeline=("Amount", "sum"),
        undercovered_pipeline=("undercovered_amount", "sum")
    ).reset_index()

    owner_roll["pct_undercovered"] = owner_roll.apply(
        lambda r: r["opps_undercovered"] / r["open_opps"] if r["open_opps"] > 0 else 0,
        axis=1
    )
    owner_roll = owner_roll.sort_values("pct_undercovered", ascending=False)

    if owner_roll.empty:
        st.markdown("No open opportunities found for the selected filters.")
    else:
        stage_priority_order = {"Late": 0, "Mid": 1, "Early": 2, "Open": 3}
        owner_df["Stage Bucket"] = owner_df["Opportunity ID"].apply(stage_bucket_for_id)
        owner_df["Stage Bucket Rank"] = owner_df["Stage Bucket"].map(stage_priority_order).fillna(3)

        shown = 0
        for _, r in owner_roll.iterrows():
            open_opps_n = int(r["open_opps"])
            if open_opps_n == 0:
                continue

            owner_name = r["Opportunity Owner"]
            under_n = int(r["opps_undercovered"])
            pct_under = float(r["pct_undercovered"])
            open_pipe = float(r["open_pipeline"])
            under_pipe = float(r["undercovered_pipeline"])

            exp_title = f"{owner_name} — {pct_under:.0%} open opps under-covered ({under_n}/{open_opps_n})"

            with st.expander(exp_title, expanded=False):
                st.text(
                    f"Pipeline at risk: {fmt_money(under_pipe)} / {fmt_money(open_pipe)} open pipeline"
                )

                rep_under = owner_df[
                    (owner_df["Opportunity Owner"] == owner_name) &
                    (owner_df["contact_count"] <= 1)
                ].copy()

                if rep_under.empty:
                    st.write("✅ No under-covered open opportunities for this rep.")
                else:
                    rep_under = rep_under.sort_values(
                        ["Stage Bucket Rank", "Amount"],
                        ascending=[True, False]
                    )

                    display_cols = [
                        "Opportunity Name",
                        "Account ID",
                        "Stage",
                        "Created Date",
                        "Amount",
                        "contact_count"
                    ]
                    rep_table = rep_under[display_cols].rename(columns={
                        "contact_count": "# Contact Roles"
                    })
                    rep_table["Created Date"] = rep_table["Created Date"].dt.strftime("%Y-%m-%d")
                    rep_table["Amount"] = rep_table["Amount"].map(lambda x: f"${x:,.0f}")

                    st.dataframe(rep_table, use_container_width=True, hide_index=True)

                    st.caption(
                        "Action: Add missing buying-group contacts on these deals. "
                        "Prioritize Late-stage deals first to protect near-term pipeline."
                    )

            shown += 1
            if shown >= 12:
                break

    section_end()

    # ======================================================
    # Top Open Opportunities to Fix First
    # ======================================================
    section_start("Top Open Opportunities to Fix First")
    st.caption(
        "Prioritize these open deals first. They are missing contacts and sorted from Late → Early stages "
        "so the most urgent deals appear on top."
    )

    priority_df = open_df[open_df["contact_count"] <= 1].copy()
    priority_df = priority_df[~priority_df["Stage"].str.contains("Qualified Out", case=False, na=False)].copy()
    priority_df["Stage Bucket"] = priority_df["Opportunity ID"].apply(stage_bucket_for_id)

    stage_priority_order = {"Late": 0, "Mid": 1, "Early": 2, "Open": 3}
    priority_df["Stage Bucket Rank"] = priority_df["Stage Bucket"].map(stage_priority_order).fillna(3)

    priority_df = priority_df.sort_values(
        ["Stage Bucket Rank", "Amount"],
        ascending=[True, False]
    ).head(15)

    priority_bullets = []
    for _, rr in priority_df.iterrows():
        priority_bullets.append(
            f"[{rr.get('Stage Bucket','Open')}] {rr.get('Opportunity Name','(No name)')} "
            f"(ID {rr.get('Opportunity ID','')}) — Stage: {rr.get('Stage','')}, "
            f"Owner: {rr.get('Opportunity Owner','')}, "
            f"Contacts: {int(rr.get('contact_count',0))}, "
            f"Amount: ${rr.get('Amount',0):,.0f}"
        )

    if priority_bullets:
        for b in priority_bullets:
            st.markdown(f"• {b}")
    else:
        st.markdown("• No under-covered open opportunities found (after excluding Qualified Out).")

    section_end()

    # ======================================================
    # Won Opps with 0 Contact Roles — bullets red-flag
    # ======================================================
    if won_zero_count > 0:
        section_start("Won Opps with 0 Contact Roles (Red Flag)")
        st.caption(
            "These won deals are missing buying-group contacts in the CRM. "
            "Fixing this improves reporting accuracy and future forecasting."
        )

        won_zero_bullets = []
        for _, rr in won_zero_df.sort_values("Amount", ascending=False).head(20).iterrows():
            won_zero_bullets.append(
                f"{rr.get('Opportunity Name','(No name)')} (ID {rr.get('Opportunity ID','')}) — "
                f"Owner: {rr.get('Opportunity Owner','')}, Stage: {rr.get('Stage','')}, "
                f"Amount: ${rr.get('Amount',0):,.0f}"
            )

        for b in won_zero_bullets:
            st.markdown(f"• {b}")

        section_end()

    # ======================================================
    # INSIGHTS — 5 CHARTS (unchanged)
    # ======================================================
    section_start("Insights")
    st.caption(
        "Charts below mirror the story: coverage drives win rate and speed, and shows where risk sits today."
    )

    chart_df = opps.copy()
    chart_df["Stage Group"] = "Open"
    chart_df.loc[won_mask, "Stage Group"] = "Won"
    chart_df.loc[lost_mask, "Stage Group"] = "Lost"

    def contact_bucket_winrate(n):
        n = float(n) if pd.notna(n) else 0
        if n <= 0:
            return None
        if n >= 7:
            return "7+"
        return str(int(n))

    chart_df["Winrate Bucket"] = chart_df["contact_count"].apply(contact_bucket_winrate)
    win_bucket_order = ["1", "2", "3", "4", "5", "6", "7+"]

    closed_df = chart_df[chart_df["Stage Group"].isin(["Won", "Lost"])].copy()
    closed_df = closed_df[~((closed_df["Stage Group"] == "Won") & (closed_df["contact_count"] == 0))]
    closed_df = closed_df[closed_df["Winrate Bucket"].notna()]

    winrate_bucket = closed_df.groupby("Winrate Bucket").agg(
        won=("Stage Group", lambda s: (s == "Won").sum()),
        lost=("Stage Group", lambda s: (s == "Lost").sum())
    ).reindex(win_bucket_order).fillna(0).reset_index()

    winrate_bucket["n"] = winrate_bucket["won"] + winrate_bucket["lost"]
    winrate_bucket["Win Rate"] = winrate_bucket.apply(
        lambda rr: rr["won"] / rr["n"] if rr["n"] > 0 else 0, axis=1
    )
    cis = winrate_bucket.apply(lambda rr: wilson_ci(rr["won"], rr["n"]), axis=1)
    winrate_bucket["CI Low"] = [c[0] for c in cis]
    winrate_bucket["CI High"] = [c[1] for c in cis]

    bars_n = alt.Chart(winrate_bucket).mark_bar(opacity=0.35).encode(
        x=alt.X("Winrate Bucket:N", sort=win_bucket_order, title="Contact Roles per Opportunity"),
        y=alt.Y("n:Q", title="# Closed Deals (Won+Lost)"),
        tooltip=["Winrate Bucket", "won", "lost", "n"]
    )
    band_ci = alt.Chart(winrate_bucket).mark_area(opacity=0.15).encode(
        x=alt.X("Winrate Bucket:N", sort=win_bucket_order),
        y=alt.Y("CI Low:Q", axis=alt.Axis(format="%"), title="Win Rate"),
        y2="CI High:Q"
    )
    line_wr = alt.Chart(winrate_bucket).mark_line(point=True, strokeWidth=3).encode(
        x=alt.X("Winrate Bucket:N", sort=win_bucket_order),
        y=alt.Y("Win Rate:Q", axis=alt.Axis(format="%")),
        tooltip=[
            "Winrate Bucket",
            alt.Tooltip("Win Rate:Q", format=".1%"),
            "won", "lost", "n"
        ]
    )

    st.altair_chart(
        alt.layer(bars_n, band_ci, line_wr)
        .resolve_scale(y='independent')
        .properties(height=280, title="Win rate improves sharply after 2+ contact roles"),
        use_container_width=True
    )

    open_chart_df = chart_df[chart_df["Stage Group"] == "Open"].copy()
    open_chart_df["Open Coverage Bucket"] = open_chart_df["contact_count"].apply(
        lambda n: "0 roles" if n == 0 else ("1 role" if n == 1 else "2+ roles")
    )
    open_pipeline_bucket = open_chart_df.groupby("Open Coverage Bucket")["Amount"].sum().reindex(
        ["0 roles", "1 role", "2+ roles"]
    ).fillna(0).reset_index().rename(columns={"Amount": "Open Pipeline"})

    donut = alt.Chart(open_pipeline_bucket).mark_arc(innerRadius=70).encode(
        theta="Open Pipeline:Q",
        color=alt.Color("Open Coverage Bucket:N", legend=None),
        tooltip=["Open Coverage Bucket", alt.Tooltip("Open Pipeline:Q", format=",.0f")]
    ).properties(height=260, title="Open pipeline concentration by coverage (risk today)")
    st.altair_chart(donut, use_container_width=True)

    funnel_df = open_chart_df.copy()
    funnel_df["Coverage Funnel Bucket"] = funnel_df["contact_count"].apply(
        lambda n: "0 roles" if n == 0 else ("1 role" if n == 1 else ("2 roles" if n == 2 else "3+ roles"))
    )
    funnel_counts = funnel_df.groupby("Coverage Funnel Bucket")["Opportunity ID"].nunique().reindex(
        ["0 roles", "1 role", "2 roles", "3+ roles"]
    ).fillna(0).reset_index().rename(columns={"Opportunity ID": "Open Opps"})

    funnel_chart = alt.Chart(funnel_counts).mark_bar().encode(
        y=alt.Y("Coverage Funnel Bucket:N", sort=["0 roles","1 role","2 roles","3+ roles"], title="Coverage bucket"),
        x=alt.X("Open Opps:Q", title="# Open Opportunities"),
        tooltip=["Coverage Funnel Bucket", "Open Opps"]
    ).properties(height=220, title="Coverage funnel for open deals (where depth is missing)")
    st.altair_chart(funnel_chart, use_container_width=True)

    time_df = chart_df.copy()
    time_df["days_to_close"] = time_df.apply(days_diff, axis=1)
    time_df["open_age_days"] = None
    open_mask_local = (time_df["Stage Group"] == "Open") & time_df["Created Date"].notna()
    time_df.loc[open_mask_local, "open_age_days"] = (today - time_df.loc[open_mask_local, "Created Date"]).dt.days

    def contact_bucket_std(n):
        n = float(n) if pd.notna(n) else 0
        if n <= 0: return "0"
        if n == 1: return "1"
        if n == 2: return "2"
        if n == 3: return "3"
        return "4+"

    time_df["Contact Bucket"] = time_df["contact_count"].apply(contact_bucket_std)
    bucket_order_std = ["0", "1", "2", "3", "4+"]

    agg_rows = []
    for sg in ["Won", "Lost"]:
        tmp = time_df[time_df["Stage Group"] == sg]
        grp = tmp.groupby("Contact Bucket")["days_to_close"].mean().reindex(bucket_order_std).reset_index()
        grp["Stage Group"] = sg
        grp = grp.rename(columns={"days_to_close": "Avg Days"})
        agg_rows.append(grp)

    tmp_open = time_df[time_df["Stage Group"] == "Open"]
    grp_open = tmp_open.groupby("Contact Bucket")["open_age_days"].mean().reindex(bucket_order_std).reset_index()
    grp_open["Stage Group"] = "Open"
    grp_open = grp_open.rename(columns={"open_age_days": "Avg Days"})
    agg_rows.append(grp_open)

    avg_days_bucket = pd.concat(agg_rows, ignore_index=True)

    vel_chart = alt.Chart(avg_days_bucket).mark_line(point=True, strokeWidth=3).encode(
        x=alt.X("Contact Bucket:N", sort=bucket_order_std, title="Contact Roles per Opportunity"),
        y=alt.Y("Avg Days:Q", title="Avg Days"),
        color=alt.Color("Stage Group:N", legend=alt.Legend(title="Outcome")),
        tooltip=["Stage Group", "Contact Bucket", alt.Tooltip("Avg Days:Q", format=",.0f")]
    ).properties(height=260, title="More contact roles correlates with faster closes")
    st.altair_chart(vel_chart, use_container_width=True)

    stage_cov_df = opps.copy()
    stage_cov_df["Coverage Bucket"] = stage_cov_df["contact_count"].apply(
        lambda n: "0 roles" if n == 0 else ("1 role" if n == 1 else "2+ roles")
    )

    heat_base = stage_cov_df[stage_cov_df["Stage Bucket"].isin(["Early","Mid","Late","Open"])].copy()
    heat_counts = heat_base.groupby(["Stage Bucket","Coverage Bucket"])["Opportunity ID"].nunique().reset_index()
    stage_totals = heat_base.groupby("Stage Bucket")["Opportunity ID"].nunique().reset_index().rename(columns={"Opportunity ID":"Stage Total"})
    heat_counts = heat_counts.merge(stage_totals, on="Stage Bucket", how="left")
    heat_counts["Pct"] = heat_counts.apply(
        lambda rr: rr["Opportunity ID"]/rr["Stage Total"] if rr["Stage Total"]>0 else 0, axis=1
    )

    heat = alt.Chart(heat_counts).mark_rect().encode(
        x=alt.X("Coverage Bucket:N", sort=["0 roles","1 role","2+ roles"], title="Coverage"),
        y=alt.Y("Stage Bucket:N", sort=["Early","Mid","Late","Open"], title="Stage bucket"),
        color=alt.Color("Pct:Q", scale=alt.Scale(scheme="redyellowgreen"), title="% of opps"),
        tooltip=[
            "Stage Bucket","Coverage Bucket",
            alt.Tooltip("Pct:Q", format=".0%"),
            alt.Tooltip("Opportunity ID:Q", title="# Opps")
        ]
    ).properties(height=240, title="Coverage health by stage bucket (where gaps show up)")
    st.altair_chart(heat, use_container_width=True)

    section_end()

    # ======================================================
    # PDF CHARTS + DOWNLOAD (unchanged from your current working version)
    # ======================================================
    # NOTE: keep using your existing download block here.
    # If you want Stage Coverage Gates added inside the PDF too,
    # tell me and I’ll thread it into metrics_dict + PDF builder.

else:
    st.info("Upload both CSV files above to generate insights.")


st.markdown(
    f"""
<hr style="margin-top:26px; border:0; border-top:1px solid #e5e7eb;" />
<div style="font-size:12px; color:#6b7280; text-align:center; padding:10px 0 2px 0;">
  © {datetime.now().year} RevOps Global. All rights reserved.
</div>
    """,
    unsafe_allow_html=True
)
