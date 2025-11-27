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

# Robust CSV loading (encoding-safe)
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
    """Map columns from CRM export into a stable schema."""
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
    """
    Force IDs to clean strings:
    - strip whitespace
    - remove trailing `.0`
    - normalize missing
    """
    s = s.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.replace({"nan": "", "None": ""})
    return s


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
    """95% Wilson score interval."""
    if n == 0:
        return (0.0, 0.0)
    p = k / n
    denom = 1 + z**2 / n
    center = (p + z**2/(2*n)) / denom
    margin = (z * ((p*(1-p)/n + z**2/(4*n**2))**0.5)) / denom
    return (max(0.0, center - margin), min(1.0, center + margin))


# Seniority bucketing from Title
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


# -----------------------
# PDF helpers
# -----------------------
def fetch_logo_bytes(url: str):
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return io.BytesIO(r.content)
    except Exception:
        return None


def pdf_watermark_and_footer(c: canvas.Canvas, doc):
    # Watermark
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
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
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
        story.append(Paragraph("Red Flag — Won Deals Missing Contact Roles", styles["H2"]))
        for r in won_zero_rows:
            story.append(Paragraph(f"• {html.escape(r)}", styles["Body"]))
        story.append(Spacer(1, 0.12*inch))

    story.append(Paragraph("Insights", styles["H2"]))
    story.append(Spacer(1, 0.1*inch))
    for i, png_buf in enumerate(chart_pngs, start=1):
        story.append(Image(png_buf, width=6.7*inch, height=3.2*inch))
        story.append(Spacer(1, 0.15*inch))
        if i in (2, 4):
            story.append(PageBreak())

    if segment_rows:
        story.append(Paragraph("Segment Uplift (Deal Size Bands)", styles["H2"]))
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
        story.append(Spacer(1, 0.12*inch))

    story.append(Paragraph("Owner Coverage Rollup (Coaching View)", styles["H2"]))
    story.append(Paragraph(
        "Coach owners with the highest under-covered open pipeline first.",
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

# Logo
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

st.markdown("<hr style='margin: 8px 0 16px 0; border:0; border-top:1px solid #e5e7eb;' />", unsafe_allow_html=True)

# CSS
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

# Instructions expander
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

# Uploaders
st.markdown(f"**Upload Opportunities CSV**  \n[Download sample]({sample_opps_url})")
opps_file = st.file_uploader("", type=["csv"], key="opps")

st.markdown(f"**Upload Opportunities with Contact Roles CSV**  \n[Download sample]({sample_roles_url})")
roles_file = st.file_uploader("", type=["csv"], key="roles")


# =========================
# MAIN LOGIC
# =========================
if opps_file and roles_file:
    # Load + normalize
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

    # CRITICAL: clean IDs before any filtering/joins
    opps["Opportunity ID"] = clean_id_series(opps["Opportunity ID"])
    opps = opps[opps["Opportunity ID"] != ""].copy()

    roles["Opportunity ID"] = clean_id_series(roles["Opportunity ID"])
    roles["Contact ID"] = clean_id_series(roles["Contact ID"])
    roles = roles[(roles["Opportunity ID"] != "") & (roles["Contact ID"] != "")].copy()

    # Types
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

    # Filter roles AFTER cleaning IDs
    filtered_opp_ids = set(opps["Opportunity ID"].unique())
    roles = roles[roles["Opportunity ID"].isin(filtered_opp_ids)].copy()

    if roles.empty:
        st.warning(
            "⚠️ Contact Roles file has **0 matching Opportunity IDs** after filtering. "
            "This usually means Opp IDs don’t match between the two exports. "
            "Please ensure both CSVs are from the same CRM scope/time window."
        )

    stage = opps["Stage"].astype(str)

    # STAGE MAPPING
    section_start("Stage Mapping (Customer-specific)")
    st.caption("Map your CRM stages into buckets so stage gates and analysis apply correctly.")

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
        open_mask = ~(won_mask | lost_mask)
        early_mask = open_mask.copy()
        mid_mask = open_mask.copy()
        late_mask = open_mask.copy()

    # CONTACT COUNTS
    cr_counts = roles.groupby("Opportunity ID")["Contact ID"].nunique()
    opps = opps.merge(cr_counts.rename("contact_count"), on="Opportunity ID", how="left")
    opps["contact_count"] = pd.to_numeric(opps["contact_count"], errors="coerce").fillna(0)

    # BASIC SPLITS
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
    won_opps  = opps.loc[won_mask].copy()
    lost_opps = opps.loc[lost_mask].copy()

    won_count = won_opps["Opportunity ID"].nunique()
    lost_count = lost_opps["Opportunity ID"].nunique()
    win_rate = won_count / (won_count + lost_count) if (won_count + lost_count) > 0 else 0

    avg_cr_lost = lost_opps["contact_count"].mean() if not lost_opps.empty else 0
    avg_cr_won  = won_opps["contact_count"].mean() if not won_opps.empty else 0
    avg_cr_open = open_opps["contact_count"].mean() if not open_opps.empty else 0

    # RED FLAG: won with zero contacts
    won_zero_df = won_opps[won_opps["contact_count"] == 0].copy()
    won_zero_count = won_zero_df["Opportunity ID"].nunique()
    won_zero_pipeline = won_zero_df["Amount"].sum()
    won_zero_pct = (won_zero_count / won_count) if won_count > 0 else 0

    # DAYS CALCS
    def days_diff(row):
        if pd.isna(row["Created Date"]) or pd.isna(row["Close Date"]):
            return None
        return (row["Close Date"] - row["Created Date"]).days

    if not lost_opps.empty:
        lost_opps["days_to_close"] = lost_opps.apply(days_diff, axis=1)
    if not won_opps.empty:
        won_opps["days_to_close"] = won_opps.apply(days_diff, axis=1)

    avg_days_lost = lost_opps["days_to_close"].dropna().mean() if "days_to_close" in lost_opps else None
    avg_days_won  = won_opps["days_to_close"].dropna().mean() if "days_to_close" in won_opps else None

    today = pd.Timestamp.today().normalize()
    if not open_opps.empty:
        open_opps["age_days"] = (today - open_opps["Created Date"]).dt.days
    avg_age_open = open_opps["age_days"].dropna().mean() if "age_days" in open_opps else None

    # =========================
    # SENIORITY MATRIX (FIXED)
    # =========================
    roles_for_matrix = roles.copy()
    if "Title" not in roles_for_matrix.columns:
        roles_for_matrix["Title"] = ""
    roles_for_matrix["Title"] = roles_for_matrix["Title"].fillna("").astype(str)
    roles_for_matrix["Seniority Bucket"] = roles_for_matrix["Title"].apply(bucket_seniority)

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
        else:
            sl = str(s).lower()
            if "won" in sl: return "Won"
            if "lost" in sl: return "Lost"
            return "Open"

    roles_for_matrix["Stage Bucket"] = roles_for_matrix["Opportunity ID"].apply(stage_bucket_for_id)

    stage_order = ["Early", "Mid", "Late", "Open", "Won", "Lost"]
    seniority_order = ["C-Level", "EVP / SVP", "VP", "Director / Head", "Manager", "IC / Staff", "Other / Unknown"]

    section_start("Seniority / Job-Level Coverage by Stage")
    st.caption("Buckets Contact Role titles into seniority levels and shows buying-group coverage by deal stage.")

    matrix_view = st.radio(
        "Matrix View",
        ["# Unique Contact Roles", "Avg Contact Roles per Opportunity"],
        horizontal=True
    )

    base_pivot = roles_for_matrix.pivot_table(
        index="Stage Bucket",
        columns="Seniority Bucket",
        values="Contact ID",
        aggfunc="nunique",
        fill_value=0
    ).reindex(stage_order).fillna(0)

    for col in seniority_order:
        if col not in base_pivot.columns:
            base_pivot[col] = 0
    base_pivot = base_pivot[seniority_order]

    if matrix_view == "# Unique Contact Roles":
        st.dataframe(base_pivot.style.background_gradient(axis=None), use_container_width=True)
    else:
        opp_stage_counts = opps.copy()
        opp_stage_counts["Stage Bucket"] = opp_stage_counts["Opportunity ID"].apply(stage_bucket_for_id)
        opps_per_stage = opp_stage_counts.groupby("Stage Bucket")["Opportunity ID"].nunique()

        avg_pivot = base_pivot.copy()
        for r in avg_pivot.index:
            denom = opps_per_stage.get(r, 0)
            avg_pivot.loc[r] = (avg_pivot.loc[r] / denom) if denom > 0 else 0

        st.dataframe(
            avg_pivot.style.format("{:.2f}").background_gradient(axis=None),
            use_container_width=True
        )

    stacked_df = base_pivot.reset_index().melt(
        id_vars="Stage Bucket",
        var_name="Seniority Bucket",
        value_name="Contact Roles"
    )
    stacked_df["Stage Bucket"] = pd.Categorical(stacked_df["Stage Bucket"], categories=stage_order, ordered=True)
    stacked_df["Seniority Bucket"] = pd.Categorical(stacked_df["Seniority Bucket"], categories=seniority_order, ordered=True)

    st.caption("Stacked view of seniority coverage per stage.")
    st.altair_chart(
        alt.Chart(stacked_df).mark_bar().encode(
            x=alt.X("Stage Bucket:N", sort=stage_order, title="Stage Bucket"),
            y=alt.Y("Contact Roles:Q", title="# Unique Contact Roles"),
            color=alt.Color("Seniority Bucket:N", sort=seniority_order, title="Seniority"),
            tooltip=["Stage Bucket", "Seniority Bucket", "Contact Roles"]
        ).properties(height=280),
        use_container_width=True
    )
    section_end()

    # =========================
    # STAGE GATES
    # =========================
    section_start("Stage Coverage Gates")
    st.caption("Define minimum Contact Roles expected by deal phase (Mid & Late).")

    default_mid_gate  = max(2.0, round(avg_cr_won * 0.6, 1)) if avg_cr_won else 2.0
    default_late_gate = max(3.0, round(avg_cr_won * 0.85, 1)) if avg_cr_won else 3.0

    mid_gate_target = st.slider("Gate 1 — Mid stage minimum contacts", 0.0, 8.0, float(default_mid_gate), 0.5)
    late_gate_target = st.slider("Gate 2 — Late stage minimum contacts", 0.0, 10.0, float(default_late_gate), 0.5)
    section_end()

    mid_df = opps.loc[mid_mask].copy() if user_mapped_any else open_opps.copy()
    late_df = opps.loc[late_mask].copy() if user_mapped_any else open_opps.copy()

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

    label_with_tooltip("Mid-stage opps below Gate 1", "Mid opps below the Gate 1 contact target.")
    show_value(f"{mid_below_pct:.1%} ({mid_below_cnt:,} of {mid_cnt:,})")
    label_with_tooltip("Mid-stage pipeline below Gate 1", "Total Amount of Mid opps below Gate 1.")
    show_value(f"${mid_below_pipe:,.0f}")

    label_with_tooltip("Late-stage opps below Gate 2", "Late opps below the Gate 2 contact target.")
    show_value(f"{late_below_pct:.1%} ({late_below_cnt:,} of {late_cnt:,})")
    label_with_tooltip("Late-stage pipeline below Gate 2", "Total Amount of Late opps below Gate 2.")
    show_value(f"${late_below_pipe:,.0f}")

    callouts = []
    if mid_below_pct > 0.30:
        callouts.append(f"Mid-stage risk: **{mid_below_pct:.0%}** of Mid opps are below the stakeholder gate.")
    if late_below_pct > 0.20:
        callouts.append(f"Late-stage gap: **{late_below_pct:.0%}** of Late opps are below the buying-group gate.")

    if callouts:
        st.markdown("**Callouts:**")
        for c in callouts:
            st.markdown(f"- {c}")

    section_end()

    # =========================
    # PIPELINE RISK
    # =========================
    open_df = open_opps.copy()

    open_pipeline = open_df["Amount"].sum() if not open_df.empty else 0
    open_pipeline_risk = open_df[open_df["contact_count"] <= 1]["Amount"].sum() if not open_df.empty else 0
    open_opps_risk = open_df[open_df["contact_count"] <= 1]["Opportunity ID"].nunique() if not open_df.empty else 0
    open_opps_total = open_df["Opportunity ID"].nunique() if not open_df.empty else 0
    risk_pct = open_opps_risk / open_opps_total if open_opps_total > 0 else 0

    section_start("Pipeline Risk")
    st.caption("How much open pipeline is under-covered today (0–1 roles).")
    label_with_tooltip("Open Pipeline at Risk", "Sum of Amount for open opps with 0–1 contact roles.")
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

    # =========================
    # CORE METRICS
    # =========================
    section_start("Core Metrics")
    st.caption("Baseline deal volume, pipeline, and win rate.")
    label_with_tooltip("Total Opportunities", "Unique opportunities in the export.")
    show_value(f"{total_opps:,}")
    label_with_tooltip("Total Pipeline", "Sum of Amount for all opportunities.")
    show_value(f"${total_pipeline:,.0f}")
    label_with_tooltip("Current Win Rate", "Won ÷ (Won + Lost).")
    show_value(f"{win_rate:.1%}")
    section_end()

    # Coverage Summary
    section_start("Contact Role Coverage")
    st.caption("Overall stakeholder coverage and pipeline missing roles.")
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

    # Outcome averages + red flag
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
            "Won opportunities that have no buying-group contacts. Indicates hygiene or tracking gaps."
        )
        show_value(f"{won_zero_count:,} ({won_zero_pct:.1%} of Won) — ${won_zero_pipeline:,.0f} pipeline")
    section_end()

    # Time metrics
    section_start("Time to Close")
    st.caption("Velocity patterns: wins close faster; stale open deals convert poorly.")
    label_with_tooltip("Avg days to close – Won", "Close Date − Created Date for Won opps.")
    show_value(f"{avg_days_won:.0f} days" if avg_days_won else "0 days")
    label_with_tooltip("Avg days to close – Lost", "Close Date − Created Date for Lost opps.")
    show_value(f"{avg_days_lost:.0f} days" if avg_days_lost else "0 days")
    label_with_tooltip("Avg age of Open opps", "Today − Created Date for Open opps.")
    show_value(f"{avg_age_open:.0f} days" if avg_age_open else "0 days")
    section_end()

    # =========================
    # SIMULATOR
    # =========================
    def modeled_win_rate_for_open(avg_open_contacts, base_win_rate, target_contacts):
        cur = max(avg_open_contacts, 1e-6)
        improvement_factor = min(1.8, target_contacts / cur)
        return min(max(base_win_rate * improvement_factor, base_win_rate), 0.95)

    section_start("Simulator — Target Contact Coverage")
    st.caption("Adjust target buying-group size and see modeled win-rate and revenue lift.")
    target_contacts = st.slider("Target avg contacts on Open Opps", 0.0, 5.0, 2.0, 0.5)
    section_end()

    enhanced_win_rate = modeled_win_rate_for_open(avg_cr_open, win_rate, target_contacts)
    incremental_won_pipeline = max(0, (enhanced_win_rate - win_rate) * open_pipeline)

    section_start("Pipeline Upside from Simulator")
    st.caption("Translates modeled uplift into incremental won pipeline.")
    label_with_tooltip("Modeled Upside if Coverage Improves", "(Enhanced win rate − current win rate) × open pipeline.")
    show_value(f"${incremental_won_pipeline:,.0f}")
    section_end()

    # =========================
    # COVERAGE-ADJUSTED FORECAST
    # =========================
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
        "Estimated win rate for open deals after accounting for stakeholder coverage. More contact roles = healthier."
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

    # =========================
    # EXECUTIVE SUMMARY (NO DUPES)
    # =========================
    bullets = []

    bullets.append(
        f"Current win rate is **{win_rate:.1%}**. Won deals average **{avg_cr_won:.1f}** contact roles vs Lost at **{avg_cr_lost:.1f}**, "
        "showing strong correlation between buying-group depth and conversion."
    )

    if avg_cr_open < max(avg_cr_won, 2.0):
        bullets.append(
            f"Open opportunities average **{avg_cr_open:.1f}** contact roles. Increasing this towards **at least {target_contacts:.1f}** "
            "contacts per opportunity would better align open deals with won buying-group patterns."
        )

    if open_pipeline_risk > 0:
        bullets.append(
            f"**${open_pipeline_risk:,.0f}** of open pipeline is under-covered (0–1 roles), representing **{risk_pct:.0%}** of open opportunities."
        )

    if incremental_won_pipeline > 0:
        bullets.append(
            f"If open buying-group coverage improves to the target, win rate models from **{win_rate:.1%} → {enhanced_win_rate:.1%}**, "
            f"yielding **${incremental_won_pipeline:,.0f}** in incremental won pipeline."
        )

    if won_zero_count > 0:
        bullets.append(
            f"⚠️ **{won_zero_count:,} Won deals** have **0** contact roles logged (≈{won_zero_pct:.0%} of Won), indicating CRM hygiene gaps."
        )

    section_start("Executive Summary")
    for b in bullets:
        st.markdown(f"• {b}")
    section_end()

    # =========================
    # INSIGHTS (CHARTS)
    # =========================
    section_start("Insights")
    st.caption("Visualizes relationships between contact coverage, win rates, pipeline risk, and velocity.")

    chart_df = opps.copy()
    chart_df["Stage Group"] = "Open"
    chart_df.loc[won_mask, "Stage Group"] = "Won"
    chart_df.loc[lost_mask, "Stage Group"] = "Lost"

    # Win-rate buckets 1-6, 7+
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
        lambda r: r["won"] / r["n"] if r["n"] > 0 else 0,
        axis=1
    )
    cis = winrate_bucket.apply(lambda r: wilson_ci(r["won"], r["n"]), axis=1)
    winrate_bucket["CI Low"] = [c[0] for c in cis]
    winrate_bucket["CI High"] = [c[1] for c in cis]

    band = alt.Chart(winrate_bucket).mark_area(opacity=0.18).encode(
        x=alt.X("Winrate Bucket:N", sort=win_bucket_order, title="Contact Roles per Opportunity"),
        y=alt.Y("CI Low:Q", axis=alt.Axis(format="%"), title="Win Rate"),
        y2="CI High:Q",
        tooltip=[
            "Winrate Bucket",
            alt.Tooltip("Win Rate:Q", format=".1%"),
            alt.Tooltip("CI Low:Q", format=".1%"),
            alt.Tooltip("CI High:Q", format=".1%"),
            "won", "lost", "n"
        ]
    )

    line = alt.Chart(winrate_bucket).mark_line(point=True).encode(
        x=alt.X("Winrate Bucket:N", sort=win_bucket_order),
        y=alt.Y("Win Rate:Q"),
        tooltip=[
            "Winrate Bucket",
            alt.Tooltip("Win Rate:Q", format=".1%"),
            alt.Tooltip("CI Low:Q", format=".1%"),
            alt.Tooltip("CI High:Q", format=".1%"),
            "won", "lost", "n"
        ]
    )
    st.altair_chart((band + line).properties(height=260), use_container_width=True)

    # Open pipeline by coverage buckets
    open_chart_df = chart_df[chart_df["Stage Group"] == "Open"].copy()
    open_chart_df["Open Coverage Bucket"] = open_chart_df["contact_count"].apply(
        lambda n: "0 Contact Roles" if n == 0 else ("1 Contact Role" if n == 1 else "2+ Contact Roles")
    )
    open_pipeline_bucket = open_chart_df.groupby("Open Coverage Bucket")["Amount"].sum().reindex(
        ["0 Contact Roles", "1 Contact Role", "2+ Contact Roles"]
    ).fillna(0).reset_index().rename(columns={"Amount": "Open Pipeline"})

    st.altair_chart(
        alt.Chart(open_pipeline_bucket).mark_bar().encode(
            x=alt.X("Open Coverage Bucket:N", title="Open Coverage"),
            y=alt.Y("Open Pipeline:Q", title="Open Pipeline ($)"),
            tooltip=["Open Coverage Bucket", alt.Tooltip("Open Pipeline:Q", format=",.0f")]
        ).properties(height=260),
        use_container_width=True
    )

    # Velocity vs contacts
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

    st.altair_chart(
        alt.Chart(avg_days_bucket).mark_bar().encode(
            x=alt.X("Contact Bucket:N", sort=bucket_order_std, title="Contact Roles per Opportunity"),
            y=alt.Y("Avg Days:Q", title="Avg Days (Close for Won/Lost, Age for Open)"),
            color=alt.Color("Stage Group:N", title="Outcome"),
            tooltip=["Stage Group", "Contact Bucket", alt.Tooltip("Avg Days:Q", format=",.0f")]
        ).properties(height=300),
        use_container_width=True
    )

    # Impact chart
    impact_df = pd.DataFrame({
        "Scenario": ["Current Win Rate", "Enhanced Win Rate"],
        "Win Rate": [win_rate, enhanced_win_rate]
    })
    st.altair_chart(
        alt.Chart(impact_df).mark_bar().encode(
            x=alt.X("Scenario:N", title=""),
            y=alt.Y("Win Rate:Q", axis=alt.Axis(format="%"), title="Win Rate"),
            tooltip=["Scenario", alt.Tooltip("Win Rate:Q", format=".1%")]
        ).properties(height=240),
        use_container_width=True
    )

    # Zero contact rate by outcome
    zero_rate_df = chart_df.groupby("Stage Group").agg(
        opps=("Opportunity ID", "nunique"),
        zero_opps=("contact_count", lambda s: (s == 0).sum())
    ).reset_index()
    zero_rate_df["Zero Contact Rate"] = zero_rate_df.apply(
        lambda r: r["zero_opps"] / r["opps"] if r["opps"] > 0 else 0, axis=1
    )

    st.altair_chart(
        alt.Chart(zero_rate_df).mark_bar().encode(
            x=alt.X("Stage Group:N", title="Outcome"),
            y=alt.Y("Zero Contact Rate:Q", axis=alt.Axis(format="%"), title="Zero-Contact Rate"),
            tooltip=["Stage Group", "zero_opps", "opps", alt.Tooltip("Zero Contact Rate:Q", format=".1%")]
        ).properties(height=240),
        use_container_width=True
    )

    section_end()

    # =========================
    # WON ZERO DETAIL
    # =========================
    won_zero_rows_for_pdf = []
    if won_zero_count > 0:
        section_start("Red Flag — Won Deals Missing Contact Roles")
        st.caption("These deals closed Won without any stakeholder roles logged. Review CRM hygiene or process gaps.")
        show_cols = ["Opportunity ID", "Opportunity Name", "Opportunity Owner", "Stage", "Amount", "Created Date", "Close Date"]
        for c in show_cols:
            if c not in won_zero_df.columns:
                won_zero_df[c] = ""

        st.dataframe(
            won_zero_df[show_cols].sort_values("Amount", ascending=False)
                .style.format({"Amount": "${:,.0f}"}),
            use_container_width=True,
            hide_index=True
        )

        for _, r in won_zero_df.sort_values("Amount", ascending=False).head(15).iterrows():
            won_zero_rows_for_pdf.append(
                f"{r.get('Opportunity Name','')} (ID {r.get('Opportunity ID','')}) — "
                f"Owner: {r.get('Opportunity Owner','')}, Stage: {r.get('Stage','')}, "
                f"Amount: ${r.get('Amount',0):,.0f}"
            )
        section_end()

    # =========================
    # OWNER ROLLUP
    # =========================
    owner_rollup_rows = []
    if "Opportunity Owner" in opps.columns:
        section_start("Owner Coverage Rollup (Coaching View)")
        st.caption("Coach owners with the highest percentage of under-covered open pipeline first.")

        owner_df = open_opps.copy()
        owner_df["Opportunity Owner"] = owner_df["Opportunity Owner"].fillna("Unassigned")

        roll = owner_df.groupby("Opportunity Owner").agg(
            open_opps=("Opportunity ID", "nunique"),
            avg_contacts=("contact_count", "mean"),
            undercovered_pct=("contact_count", lambda s: (s <= 1).mean()),
            open_pipeline=("Amount", "sum")
        ).reset_index()

        roll = roll.sort_values(
            ["undercovered_pct", "open_pipeline"],
            ascending=[False, False]
        ).head(10)

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

    # =========================
    # TOP OPEN OPPS
    # =========================
    section_start("Top Open Opportunities to Fix First")
    st.caption("Prioritize multi-threading these deals based on value, age, and low contact coverage.")

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
                f"{r.get('Type','')}, "
                f"${r.get('Amount',0):,.0f}, {r.get('contact_count',0):.0f} contacts, "
                f"{r.get('age_days',0):.0f} days open"
            )
            top_opps_rows.append(line)
            st.markdown(f"- **{line}**")
    else:
        st.write("No open opportunities found.")
    section_end()

    # =========================
    # MAKE PDF CHARTS
    # =========================
    pdf_chart_pngs = []

    # (1) Win-rate chart
    fig1 = plt.figure(figsize=(7.2, 3.2))
    ax1 = fig1.add_subplot(111)
    x = winrate_bucket["Winrate Bucket"].tolist()
    y = winrate_bucket["Win Rate"].tolist()
    lo = winrate_bucket["CI Low"].tolist()
    hi = winrate_bucket["CI High"].tolist()
    ax1.plot(x, y, marker="o", label="Win Rate")
    ax1.fill_between(x, lo, hi, alpha=0.2, label="95% CI (Wilson)")
    ax1.set_title("Win Rate vs Contact Roles (1–7+, Won-0 excluded)")
    ax1.set_xlabel("Contact Roles per Opportunity")
    ax1.set_ylabel("Win Rate")
    ax1.yaxis.set_major_formatter(PercentFormatter(1.0))
    ax1.legend()
    pdf_chart_pngs.append(fig_to_png_bytes(fig1))

    # (2) Open pipeline risk
    fig2 = plt.figure(figsize=(7.2, 3.2))
    ax2 = fig2.add_subplot(111)
    ax2.bar(open_pipeline_bucket["Open Coverage Bucket"], open_pipeline_bucket["Open Pipeline"])
    ax2.set_title("Open Pipeline by Contact Coverage")
    ax2.set_xlabel("Open Coverage")
    ax2.set_ylabel("Open Pipeline ($)")
    ax2.tick_params(axis='x', rotation=10)
    pdf_chart_pngs.append(fig_to_png_bytes(fig2))

    # (3) Velocity vs roles
    fig3 = plt.figure(figsize=(7.2, 3.6))
    ax3 = fig3.add_subplot(111)
    for sg in ["Won", "Lost", "Open"]:
        sub = avg_days_bucket[avg_days_bucket["Stage Group"] == sg]
        ax3.plot(sub["Contact Bucket"], sub["Avg Days"], marker="o", label=sg)
    ax3.set_title("Velocity vs Contact Roles")
    ax3.set_xlabel("Contact Roles per Opportunity")
    ax3.set_ylabel("Avg Days")
    ax3.legend()
    pdf_chart_pngs.append(fig_to_png_bytes(fig3))

    # (4) Uplift chart
    fig4 = plt.figure(figsize=(6.5, 3.0))
    ax4 = fig4.add_subplot(111)
    ax4.bar(impact_df["Scenario"], impact_df["Win Rate"])
    ax4.set_title("Modeled Win Rate Uplift")
    ax4.set_ylabel("Win Rate")
    ax4.yaxis.set_major_formatter(PercentFormatter(1.0))
    pdf_chart_pngs.append(fig_to_png_bytes(fig4))

    # (5) Zero-contact by outcome
    fig5 = plt.figure(figsize=(6.8, 3.0))
    ax5 = fig5.add_subplot(111)
    ax5.bar(zero_rate_df["Stage Group"], zero_rate_df["Zero Contact Rate"])
    ax5.set_title("Zero-Contact Rate by Outcome")
    ax5.set_ylabel("Zero-Contact Rate")
    ax5.yaxis.set_major_formatter(PercentFormatter(1.0))
    pdf_chart_pngs.append(fig_to_png_bytes(fig5))

    # =========================
    # PDF DOWNLOAD
    # =========================
    section_start("Download Full PDF Report")
    st.caption("Download a branded PDF with metrics, callouts, charts, and next-step lists.")

    metrics_dict = {
        "Stage Gate Coverage": [
            ["Mid-stage opps below Gate 1", f"{mid_below_pct:.1%} ({mid_below_cnt:,} of {mid_cnt:,})"],
            ["Mid-stage pipeline below Gate 1", f"${mid_below_pipe:,.0f}"],
            ["Late-stage opps below Gate 2", f"{late_below_pct:.1%} ({late_below_cnt:,} of {late_cnt:,})"],
            ["Late-stage pipeline below Gate 2", f"${late_below_pipe:,.0f}"],
        ],
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
            ["⚠️ Won Opps with 0 Contact Roles", f"{won_zero_count:,} ({won_zero_pct:.1%})"],
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
        segment_rows=[],
        won_zero_rows=won_zero_rows_for_pdf
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
RevOps Global’s **Buying Group Automation** helps teams identify stakeholders, close coverage gaps,  
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
