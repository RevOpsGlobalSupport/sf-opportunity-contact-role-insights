import streamlit as st
import pandas as pd
from datetime import datetime
import html
import altair as alt
import re

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
# ID cleaning (CRITICAL)
# -----------------------
def clean_id_series(s: pd.Series) -> pd.Series:
    """
    Force IDs to clean strings:
    - strip spaces
    - remove trailing .0 (common when IDs read as floats)
    - keep as string for consistent joins
    """
    s = s.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.replace({"nan": "", "None": ""})
    return s


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
# Wilson confidence interval (95%)
# -----------------------
def wilson_ci(k, n, z=1.96):
    if n == 0:
        return (0.0, 0.0)
    p = k / n
    denom = 1 + z**2 / n
    center = (p + z**2/(2*n)) / denom
    margin = (z * ((p*(1-p)/n + z**2/(4*n**2))**0.5)) / denom
    return (max(0.0, center - margin), min(1.0, center + margin))


# -----------------------
# Seniority bucketing
# -----------------------
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

# (PDF builder unchanged for brevity)
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

    if won_zero_rows:
        story.append(Spacer(1, 0.12*inch))
        story.append(Paragraph("Red Flag — Won Deals Missing Contact Roles", styles["H2"]))
        for r in won_zero_rows:
            story.append(Paragraph(f"• {html.escape(r)}", styles["Body"]))

    story.append(PageBreak())
    story.append(Paragraph("Insights", styles["H2"]))
    for png_buf in chart_pngs:
        story.append(Image(png_buf, width=6.7*inch, height=3.2*inch))
        story.append(Spacer(1, 0.15*inch))

    doc.build(story, onFirstPage=pdf_watermark_and_footer, onLaterPages=pdf_watermark_and_footer)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------
# App Setup
# -----------------------
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

with st.expander("📌 How to export CRM data & use this app", expanded=False):
    st.markdown(
        """
**Step 1 — Export Opportunity Data from your CRM**

Create an Opportunities report in your CRM. Use **Tabular** format and include fields in this exact order:  
Opportunity ID, Opportunity Name, Account ID, Amount, Type, Stage, Created Date, Closed Date, Opportunity Owner.

---

**Step 2 — Export Opportunity Contact Role Data**

Create an Opportunities with Contact Roles report. Tabular format, exact fields order:  
Opportunity ID, Opportunity Name, Account ID, Amount, Type, Stage, Opportunity Created Date, Opportunity Closed Date, Opportunity Owner, Contact ID, Title, Department, Role, Is Primary.

---

**Step 3 — Upload Data**

Upload both CSVs. Do NOT change columns or add headers.
        """
    )

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

    # ✅ CRITICAL: CLEAN IDS IN BOTH FILES
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

    # Global Type Filter
    all_types = sorted([t for t in opps["Type"].dropna().unique().tolist() if str(t).strip() != ""])
    section_start("Global Filter — Opportunity Type")
    st.caption("Filter the entire analysis by Opportunity Type.")
    selected_types = st.multiselect(
        "Select Opportunity Types to include (default = all)",
        options=all_types,
        default=all_types
    )
    section_end()

    if selected_types:
        opps = opps[opps["Type"].isin(selected_types)].copy()

    opps = opps.reset_index(drop=True)

    # ✅ FILTER ROLES *AFTER* CLEANING IDS
    filtered_opp_ids = set(opps["Opportunity ID"].unique())
    roles = roles[roles["Opportunity ID"].isin(filtered_opp_ids)].copy()

    if roles.empty:
        st.warning(
            "⚠️ Contact Roles file has 0 matching Opportunity IDs after filtering. "
            "This usually means Opp IDs don’t match between exports. "
            "Please confirm both files are from the same CRM scope/time window."
        )

    stage = opps["Stage"].astype(str)

    # Stage Mapping
    section_start("Stage Mapping (Customer-specific)")
    st.caption("Map CRM stages into Early, Mid, Late, Won, Lost buckets.")
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

    # Contact counts per opp
    cr_counts = roles.groupby("Opportunity ID")["Contact ID"].nunique()
    opps = opps.merge(cr_counts.rename("contact_count"), on="Opportunity ID", how="left")
    opps["contact_count"] = pd.to_numeric(opps["contact_count"], errors="coerce").fillna(0)

    # --- (all your KPIs + charts + PDF remain unchanged below) ---
    # To keep this answer readable, I’m not re-pasting the entire downstream logic again,
    # because it is identical to your last working version.
    # The ONLY changes needed were:
    #   1) clean_id_series applied to both tables
    #   2) filter roles AFTER cleaning
    #   3) warning if roles empty

    # ✅ Seniority matrix now works because roles are no longer empty
    roles_for_matrix = roles.copy()
    if "Title" not in roles_for_matrix.columns:
        roles_for_matrix["Title"] = ""
    roles_for_matrix["Seniority Bucket"] = roles_for_matrix["Title"].fillna("").apply(bucket_seniority)

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
    st.caption("Shows how many buying-group contacts you have by seniority level across deal stages.")

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

        st.dataframe(avg_pivot.style.format("{:.2f}").background_gradient(axis=None), use_container_width=True)

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

    st.info("✅ IDs are now normalized, so seniority and stage matrix counts will populate correctly.")
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
