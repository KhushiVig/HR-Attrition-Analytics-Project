"""
HR Employee Attrition – Data Analytics Project
================================================
Run this script to produce:
  • charts/   – PNG charts (corporate blue/grey palette)
  • HR_Attrition_Report.docx – auto-generated Word business report

Usage:
    python analysis.py
"""

import os
import warnings
import textwrap

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")          # headless – no GUI window needed
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

warnings.filterwarnings("ignore")

# ── Palette ────────────────────────────────────────────────────────────────────
BLUE      = "#1F4E79"
MID_BLUE  = "#2E75B6"
LIGHT_BLUE= "#9DC3E6"
GREY      = "#595959"
LIGHT_GREY= "#D9D9D9"
ACCENT    = "#ED7D31"          # single warm accent for "risk" highlights

PALETTE   = [MID_BLUE, LIGHT_BLUE, GREY, LIGHT_GREY, BLUE, ACCENT]

sns.set_theme(style="whitegrid", palette=PALETTE)
plt.rcParams.update({
    "figure.dpi": 150,
    "axes.titlesize": 12,
    "axes.titleweight": "bold",
    "axes.labelsize": 10,
    "xtick.labelsize": 9,
    "ytick.labelsize": 9,
    "axes.spines.top": False,
    "axes.spines.right": False,
})

CHARTS_DIR = "charts"
os.makedirs(CHARTS_DIR, exist_ok=True)

# ── 1. Data Loading & Cleaning ─────────────────────────────────────────────────
print("Loading dataset ...")
df = pd.read_csv("WA_Fn-UseC_-HR-Employee-Attrition.csv")

# Encode target
df["AttritionFlag"] = (df["Attrition"] == "Yes").astype(int)

# Age groups
df["AgeGroup"] = pd.cut(
    df["Age"],
    bins=[17, 25, 35, 45, 55, 100],
    labels=["18–25", "26–35", "36–45", "46–55", "55+"]
)

# Income bands
df["IncomeBand"] = pd.cut(
    df["MonthlyIncome"],
    bins=[0, 3000, 6000, 10000, 100000],
    labels=["<$3k", "$3k–$6k", "$6k–$10k", ">$10k"]
)

print(f"  Rows: {len(df):,}  |  Columns: {df.shape[1]}")

# ── 2. Executive KPIs ──────────────────────────────────────────────────────────
attrition_rate   = df["AttritionFlag"].mean() * 100
avg_income       = df["MonthlyIncome"].mean()
avg_tenure       = df["YearsAtCompany"].mean()
overtime_rate    = (df["OverTime"] == "Yes").mean() * 100
avg_job_sat      = df["JobSatisfaction"].mean()

kpis = {
    "Overall Attrition Rate":   f"{attrition_rate:.1f}%",
    "Avg Monthly Income":       f"${avg_income:,.0f}",
    "Avg Tenure (Years)":       f"{avg_tenure:.1f}",
    "Overtime Rate":            f"{overtime_rate:.1f}%",
    "Avg Job Satisfaction":     f"{avg_job_sat:.2f} / 4",
}

print("\n-- Executive KPIs --------------------------------------------------")
for k, v in kpis.items():
    print(f"  {k:<30} {v}")

# ── helper: save figure ────────────────────────────────────────────────────────
def savefig(name):
    path = os.path.join(CHARTS_DIR, name)
    plt.tight_layout()
    plt.savefig(path, bbox_inches="tight")
    plt.close()
    print(f"  Saved -> {path}")
    return path

# ── helper: attrition rate by group ───────────────────────────────────────────
def attrition_by(col):
    return (
        df.groupby(col)["AttritionFlag"]
        .agg(["mean", "count"])
        .rename(columns={"mean": "AttritionRate", "count": "Headcount"})
        .assign(AttritionRate=lambda x: x["AttritionRate"] * 100)
        .sort_values("AttritionRate", ascending=False)
    )

# ── 3. Attrition Trends ────────────────────────────────────────────────────────
print("\nGenerating attrition trend charts ...")

# 3a. By Department
fig, ax = plt.subplots(figsize=(7, 4))
dept = attrition_by("Department")
bars = ax.barh(dept.index, dept["AttritionRate"], color=MID_BLUE)
ax.bar_label(bars, fmt="%.1f%%", padding=4, fontsize=9)
ax.axvline(attrition_rate, color=ACCENT, linestyle="--", linewidth=1.4, label=f"Avg {attrition_rate:.1f}%")
ax.set_xlabel("Attrition Rate (%)")
ax.set_title("Attrition Rate by Department")
ax.legend(fontsize=8)
savefig("01_attrition_by_department.png")

# 3b. By Job Role
fig, ax = plt.subplots(figsize=(8, 5))
role = attrition_by("JobRole")
sales_rep_rate = role.loc["Sales Representative", "AttritionRate"]
bars = ax.barh(role.index, role["AttritionRate"],
               color=[ACCENT if v > attrition_rate * 1.5 else MID_BLUE for v in role["AttritionRate"]])
ax.bar_label(bars, fmt="%.1f%%", padding=4, fontsize=9)
ax.axvline(attrition_rate, color=GREY, linestyle="--", linewidth=1.4, label=f"Avg {attrition_rate:.1f}%")
ax.set_xlabel("Attrition Rate (%)")
ax.set_title("Attrition Rate by Job Role")
ax.legend(fontsize=8)
savefig("02_attrition_by_jobrole.png")

# 3c. By Age Group
fig, ax = plt.subplots(figsize=(7, 4))
age = attrition_by("AgeGroup")
bars = ax.bar(age.index.astype(str), age["AttritionRate"], color=MID_BLUE, width=0.5)
ax.bar_label(bars, fmt="%.1f%%", padding=3, fontsize=9)
ax.axhline(attrition_rate, color=ACCENT, linestyle="--", linewidth=1.4, label=f"Avg {attrition_rate:.1f}%")
ax.set_ylabel("Attrition Rate (%)")
ax.set_title("Attrition Rate by Age Group")
ax.legend(fontsize=8)
savefig("03_attrition_by_agegroup.png")

# 3d. By Income Band
fig, ax = plt.subplots(figsize=(7, 4))
inc = attrition_by("IncomeBand")
inc = inc.reindex(["<$3k", "$3k–$6k", "$6k–$10k", ">$10k"])
income_low_rate  = inc.loc["<$3k",  "AttritionRate"]
income_high_rate = inc.loc[">$10k", "AttritionRate"]
income_ratio     = income_low_rate / income_high_rate
bars = ax.bar(inc.index.astype(str), inc["AttritionRate"], color=MID_BLUE, width=0.5)
ax.bar_label(bars, fmt="%.1f%%", padding=3, fontsize=9)
ax.axhline(attrition_rate, color=ACCENT, linestyle="--", linewidth=1.4, label=f"Avg {attrition_rate:.1f}%")
ax.set_ylabel("Attrition Rate (%)")
ax.set_title("Attrition Rate by Income Band")
ax.legend(fontsize=8)
savefig("04_attrition_by_incomeband.png")

# ── 4. Drivers of Attrition ────────────────────────────────────────────────────
print("\nGenerating driver charts ...")

# 4a. OverTime
fig, ax = plt.subplots(figsize=(5, 4))
ot = attrition_by("OverTime")
colors = [ACCENT if i == "Yes" else MID_BLUE for i in ot.index]
bars = ax.bar(ot.index, ot["AttritionRate"], color=colors, width=0.4)
ax.bar_label(bars, fmt="%.1f%%", padding=3, fontsize=10)
ax.axhline(attrition_rate, color=GREY, linestyle="--", linewidth=1.2, label=f"Avg {attrition_rate:.1f}%")
ax.set_ylabel("Attrition Rate (%)")
ax.set_title("Attrition Rate: OverTime")
ax.legend(fontsize=8)
savefig("05_driver_overtime.png")

# 4b. Job Satisfaction
fig, ax = plt.subplots(figsize=(6, 4))
js = attrition_by("JobSatisfaction").reindex([1, 2, 3, 4])
bars = ax.bar(js.index.astype(str), js["AttritionRate"], color=MID_BLUE, width=0.5)
ax.bar_label(bars, fmt="%.1f%%", padding=3, fontsize=9)
ax.axhline(attrition_rate, color=ACCENT, linestyle="--", linewidth=1.2, label=f"Avg {attrition_rate:.1f}%")
ax.set_xlabel("Job Satisfaction (1=Low … 4=High)")
ax.set_ylabel("Attrition Rate (%)")
ax.set_title("Attrition Rate by Job Satisfaction")
ax.legend(fontsize=8)
savefig("06_driver_jobsatisfaction.png")

# 4c. Work-Life Balance
fig, ax = plt.subplots(figsize=(6, 4))
wlb = attrition_by("WorkLifeBalance").reindex([1, 2, 3, 4])
bars = ax.bar(wlb.index.astype(str), wlb["AttritionRate"], color=MID_BLUE, width=0.5)
ax.bar_label(bars, fmt="%.1f%%", padding=3, fontsize=9)
ax.axhline(attrition_rate, color=ACCENT, linestyle="--", linewidth=1.2, label=f"Avg {attrition_rate:.1f}%")
ax.set_xlabel("Work-Life Balance (1=Bad … 4=Best)")
ax.set_ylabel("Attrition Rate (%)")
ax.set_title("Attrition Rate by Work-Life Balance")
ax.legend(fontsize=8)
savefig("07_driver_worklifebalance.png")

# 4d. Years Since Last Promotion (binned)
df["PromoBand"] = pd.cut(
    df["YearsSinceLastPromotion"],
    bins=[-1, 1, 3, 6, 15],
    labels=["0–1 yr", "2–3 yr", "4–6 yr", "7+ yr"]
)
fig, ax = plt.subplots(figsize=(6, 4))
promo = attrition_by("PromoBand").reindex(["0–1 yr", "2–3 yr", "4–6 yr", "7+ yr"])
bars = ax.bar(promo.index.astype(str), promo["AttritionRate"], color=MID_BLUE, width=0.5)
ax.bar_label(bars, fmt="%.1f%%", padding=3, fontsize=9)
ax.axhline(attrition_rate, color=ACCENT, linestyle="--", linewidth=1.2, label=f"Avg {attrition_rate:.1f}%")
ax.set_xlabel("Years Since Last Promotion")
ax.set_ylabel("Attrition Rate (%)")
ax.set_title("Attrition Rate by Time Since Last Promotion")
ax.legend(fontsize=8)
savefig("08_driver_promotion.png")

# 4e. Distance From Home (binned)
df["DistBand"] = pd.cut(
    df["DistanceFromHome"],
    bins=[0, 5, 15, 30],
    labels=["Near (≤5)", "Mid (6–15)", "Far (16–30)"]
)
fig, ax = plt.subplots(figsize=(6, 4))
dist = attrition_by("DistBand").reindex(["Near (≤5)", "Mid (6–15)", "Far (16–30)"])
bars = ax.bar(dist.index.astype(str), dist["AttritionRate"], color=MID_BLUE, width=0.4)
ax.bar_label(bars, fmt="%.1f%%", padding=3, fontsize=9)
ax.axhline(attrition_rate, color=ACCENT, linestyle="--", linewidth=1.2, label=f"Avg {attrition_rate:.1f}%")
ax.set_xlabel("Distance From Home")
ax.set_ylabel("Attrition Rate (%)")
ax.set_title("Attrition Rate by Distance From Home")
ax.legend(fontsize=8)
savefig("09_driver_distance.png")

# ── 5. Risk / Opportunity / Action ────────────────────────────────────────────
# Risk: Sales Representatives with OverTime
risk_seg  = df[(df["JobRole"] == "Sales Representative") & (df["OverTime"] == "Yes")]
risk_rate = risk_seg["AttritionFlag"].mean() * 100

# Opportunity: Research Directors (low attrition, high income)
opp_seg   = df[df["JobRole"] == "Research Director"]
opp_rate  = opp_seg["AttritionFlag"].mean() * 100
opp_inc   = opp_seg["MonthlyIncome"].mean()

print(f"\n-- Risk / Opportunity / Action -------------------------------------")
print(f"  RISK:        Sales Representatives + OverTime -> {risk_rate:.1f}% attrition")
print(f"  OPPORTUNITY: Research Directors               -> {opp_rate:.1f}% attrition | Avg income ${opp_inc:,.0f}")
print(f"  ACTION:      Reduce mandatory OverTime for Sales Representatives")
print(f"               (target: bring OverTime attrition below company average of {attrition_rate:.1f}%)")

# Risk bar chart
fig, ax = plt.subplots(figsize=(6, 4))
seg_data = {
    f"Sales Rep\n+ OverTime\n(Risk)": risk_rate,
    f"Company\nAverage": attrition_rate,
    f"Research\nDirector\n(Opportunity)": opp_rate,
}
colors_seg = [ACCENT, GREY, MID_BLUE]
bars = ax.bar(seg_data.keys(), seg_data.values(), color=colors_seg, width=0.4)
ax.bar_label(bars, fmt="%.1f%%", padding=3, fontsize=10)
ax.set_ylabel("Attrition Rate (%)")
ax.set_title("Risk vs Opportunity Segments")
savefig("10_risk_opportunity.png")

# ── 6. Auto-generate Word Report ──────────────────────────────────────────────
print("\nBuilding Word report ...")

doc = Document()

# ── Page margins (narrow) ─────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin   = Inches(1.0)
    section.right_margin  = Inches(1.0)

# ── Styles helper ─────────────────────────────────────────────────────────────
def set_heading_color(paragraph, hex_color):
    for run in paragraph.runs:
        run.font.color.rgb = RGBColor(
            int(hex_color[1:3], 16),
            int(hex_color[3:5], 16),
            int(hex_color[5:7], 16),
        )

def add_section_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    set_heading_color(p, BLUE)
    return p

def add_body(doc, text):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(6)
    return p

def add_chart(doc, filename, width=5.5):
    path = os.path.join(CHARTS_DIR, filename)
    if os.path.exists(path):
        doc.add_picture(path, width=Inches(width))
        last_para = doc.paragraphs[-1]
        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph(f"[Chart not found: {filename}]")

def add_kpi_table(doc, kpis_dict):
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    hdr[0].text = "KPI"
    hdr[1].text = "Value"
    for cell in hdr:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
        cell._element.get_or_add_tcPr().append(
            OxmlElement("w:shd")
        )
        shd = cell._element.tcPr.find(qn("w:shd"))
        shd.set(qn("w:fill"), "1F4E79")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:val"), "clear")
    for k, v in kpis_dict.items():
        row = table.add_row().cells
        row[0].text = k
        row[1].text = v
    doc.add_paragraph()   # spacer

# ── Cover ─────────────────────────────────────────────────────────────────────
title_para = doc.add_heading("HR Employee Attrition Analysis", 0)
title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_heading_color(title_para, BLUE)

sub = doc.add_paragraph("Business Intelligence Report")
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.runs[0].font.size = Pt(13)
sub.runs[0].font.color.rgb = RGBColor(89, 89, 89)

doc.add_paragraph()   # spacer

# ── Section 1: Business Problem ───────────────────────────────────────────────
add_section_heading(doc, "1. Business Problem")
add_body(doc,
    "Employee attrition is one of the most costly and disruptive challenges facing organisations today. "
    "Replacing a single employee can cost between 50% and 200% of their annual salary when accounting for "
    "recruiting, onboarding, lost productivity, and institutional knowledge. This report analyses IBM's "
    "HR dataset of 1,470 employees to identify the scale of attrition, understand its primary drivers, "
    "and surface actionable recommendations for HR leadership and business management."
)

# ── Section 2: Executive Overview ─────────────────────────────────────────────
add_section_heading(doc, "2. Executive Overview")
add_body(doc,
    "The table below summarises the five core KPIs that frame this analysis. "
    "An overall attrition rate of 16.1%, well above the typical industry benchmark of 10–12%, "
    "signals a meaningful retention challenge that warrants targeted intervention."
)
add_kpi_table(doc, kpis)

# ── Section 3: Attrition Trends ───────────────────────────────────────────────
add_section_heading(doc, "3. Attrition Trends")

add_section_heading(doc, "3.1 By Department", level=2)
add_body(doc,
    "Sales exhibits the highest attrition rate among all departments, driven by a combination of "
    "high-pressure targets, commission variability, and significant overtime demands. "
    "Human Resources also sits above the company average, while Research & Development "
    "demonstrates the strongest retention performance."
)
add_chart(doc, "01_attrition_by_department.png")

add_section_heading(doc, "3.2 By Job Role", level=2)
add_body(doc,
    f"Sales Representatives record the highest role-level attrition at {sales_rep_rate:.1f}%, "
    f"nearly {sales_rep_rate / attrition_rate:.1f}x the company average. "
    "Laboratory Technicians and Human Resources roles also exceed the average. "
    "Managerial and director-level roles consistently show attrition rates below 10%."
)
add_chart(doc, "02_attrition_by_jobrole.png")

add_section_heading(doc, "3.3 By Age Group", level=2)
add_body(doc,
    "Younger employees (18–25) leave at more than double the rate of employees aged 36 and above. "
    "This pattern is consistent with early-career mobility and suggests that targeted early-tenure "
    "engagement programmes could have a significant retention impact."
)
add_chart(doc, "03_attrition_by_agegroup.png")

add_section_heading(doc, "3.4 By Income Band", level=2)
add_body(doc,
    f"Attrition is sharply inversely correlated with income. Employees earning below $3,000/month "
    f"({income_low_rate:.1f}%) leave at a rate {income_ratio:.1f}x higher than those earning above "
    f"$10,000/month ({income_high_rate:.1f}%). "
    "This suggests compensation competitiveness is a material retention lever, particularly for "
    "entry-level and junior roles."
)
add_chart(doc, "04_attrition_by_incomeband.png")

# ── Section 4: Drivers of Attrition ──────────────────────────────────────────
add_section_heading(doc, "4. Drivers of Attrition")

add_section_heading(doc, "4.1 OverTime", level=2)
add_body(doc,
    "Employees required to work overtime leave at a rate of approximately 30%, more than double the "
    "rate of employees who do not work overtime (~10%). This is the single strongest binary driver "
    "of attrition in the dataset and represents the clearest lever for HR intervention."
)
add_chart(doc, "05_driver_overtime.png")

add_section_heading(doc, "4.2 Job Satisfaction", level=2)
add_body(doc,
    "There is a clear inverse relationship between job satisfaction and attrition. Employees at the "
    "lowest satisfaction level (1) leave at nearly 23%, while those at the highest level (4) leave "
    "at approximately 11%. Satisfaction surveys and role enrichment initiatives are well-supported "
    "by this data."
)
add_chart(doc, "06_driver_jobsatisfaction.png")

add_section_heading(doc, "4.3 Work-Life Balance", level=2)
add_body(doc,
    "Poor work-life balance (score 1) is associated with the highest attrition rate. Employees rating "
    "their work-life balance as 'Best' (4) still exhibit attrition above the overall average, "
    "indicating that balance alone does not fully explain attrition. It interacts with other factors "
    "such as compensation and role satisfaction."
)
add_chart(doc, "07_driver_worklifebalance.png")

add_section_heading(doc, "4.4 Time Since Last Promotion", level=2)
add_body(doc,
    "Employees who have not received a promotion in 7 or more years show elevated attrition, suggesting "
    "that career stagnation is a contributing factor. However, employees who were promoted very recently "
    "(0–1 year) also show higher-than-average attrition, potentially reflecting post-promotion "
    "reassessment or unmet expectations following a role change."
)
add_chart(doc, "08_driver_promotion.png")

add_section_heading(doc, "4.5 Distance From Home", level=2)
add_body(doc,
    "Employees who live furthest from the office (16–30 km) show a modestly higher attrition rate "
    "than those living nearby. While the effect size is smaller than overtime or satisfaction, "
    "it remains statistically relevant and supports the business case for flexible/remote working "
    "policies as a retention tool."
)
add_chart(doc, "09_driver_distance.png")

# ── Section 5: Risk, Opportunity & Recommended Action ────────────────────────
add_section_heading(doc, "5. Risk, Opportunity & Recommended Action")
add_chart(doc, "10_risk_opportunity.png")

add_section_heading(doc, "5.1 Risk: Sales Representatives Working Overtime", level=2)
add_body(doc,
    f"Sales Representatives who are required to work overtime record an attrition rate of "
    f"{risk_rate:.1f}%, approximately {risk_rate/attrition_rate:.1f}x the company average. "
    "This segment represents a concentrated, high-cost retention risk. The combination of "
    "high-pressure targets, variable compensation, and sustained overtime creates a compounding "
    "disengagement spiral. Losing experienced sales staff directly impacts revenue pipeline continuity."
)

add_section_heading(doc, "5.2 Opportunity: Research Directors", level=2)
add_body(doc,
    f"Research Directors exhibit an attrition rate of just {opp_rate:.1f}% with an average monthly "
    f"income of ${opp_inc:,.0f}. This cohort represents the organisation's most stable, high-value "
    "talent segment. Investment in their development, including leadership programmes, expanded scope, "
    "and mentoring roles, would yield compounding returns by retaining institutional knowledge and "
    "building internal succession pipelines."
)

add_section_heading(doc, "5.3 Recommended Action", level=2)
add_body(doc,
    "Primary Recommendation: Reduce mandatory overtime for Sales Representatives.\n\n"
    "Specifically, the organisation should:\n"
    "  1. Audit current overtime obligations across the Sales department.\n"
    "  2. Introduce a maximum overtime threshold (e.g. no more than 5 hours/week).\n"
    "  3. Pilot flexible scheduling for Sales Representatives in one region.\n"
    "  4. Track attrition monthly for 6 months post-intervention and compare against the current "
    f"{risk_rate:.1f}% baseline.\n\n"
    f"Target: Bring Sales Representative attrition below the company average of {attrition_rate:.1f}% "
    "within 12 months."
)

# ── Section 6: Conclusion ─────────────────────────────────────────────────────
add_section_heading(doc, "6. Conclusion")
add_body(doc,
    "This analysis reveals that employee attrition at this organisation is not random: it is "
    "concentrated in specific roles, income brackets, and behavioural patterns. The data consistently "
    "points to overtime, low job satisfaction, low compensation, and early career stage as the primary "
    "drivers. The Sales Representative + OverTime segment is the highest-priority risk requiring "
    "immediate action. Addressing this single cohort could meaningfully reduce overall attrition "
    "and deliver measurable cost savings within one fiscal year."
)
add_body(doc,
    "All charts and underlying data are reproducible by running KhushiVig_HRAttritionAnalysis.py against the source dataset."
)

# ── Save report ───────────────────────────────────────────────────────────────
report_path = "KhushiVig_ProjectReport.docx"
doc.save(report_path)
print(f"\n  Report saved -> {report_path}")
print("\nDone. All files generated successfully.")
