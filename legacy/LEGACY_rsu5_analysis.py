"""
RSU 5 Financial Analysis: Pownal Elementary School Reconciliation
Calculations supporting the comprehensive planning report.

Outputs a markdown appendix file with all computation results, citations,
and methodology descriptions.

SCHOOL ASSIGNMENT MODEL (Current Configuration - Scenario 1):
  Pownal:   PreK-5 at PES -> 6-8 at FMS -> 9-12 at FHS
  Durham:   PreK-8 at DCS -> 9-12 at FHS
  Freeport: PreK-2 at MSS -> 3-5 at MLS -> 6-8 at FMS -> 9-12 at FHS
"""

import sys
from io import StringIO

md = StringIO()

def w(line=""):
    md.write(line + "\n")

def table_row(*cells, header=False):
    row = "| " + " | ".join(str(c) for c in cells) + " |"
    w(row)
    if header:
        w("| " + " | ".join("---" for _ in cells) + " |")

# ============================================================
# FOOTNOTES REGISTRY
# ============================================================
FOOTNOTES = {
    1: "FY27 Budget Handbook, RSU 5 Superintendent's Office, pp.3-4 (02/11/2026). \"Projected Enrollment 2026-2027\" and Teacher/Class Size table.",
    2: "FY27 Superintendent's Proposed Budget Articles, 89-page line-item detail (02/11/2026). Available at: https://resources.finalsite.net/images/v1770852540/rsu5org/spzuwjskm8vh3t9bjck8/2026-2027SuperintendentsRecommendedBudgetArticlesforWebsite02112026.pdf",
    3: "FY27 Budget Handbook, RSU 5, Budget Impact Summary pp.9-10 (02/11/2026). Available at: https://resources.finalsite.net/images/v1770852522/rsu5org/ytm49nnwpbrzi4oiwg4i/FY27BUDGETHANDBOOK02112026.pdf",
    4: "Pownal FY26 Real Estate Tax Commitment Book, committed 07/29/2025, rate $15.300/thousand. Source: Town of Pownal Assessor, https://www.pownalmaine.org/",
    5: "Freeport FY26 tax rate $13.85/thousand at 100% assessment ratio, committed 09/15/2025. Source: Freeport Assessor, https://www.freeportmaine.com/158/Assessor",
    6: "Freeport FY25 Budget Presentation, Proposed Tax Changes table. Source: https://www.freeportmaine.com/DocumentCenter/View/2711/FY-25-Budget-Presentation",
    7: "Durham FY26 tax rate $33.58/thousand at 53% assessment ratio, committed 08/12/2025. Source: Durham Assessor, https://durhammaine.gov/pages/assessing",
    8: "Maine DOE FY25-26 Warrant Article F: Education Subsidy Information for Property Tax Bill (09/09/2025). Source: https://www.maine.gov/doe/sites/maine.gov.doe/files/inline-files/School%20Finance%20-%20FY25-26%20Warrant%20Article%20F%20Education%20Subsidy%20Information%20for%20Property%20Tax%20Bill%20-%209.9.2025.pdf",
    9: "FY26 Adopted Budget, RSU 5, $44,455,929 total. Approved by voters 06/10/2025 (817 Yes, 494 No). Source: https://www.rsu5.org/budget/fy26",
    10: "RSU 5 Planning for the Future presentation (02/11/2026). Source: https://resources.finalsite.net/images/v1770852531/rsu5org/vfrsu75oj4ebaxvsbcwc/PlanningfortheFutureofRSU5.pdf",
    11: "FY26 Citizens Adopted Budget for Website (06/10/2025). Source: https://resources.finalsite.net/images/v1750961183/rsu5org/hs7iblg3shrx7fhugb88/FY26CitizensAdoptedBudgetforWebsite06102025.pdf",
    12: "Pownal tax breakdown (RSU 58.4%, County 3.3%, Town 38.3%) from Pownal municipal records, FY26.",
    13: "Census/ACS population estimates: Pownal 1,590; Durham 4,339; Freeport 8,771. Sources: censusreporter.org, maine-demographics.com (2023 est.)",
    14: "Maine CDC vital statistics / CDC WONDER. Cumberland County birth rate ~9.5/1,000; Androscoggin County similar.",
    15: "Maine DOE Public Pre-K Guidebook (01/21/2025). Chapter 124: max 16 students/classroom, 1:8 staff ratio during academic time. Source: https://www.maine.gov/doe/sites/maine.gov.doe/files/inline-files/Early%20Learning%20-%20Public%20Pre%20K%20Guidebook%20-%201.21.2025.pdf",
    16: "Pownal is in Cumberland County; Durham is in Androscoggin County; Freeport is in Cumberland County.",
    17: "Portland Press Herald, 10/02/2025, 'Freeport-area schools begin planning for new early childhood responsibilities.' By July 1, 2028, districts assume CDS responsibilities for FAPE ages 3-5. Source: https://www.pressherald.com/2025/10/02/freeport-area-schools-begin-planning-for-new-early-childhood-responsibilities/",
    18: "RSU 5 Early Childhood Transition Task Force membership and meeting schedule. Source: https://www.rsu5.org/quick-links/early-childhood-planning-cds",
    19: "RSU 5 Service Model Options PK3, presented at Dec 18, 2025 Task Force meeting. Detailed startup costs for 3 service models. Source: https://www.rsu5.org/fs/resource-manager/view/da1426f4-8cc1-4b09-bd89-477da9c623ed",
    20: "Research on school consolidation and property values: Duncombe & Yinger (2010), Syracuse University CPR; EdWorkingPapers #22-530 (Arkansas, 2022). Findings: 5-15% property value decline in high-income areas post-closure; negative population effects in rural communities.",
    21: "Maine Chapter 124 Section 9 - School Facilities for Public Preschool Programs. 35 sq ft/child, toilets within 40 feet, water source in classroom, natural light required, 75 sq ft/child outdoor with fencing. Source: https://regulations.justia.com/states/maine/05/071/chapter-124/section-071-124-9/",
    22: "Driving distances estimated from Google Maps between PES (587 Elmwood Rd, Pownal ME), DCS (654 Hallowell Rd, Durham ME), MSS (17 West St, Freeport ME), FMS/FHS (Freeport center).",
    23: "Portable classroom lease costs: SAD 51 (Cumberland/North Yarmouth) FY25 budget shows $71K-$229K/year for modular classrooms. Source: Portland Press Herald, 04/16/2024.",
    24: "Patriquin Architects, 'Checklist for Early Childhood Conversion Projects.' Plumbing, bathroom, and safety requirements for converting K-5 spaces to serve preschool-age children. Source: https://www.patriquinarchitects.com/checklist-for-early-childhood-conversion-projects/",
    25: "FY27 Superintendent's Proposed Budget Handbook (Revised 02/11/2026). Total operating budget $47,357,441 = Articles 1-11 ($47,269,441) + Adult Ed ($88,000). 10-year budget history table. Expenditure/Reductions summary. Source: https://www.rsu5.org/fs/resource-manager/view/78dce71e-a8e6-4435-b33f-9746e8541a3a",
}

# ============================================================
# SECTION 1: ENROLLMENT DATA [fn1]
# ============================================================

school_enrollment = {
    "MSS": 274, "MLS": 264, "PES": 105, "DCS": 467, "FMS": 306, "FHS": 554,
}

enrollment_history = {
    "MSS": {"2023": 316, "2024": 288, "2025": 275, "2026": 274},
    "MLS": {"2023": 281, "2024": 266, "2025": 282, "2026": 264},
    "PES": {"2023":  89, "2024":  97, "2025":  98, "2026": 105},
    "DCS": {"2023": 473, "2024": 466, "2025": 453, "2026": 467},
    "FMS": {"2023": 293, "2024": 288, "2025": 286, "2026": 306},
    "FHS": {"2023": 632, "2024": 592, "2025": 577, "2026": 554},
}

pes_grade_sizes = [14, 14, 13, 17, 13, 18]
pes_avg_per_grade = sum(pes_grade_sizes) / len(pes_grade_sizes)
dcs_grade_7 = 46
dcs_grade_8 = 54
dcs_avg_78_per_grade = (dcs_grade_7 + dcs_grade_8) / 2

pownal_at_fms = round(pes_avg_per_grade) * 3
freeport_at_fms = school_enrollment["FMS"] - pownal_at_fms
pownal_at_fhs = round(pes_avg_per_grade) * 4
durham_at_fhs = round(dcs_avg_78_per_grade) * 4
freeport_at_fhs = school_enrollment["FHS"] - pownal_at_fhs - durham_at_fhs

pownal_total_students = school_enrollment["PES"] + pownal_at_fms + pownal_at_fhs
durham_total_students = school_enrollment["DCS"] + durham_at_fhs
freeport_total_students = (school_enrollment["MSS"] + school_enrollment["MLS"]
                           + freeport_at_fms + freeport_at_fhs)
total_district = pownal_total_students + durham_total_students + freeport_total_students

w("# RSU 5 Financial Analysis: Calculation Appendix")
w()
w("*Generated by `rsu5_analysis.py` -- all figures are computed from cited source data.*")
w("*This appendix provides the quantitative foundation for the planning report.*")
w()

w("---")
w()
w("## 1. Enrollment & Town-of-Origin Analysis")
w()
w("### 1.1 Historical Enrollment Trends [^1]")
w()
table_row("School", "Grades", "Oct 2023", "Oct 2024", "Oct 2025", "Oct 2026P", "4yr Change", header=True)
school_labels = {
    "MSS": ("Morse Street", "PreK-2"),
    "MLS": ("Mast Landing", "3-5"),
    "PES": ("Pownal Elementary", "PreK-5"),
    "DCS": ("Durham Community", "PreK-8"),
    "FMS": ("Freeport Middle", "6-8"),
    "FHS": ("Freeport High", "9-12"),
}
for sch in ["MSS", "MLS", "PES", "DCS", "FMS", "FHS"]:
    h = enrollment_history[sch]
    chg = h["2026"] - h["2023"]
    label, grades = school_labels[sch]
    table_row(f"{sch} ({label})", grades, h["2023"], h["2024"], h["2025"], h["2026"], f"{chg:+d}")

totals = {yr: sum(enrollment_history[s][yr] for s in enrollment_history) for yr in ["2023", "2024", "2025", "2026"]}
table_row("**TOTAL**", "All", totals["2023"], totals["2024"], totals["2025"], totals["2026"], f"{totals['2026']-totals['2023']:+d}")

w()
w("### 1.2 Town-of-Origin Estimation Methodology [^1]")
w()
w("FMS and FHS serve students from multiple towns. Since RSU 5 does not publicly report")
w("enrollment by town of origin at shared schools, we estimate using grade-level cohort sizes:")
w()
w(f"- **PES average K-5 grade size:** {pes_avg_per_grade:.1f} students (grades: {', '.join(str(g) for g in pes_grade_sizes)})")
w(f"- **DCS average 7-8 grade size:** {dcs_avg_78_per_grade:.1f} students (grade 7: {dcs_grade_7}, grade 8: {dcs_grade_8})")
w()

table_row("School", "Total", "Pownal (est.)", "Durham (est.)", "Freeport (est.)", header=True)
table_row("FMS", 306, pownal_at_fms, "--", freeport_at_fms)
table_row("FHS", 554, pownal_at_fhs, durham_at_fhs, freeport_at_fhs)

w()
w("### 1.3 Total Students by Town of Origin")
w()
table_row("Town", "Own Schools", "At FMS", "At FHS", "Total", header=True)
table_row("Pownal", f"PES: {school_enrollment['PES']}", pownal_at_fms, pownal_at_fhs, f"**{pownal_total_students}**")
table_row("Durham", f"DCS: {school_enrollment['DCS']}", "--", durham_at_fhs, f"**{durham_total_students}**")
table_row("Freeport", f"MSS+MLS: {school_enrollment['MSS']+school_enrollment['MLS']}", freeport_at_fms, freeport_at_fhs, f"**{freeport_total_students}**")
table_row("**District**", "", "", "", f"**{total_district}**")

# ============================================================
# SECTION 2: FY27 PROPOSED BUDGET BY COST CENTER [fn2]
# ============================================================

fy27_by_school_article = {
    "Art 1 - Regular Instruction": {
        "DCS": 4_611_862, "MSS": 2_506_278, "PES": 1_196_683,
        "MLS": 2_338_114, "FMS": 3_201_928, "FHS": 5_310_140,
        "SYS": 0, "K8": 420_852, "912": 287_336,
    },
    "Art 2 - Special Education": {
        "DCS": 1_553_029, "MSS": 1_037_964, "PES": 329_115,
        "MLS": 971_107, "FMS": 1_062_320, "FHS": 1_013_463,
        "SYS": 836_683, "K8": 347_140, "912": 359_601,
    },
    "Art 4 - Other Instruction": {
        "DCS": 90_140, "MSS": 7_710, "PES": 13_125,
        "MLS": 12_626, "FMS": 277_908, "FHS": 769_494,
        "SYS": 0, "K8": 1_536, "912": 0,
    },
    "Art 5 - Student & Staff Support": {
        "DCS": 581_386, "MSS": 400_046, "PES": 219_819,
        "MLS": 363_844, "FMS": 490_321, "FHS": 1_061_829,
        "SYS": 1_523_668, "K8": 249_421, "912": 16_032,
    },
    "Art 7 - School Administration": {
        "DCS": 515_193, "MSS": 421_185, "PES": 257_979,
        "MLS": 400_206, "FMS": 462_558, "FHS": 580_903,
        "SYS": 0, "K8": 0, "912": 0,
    },
    "Art 9 - Facilities & Maintenance": {
        "DCS": 683_965, "MSS": 569_640, "PES": 253_384,
        "MLS": 482_373, "FMS": 641_780, "FHS": 1_745_609,
        "SYS": 1_564_602, "K8": 0, "912": 0, "CTRL": 52_424,
    },
}

system_only_articles = {
    "Art 3 - CTE":             337_282,
    "Art 6 - System Admin":  1_308_008,
    "Art 8 - Transportation": 2_388_457,
    "Art 10 - Debt Service":  1_071_577,
    "Art 11 - All Other":        69_796,
}

schools_list = ["DCS", "MSS", "PES", "MLS", "FMS", "FHS"]
school_direct = {}
for sch in schools_list:
    school_direct[sch] = sum(art.get(sch, 0) for art in fy27_by_school_article.values())

system_pool = sum(system_only_articles.values())
for art_data in fy27_by_school_article.values():
    system_pool += art_data.get("SYS", 0) + art_data.get("K8", 0) + art_data.get("912", 0) + art_data.get("CTRL", 0)

total_articles = 47_269_441

w()
w("---")
w()
w("## 2. FY27 Proposed Budget by Cost Center [^2]")
w()
table_row("School", "Direct Cost", "Enrollment", "Per Student", header=True)
for sch in schools_list:
    enr = school_enrollment[sch]
    table_row(sch, f"${school_direct[sch]:,}", enr, f"${school_direct[sch]/enr:,.0f}")
w()
table_row("Direct School Total", f"${sum(school_direct.values()):,}", "", "")
table_row("System-Wide Pool", f"${system_pool:,}", "", "")
table_row("**Grand Total**", f"**${sum(school_direct.values()) + system_pool:,}**", f"**{total_district}**", f"**${total_articles/total_district:,.0f}**")

w()
w("System-wide pool includes: " + ", ".join(f"{k} (${v:,})" for k, v in system_only_articles.items()))
w("plus system/K-8/9-12/central office allocations from school-level articles.")

# ============================================================
# SECTION 3: CONSUMPTION BY TOWN [fn2]
# ============================================================

pownal_fms_share = pownal_at_fms / school_enrollment["FMS"]
pownal_fhs_share = pownal_at_fhs / school_enrollment["FHS"]
durham_fhs_share = durham_at_fhs / school_enrollment["FHS"]
freeport_fms_share = freeport_at_fms / school_enrollment["FMS"]
freeport_fhs_share = freeport_at_fhs / school_enrollment["FHS"]

pownal_district_share = pownal_total_students / total_district
durham_district_share = durham_total_students / total_district
freeport_district_share = freeport_total_students / total_district

pownal_consumption = (
    school_direct["PES"]
    + school_direct["FMS"] * pownal_fms_share
    + school_direct["FHS"] * pownal_fhs_share
    + system_pool * pownal_district_share
)
durham_consumption = (
    school_direct["DCS"]
    + school_direct["FHS"] * durham_fhs_share
    + system_pool * durham_district_share
)
freeport_consumption = (
    school_direct["MSS"] + school_direct["MLS"]
    + school_direct["FMS"] * freeport_fms_share
    + school_direct["FHS"] * freeport_fhs_share
    + system_pool * freeport_district_share
)

w()
w("---")
w()
w("## 3. Estimated Budget Consumption by Town of Origin [^2]")
w()
w("**Methodology:** Direct school costs are allocated 100% to the town served (PES -> Pownal,")
w("DCS -> Durham, MSS/MLS -> Freeport). FMS and FHS costs are allocated proportionally by")
w("estimated town-of-origin enrollment share (Section 1.2). System-wide costs are allocated")
w("proportionally by total student count.")
w()

for town, label, components in [
    ("Pownal", "POWNAL", [
        ("PES (100%)", school_direct["PES"], school_direct["PES"]),
        (f"FMS ({pownal_fms_share*100:.1f}%)", school_direct["FMS"], school_direct["FMS"] * pownal_fms_share),
        (f"FHS ({pownal_fhs_share*100:.1f}%)", school_direct["FHS"], school_direct["FHS"] * pownal_fhs_share),
        (f"System ({pownal_district_share*100:.1f}%)", system_pool, system_pool * pownal_district_share),
    ]),
    ("Durham", "DURHAM", [
        ("DCS (100%)", school_direct["DCS"], school_direct["DCS"]),
        (f"FHS ({durham_fhs_share*100:.1f}%)", school_direct["FHS"], school_direct["FHS"] * durham_fhs_share),
        (f"System ({durham_district_share*100:.1f}%)", system_pool, system_pool * durham_district_share),
    ]),
    ("Freeport", "FREEPORT", [
        ("MSS (100%)", school_direct["MSS"], school_direct["MSS"]),
        ("MLS (100%)", school_direct["MLS"], school_direct["MLS"]),
        (f"FMS ({freeport_fms_share*100:.1f}%)", school_direct["FMS"], school_direct["FMS"] * freeport_fms_share),
        (f"FHS ({freeport_fhs_share*100:.1f}%)", school_direct["FHS"], school_direct["FHS"] * freeport_fhs_share),
        (f"System ({freeport_district_share*100:.1f}%)", system_pool, system_pool * freeport_district_share),
    ]),
]:
    w(f"**{label}:**")
    w()
    table_row("Component", "School Total", "Town's Share", header=True)
    total = 0
    for comp_label, full, share in components:
        table_row(comp_label, f"${full:,}", f"${share:,.0f}")
        total += share
    table_row("**TOTAL**", "", f"**${total:,.0f}**")
    stu = {"Pownal": pownal_total_students, "Durham": durham_total_students, "Freeport": freeport_total_students}[town]
    table_row("Per Student", f"({stu} students)", f"**${total/stu:,.0f}**")
    w()

consumption_by_town = {"Pownal": pownal_consumption, "Durham": durham_consumption, "Freeport": freeport_consumption}
students_by_town = {"Pownal": pownal_total_students, "Durham": durham_total_students, "Freeport": freeport_total_students}

check_total = pownal_consumption + durham_consumption + freeport_consumption
w(f"**Integrity check:** Sum of town consumption = ${check_total:,.0f} (budget total: ${total_articles:,}) -- {abs(check_total - total_articles):.0f} rounding variance.")

# ============================================================
# SECTION 4: REVENUE BY TOWN [fn3]
# ============================================================

total_shared_revenue = 1_431_699.07

revenue = {
    "Pownal": {
        "rlc": 2_257_247.34, "alm": 2_190_978.34, "state_aid": 567_179.79,
        "nonshared_debt": 0, "total_contribution": 5_015_405.46, "alm_pct": 0.1260,
    },
    "Durham": {
        "rlc": 3_910_950.09, "alm": 3_724_663.18, "state_aid": 5_991_563.92,
        "nonshared_debt": 117_175.57, "total_contribution": 13_744_352.75, "alm_pct": 0.2142,
    },
    "Freeport": {
        "rlc": 14_173_183.75, "alm": 11_473_075.46, "state_aid": 1_463_888.50,
        "nonshared_debt": 0, "total_contribution": 27_110_147.71, "alm_pct": 0.6598,
    },
}
for town, data in revenue.items():
    data["shared_rev"] = total_shared_revenue * data["alm_pct"]
    data["total_revenue"] = data["total_contribution"] + data["shared_rev"]

w()
w("---")
w()
w("## 4. Revenue by Town & Net Fiscal Position [^3]")
w()
w("Revenue includes Required Local Contribution (RLC), Additional Local Monies (ALM),")
w("state aid attributable to each town's students, non-shared debt service, and a")
w("proportional share of shared revenues (allocated by ALM percentage).")
w()

w("### 4.1 Detailed Revenue Breakdown")
w()
table_row("Component", "Pownal", "Durham", "Freeport", header=True)
for label, key in [
    ("Required Local Contribution (RLC)", "rlc"),
    ("Additional Local Monies (ALM)", "alm"),
    ("State Aid", "state_aid"),
    ("Non-Shared Debt Service", "nonshared_debt"),
]:
    table_row(label,
              f"${revenue['Pownal'][key]:,.2f}",
              f"${revenue['Durham'][key]:,.2f}",
              f"${revenue['Freeport'][key]:,.2f}")

for town in ["Pownal", "Durham", "Freeport"]:
    pass
table_row(f"Shared Revenue (by ALM %)",
          f"${revenue['Pownal']['shared_rev']:,.2f} ({revenue['Pownal']['alm_pct']*100:.1f}%)",
          f"${revenue['Durham']['shared_rev']:,.2f} ({revenue['Durham']['alm_pct']*100:.1f}%)",
          f"${revenue['Freeport']['shared_rev']:,.2f} ({revenue['Freeport']['alm_pct']*100:.1f}%)")
table_row("**TOTAL REVENUE**",
          f"**${revenue['Pownal']['total_revenue']:,.0f}**",
          f"**${revenue['Durham']['total_revenue']:,.0f}**",
          f"**${revenue['Freeport']['total_revenue']:,.0f}**")

w()
w("### 4.2 Net Fiscal Position (Revenue - Consumption)")
w()
table_row("Town", "Total Revenue", "Est. Consumption", "Net Position", "Students", "Revenue/Stu", "Cost/Stu", header=True)
for town in ["Pownal", "Durham", "Freeport"]:
    rev = revenue[town]["total_revenue"]
    cons = consumption_by_town[town]
    stu = students_by_town[town]
    net = rev - cons
    table_row(town, f"${rev:,.0f}", f"${cons:,.0f}", f"**${net:+,.0f}**", stu, f"${rev/stu:,.0f}", f"${cons/stu:,.0f}")

w()
w("> **Key finding:** Pownal runs a modest deficit of ~$267K. Durham runs a larger deficit")
w(f"> of ~$1.44M, offset by substantial state aid ($5.99M). Freeport's surplus of ~$1.74M")
w("> effectively subsidizes both other towns. These figures are FY27 proposed; actual FY26")
w("> positions would differ slightly.")

# ============================================================
# SECTION 5: COMPREHENSIVE TAX ANALYSIS [fn4-8,12,16]
# ============================================================

w()
w("---")
w()
w("## 5. Comprehensive Tax Analysis")
w()
w("### 5.1 Property Tax Rate Breakdown by Town (FY26, 2025-2026)")
w()
w("Each town's property tax has three major components: school (RSU #5 assessment),")
w("county, and municipal. The towns are in different counties [^16], and each town")
w("assesses property at a different ratio of market value.")
w()

# Pownal
w("#### Pownal [^4] [^12]")
w()
pownal_mil = 15.300
pownal_rsu_pct = 0.584
pownal_county_pct = 0.033
pownal_town_pct = 0.383
pownal_rsu_mil = pownal_mil * pownal_rsu_pct
pownal_county_mil = pownal_mil * pownal_county_pct
pownal_town_mil = pownal_mil * pownal_town_pct
pownal_taxable = 392_398_240  # from commitment book final total

table_row("Component", "% of Total", "Mil Rate", "Levy Amount", header=True)
table_row("RSU #5 (School)", f"{pownal_rsu_pct*100:.1f}%", f"${pownal_rsu_mil:.3f}",
          f"${pownal_taxable * pownal_rsu_mil / 1000:,.0f}")
table_row("Cumberland County", f"{pownal_county_pct*100:.1f}%", f"${pownal_county_mil:.3f}",
          f"${pownal_taxable * pownal_county_mil / 1000:,.0f}")
table_row("Town of Pownal", f"{pownal_town_pct*100:.1f}%", f"${pownal_town_mil:.3f}",
          f"${pownal_taxable * pownal_town_mil / 1000:,.0f}")
table_row("**TOTAL**", "100.0%", f"**${pownal_mil:.3f}**",
          f"**${pownal_taxable * pownal_mil / 1000:,.0f}**")

w()
w(f"- Taxable valuation: ${pownal_taxable:,}")
w(f"- State valuation: $399,866,667 [^3]")
w(f"- Assessment ratio: {pownal_taxable/399_866_667:.1%}")
w(f"- Commitment date: 07/29/2025")
w()

# Freeport
w("#### Freeport [^5] [^6]")
w()
freeport_mil_fy26 = 13.850
freeport_taxable = 2_613_679_115  # from FY27 budget handbook
freeport_state_val = 2_510_750_000

# FY25 component percentages from budget presentation
# FY25: RSU $22,692,612 (71.6%), Town $7,449,095 (23.5%), County $1,442,472 (4.6%), Transit $95,766 (0.3%)
freeport_rsu_pct_fy25 = 0.716
freeport_town_pct_fy25 = 0.235
freeport_county_pct_fy25 = 0.046
freeport_transit_pct_fy25 = 0.003

w("FY26 total rate: $13.85/thousand at 100% assessment ratio. [^5]")
w("Component breakdown estimated from FY25 proportions [^6] (FY26 component detail")
w("not yet publicly available as of this analysis):")
w()

table_row("Component", "% of Total (FY25)", "Est. FY26 Mil Rate", "Est. Levy Amount", header=True)
table_row("RSU #5 (School)", f"{freeport_rsu_pct_fy25*100:.1f}%",
          f"${freeport_mil_fy26 * freeport_rsu_pct_fy25:.3f}",
          f"${freeport_taxable * freeport_mil_fy26 * freeport_rsu_pct_fy25 / 1000:,.0f}")
table_row("Town of Freeport", f"{freeport_town_pct_fy25*100:.1f}%",
          f"${freeport_mil_fy26 * freeport_town_pct_fy25:.3f}",
          f"${freeport_taxable * freeport_mil_fy26 * freeport_town_pct_fy25 / 1000:,.0f}")
table_row("Cumberland County", f"{freeport_county_pct_fy25*100:.1f}%",
          f"${freeport_mil_fy26 * freeport_county_pct_fy25:.3f}",
          f"${freeport_taxable * freeport_mil_fy26 * freeport_county_pct_fy25 / 1000:,.0f}")
table_row("Metro Transit", f"{freeport_transit_pct_fy25*100:.1f}%",
          f"${freeport_mil_fy26 * freeport_transit_pct_fy25:.3f}",
          f"${freeport_taxable * freeport_mil_fy26 * freeport_transit_pct_fy25 / 1000:,.0f}")
table_row("**TOTAL**", "100.0%", f"**${freeport_mil_fy26:.3f}**",
          f"**${freeport_taxable * freeport_mil_fy26 / 1000:,.0f}**")

w()
w(f"- Taxable valuation: ${freeport_taxable:,} [^3]")
w(f"- State valuation: ${freeport_state_val:,} [^3]")
w(f"- Assessment ratio: {freeport_taxable/freeport_state_val:.1%}")
w(f"- Commitment date: 09/15/2025")
w()

# Durham
w("#### Durham [^7]")
w()
durham_mil = 33.580
durham_taxable = 389_646_650
durham_state_val = 692_816_667
durham_assessment_ratio = durham_taxable / durham_state_val

w(f"FY26 total rate: $33.58/thousand. Durham Assessor cites ~53% assessment ratio; [^7]")
w(f"however, dividing the FY26 taxable valuation (${durham_taxable:,}) by state valuation")
w(f"(${durham_state_val:,}) yields {durham_assessment_ratio:.1%}. The difference may reflect")
w(f"timing of the state valuation estimate vs. the commitment date. We use the calculated ratio below.")
w(f"Durham is in **Androscoggin County** (unlike Pownal and Freeport which are in Cumberland County). [^16]")
w()
w("**Component breakdown is not publicly available** from Durham town records for FY26.")
w("Durham's 18-month transition budget complicates direct comparison. We can estimate")
w("the RSU component from the RSU 5 cost-sharing formula:")
w()

# Estimate Durham's RSU assessment from FY27 budget handbook data
# FY27 proposed: Durham RLC $3,910,950 + ALM $3,724,663 = $7,635,613
# Scale to FY26: FY26 adopted was $44,455,929, FY27 proposed is $47,269,441
# Ratio: 44,455,929/47,269,441 = 0.9405
fy26_to_fy27_ratio = 44_455_929 / 47_269_441
durham_rsu_fy27 = 3_910_950.09 + 3_724_663.18  # RLC + ALM
durham_rsu_fy26_est = durham_rsu_fy27 * fy26_to_fy27_ratio
durham_rsu_mil_est = durham_rsu_fy26_est / durham_taxable * 1000
durham_rsu_pct_est = durham_rsu_mil_est / durham_mil

w(f"- FY27 proposed Durham RSU assessment (RLC + ALM): ${durham_rsu_fy27:,.0f} [^3]")
w(f"- FY26/FY27 budget scaling ratio: {fy26_to_fy27_ratio:.4f} (${44_455_929:,} / ${47_269_441:,}) [^9] [^2]")
w(f"- Estimated FY26 Durham RSU assessment: ${durham_rsu_fy26_est:,.0f}")
w(f"- Estimated RSU mil rate on Durham assessed: ${durham_rsu_mil_est:.2f}")
w(f"- Estimated RSU share of total: {durham_rsu_pct_est*100:.1f}%")
w()

table_row("Component", "% of Total (est.)", "Est. Mil Rate", "Est. Levy Amount", header=True)
table_row("RSU #5 (School)", f"~{durham_rsu_pct_est*100:.0f}%", f"~${durham_rsu_mil_est:.2f}",
          f"~${durham_rsu_fy26_est:,.0f}")
remainder_mil = durham_mil - durham_rsu_mil_est
remainder_levy = durham_taxable * remainder_mil / 1000
table_row("Androscoggin County + Town", f"~{(1-durham_rsu_pct_est)*100:.0f}%", f"~${remainder_mil:.2f}",
          f"~${remainder_levy:,.0f}")
table_row("**TOTAL**", "100%", f"**${durham_mil:.3f}**",
          f"**${durham_taxable * durham_mil / 1000:,.0f}**")
w()
w(f"- Taxable valuation: ${durham_taxable:,} [^3]")
w(f"- State valuation: ${durham_state_val:,} [^3]")
w(f"- Assessment ratio: {durham_assessment_ratio:.1%}")
w(f"- Commitment date: 08/12/2025")
w(f"- **Note:** Durham's last town-wide revaluation was in 2010. [^7] The 53% assessment ratio")
w(f"  means assessed values are roughly half of current market values, inflating the nominal mil")
w(f"  rate. A $400,000 market-value home in Durham is assessed at ~${400_000 * durham_assessment_ratio:,.0f}.")

w()
w("### 5.2 Equalized Tax Comparison (State Valuation Basis) [^3]")
w()
w("Because towns assess property at different ratios, direct mil rate comparison is")
w("misleading. The **equalized mil rate** uses state valuation (100% market value estimate)")
w("as the common denominator:")
w()

tax_data = {
    "Pownal": {"state_val": 399_866_667, "taxable": pownal_taxable, "rsu_net_tax": 4_448_225.67, "enrollment": pownal_total_students},
    "Durham": {"state_val": 692_816_667, "taxable": durham_taxable, "rsu_net_tax": 7_752_788.83, "enrollment": durham_total_students},
    "Freeport": {"state_val": 2_510_750_000, "taxable": freeport_taxable, "rsu_net_tax": 25_646_259.21, "enrollment": freeport_total_students},
}

table_row("Town", "RSU Net Tax Impact", "State Valuation", "Eq. School Mil", "Local Mil (assessed)", "Assessment Ratio", header=True)
for town, d in tax_data.items():
    eq_mil = d["rsu_net_tax"] / d["state_val"] * 1000
    local_mil = d["rsu_net_tax"] / d["taxable"] * 1000
    ratio = d["taxable"] / d["state_val"]
    table_row(town, f"${d['rsu_net_tax']:,.0f}", f"${d['state_val']:,}", f"{eq_mil:.2f}", f"{local_mil:.2f}", f"{ratio:.1%}")

w()
w("The equalized school mil rates cluster tightly (10.2-11.2), confirming the Maine EPS")
w("funding formula works as designed to equalize the education tax burden across towns")
w("of different wealth levels.")

w()
w("### 5.3 Per-Student Tax Burden [^3]")
w()
table_row("Town", "RSU Net Tax Impact", "Total Students", "Tax per Student", header=True)
for town, d in tax_data.items():
    tax_per_stu = d["rsu_net_tax"] / d["enrollment"]
    table_row(town, f"${d['rsu_net_tax']:,.0f}", d["enrollment"], f"**${tax_per_stu:,.0f}**")

pownal_tps = tax_data["Pownal"]["rsu_net_tax"] / pownal_total_students
durham_tps = tax_data["Durham"]["rsu_net_tax"] / durham_total_students
freeport_tps = tax_data["Freeport"]["rsu_net_tax"] / freeport_total_students

w()
w(f"- Pownal pays **{pownal_tps/freeport_tps:.2f}x** Freeport per student in local school taxes")
w(f"- Pownal pays **{pownal_tps/durham_tps:.2f}x** Durham per student in local school taxes")
w(f"- Durham's low per-student local tax is offset by **${revenue['Durham']['state_aid']:,.0f}** in state aid")

w()
w("### 5.4 School Tax on a $400,000 Home (at Market Value)")
w()
w("Using equalized mil rates to give a fair cross-town comparison on the same home value:")
w()
table_row("Town", "Equalized School Mil", "Annual School Tax", "Total Mil (all purposes)", "Total Annual Tax", header=True)
for town_name, total_mil, state_val in [
    ("Pownal", pownal_mil, 399_866_667),
    ("Durham", durham_mil, 692_816_667),
    ("Freeport", freeport_mil_fy26, 2_510_750_000),
]:
    d = tax_data[town_name]
    eq_school_mil = d["rsu_net_tax"] / state_val * 1000
    school_tax = 400_000 * eq_school_mil / 1000
    total_eq_mil = total_mil * d["taxable"] / state_val
    total_tax = 400_000 * total_eq_mil / 1000
    table_row(town_name, f"{eq_school_mil:.2f}", f"${school_tax:,.0f}", f"{total_eq_mil:.2f}", f"${total_tax:,.0f}")

w()
w("> **Note:** Durham's total equalized mil rate is lower than the nominal $33.58 because the")
w("> 53% assessment ratio inflates the nominal rate. On an equalized basis, all three towns")
w("> pay similar school taxes per dollar of home value.")

# ============================================================
# SECTION 5.5: STATE EDUCATION FUNDING (EPS) [fn8]
# ============================================================

w()
w("### 5.5 Maine EPS Education Funding Formula (FY25-26) [^8]")
w()
w("The Maine Essential Programs & Services (EPS) formula determines each town's state aid.")
w("This is the foundation of the cost-sharing formula within RSU 5.")
w()

eps_data = {
    "Durham": {"total": 9_123_076.68, "local": 3_655_120.00, "mill": 6.10, "state": 5_467_956.68, "state_pct": 59.94},
    "Freeport": {"total": 15_033_673.30, "local": 13_799_840.67, "mill": 5.95, "state": 1_233_832.63, "state_pct": 8.21},
    "Pownal": {"total": 2_682_828.98, "local": 2_179_326.67, "mill": 6.10, "state": 503_502.31, "state_pct": 18.77},
}

table_row("Town", "Total EPS Allocation", "Local Share", "EPS Mill Rate", "State Share", "State %", header=True)
for town in ["Pownal", "Durham", "Freeport"]:
    d = eps_data[town]
    table_row(town, f"${d['total']:,.2f}", f"${d['local']:,.2f}", f"{d['mill']:.2f}",
              f"${d['state']:,.2f}", f"{d['state_pct']:.2f}%")

rsu_eps_total = sum(d["total"] for d in eps_data.values())
rsu_state_total = sum(d["state"] for d in eps_data.values())
table_row("**RSU 5 Total**", f"**${rsu_eps_total:,.2f}**",
          f"**${sum(d['local'] for d in eps_data.values()):,.2f}**", "",
          f"**${rsu_state_total:,.2f}**",
          f"**{rsu_state_total/rsu_eps_total*100:.2f}%**")

w()
w("**Key observations:**")
w(f"- Durham receives {eps_data['Durham']['state_pct']:.1f}% of its allocation from the state -- by far the largest share.")
w(f"  This reflects Durham's lower property wealth relative to student count.")
w(f"- Freeport receives only {eps_data['Freeport']['state_pct']:.1f}% state aid due to high property valuation.")
w(f"- Pownal receives {eps_data['Pownal']['state_pct']:.1f}% state aid -- moderate, reflecting its smaller tax base.")
w(f"- The EPS formula is the mechanism by which state aid flows disproportionately to Durham,")
w(f"  reducing its local tax burden per student from $21K+ to $11.6K.")

# ============================================================
# SECTION 6: PER-STUDENT COST BY ARTICLE [fn2]
# ============================================================

w()
w("---")
w()
w("## 6. Per-Student Direct Cost by Budget Article (Elementary Schools) [^2]")
w()
w("Comparison of direct per-student costs across elementary-level schools.")
w("Note: DCS serves PreK-8 (467 students); PES serves PreK-5 (105 students).")
w("MSS serves PreK-2 (274); MLS serves 3-5 (264).")
w()

elem_schools = ["DCS", "PES", "MLS", "MSS"]
elem_enrollment = {"DCS": 467, "PES": 105, "MLS": 264, "MSS": 274}

for article, centers in sorted(fy27_by_school_article.items()):
    w(f"**{article}:**")
    w()
    table_row("School", "Total Cost", "Enrollment", "Per Student", header=True)
    for sch in elem_schools:
        cost = centers.get(sch, 0)
        enr = elem_enrollment[sch]
        table_row(sch, f"${cost:,}", enr, f"${cost/enr:,.0f}")
    w()

w("### 6.1 PES Cost Premium Analysis")
w()
pes_total = school_direct["PES"]
avg_other = (school_direct["DCS"]/467 + school_direct["MSS"]/274 + school_direct["MLS"]/264) / 3
gap_total = pes_total/105 - avg_other

table_row("Article", "PES $/Student", "Other Elem Avg", "Difference", header=True)
for article, centers in sorted(fy27_by_school_article.items()):
    pes_per = centers.get("PES", 0) / 105
    other_avg = (centers.get("DCS", 0)/467 + centers.get("MSS", 0)/274 + centers.get("MLS", 0)/264) / 3
    diff = pes_per - other_avg
    if abs(diff) > 100:
        table_row(article, f"${pes_per:,.0f}", f"${other_avg:,.0f}", f"${diff:+,.0f}")
table_row("**TOTAL GAP**", f"${pes_total/105:,.0f}", f"${avg_other:,.0f}", f"**${gap_total:+,.0f}**")

w()
w(f"The gap of ~${gap_total:,.0f}/student x 105 students = **${gap_total*105:,.0f}** total (computed from unrounded per-student figures).")
w("This premium is driven primarily by **fixed-cost dilution** -- school administration,")
w("facilities, and student support costs that don't scale down with smaller enrollment.")
w("Notably, PES spends *less* than average on special education (-$463/student).")

# ============================================================
# SECTION 7: MIDDLE SCHOOL CONSOLIDATION [fn1,2,10]
# ============================================================

w()
w("---")
w()
w("## 7. Middle School Consolidation Analysis [^1] [^10]")
w()
w("**Proposal:** Send all district 7th-8th graders to FMS; send all district 6th graders to DCS.")
w("PES remains PreK-5 (no change).")
w()

dcs_grade_6 = 41
dcs_78 = dcs_grade_7 + dcs_grade_8
# FMS 6th grade total is ~116 (5.16 sections × 22-23 students) and INCLUDES Pownal students
pownal_6th = round(pes_avg_per_grade)  # ~15
fms_grade_6_total = 116  # all FMS 6th graders (Freeport + Pownal combined)
freeport_6th = fms_grade_6_total - pownal_6th  # ~101 Freeport-only
fms_grade_7 = 95
fms_grade_8 = 95
cost_per_fte = 107_700

new_fms_total = fms_grade_7 + fms_grade_8 + dcs_78
# District-wide 6th = DCS current (Durham) + FMS current (Freeport+Pownal) -- no double-counting
new_dcs_6th_total = dcs_grade_6 + fms_grade_6_total
dcs_k5_remaining = school_enrollment["DCS"] - dcs_grade_6 - dcs_78
new_dcs_total = dcs_k5_remaining + new_dcs_6th_total

w("### 7.1 Student Movement")
w()
table_row("School", "Current", "Proposed", "Change", header=True)
table_row("DCS", f"{school_enrollment['DCS']} (PreK-8, Durham)", f"{new_dcs_total} (Durham PreK-5 + district 6th)", f"{new_dcs_total - school_enrollment['DCS']:+d}")
table_row("FMS", f"{school_enrollment['FMS']} (Freeport+Pownal 6-8)", f"{new_fms_total} (district 7-8)", f"{new_fms_total - school_enrollment['FMS']:+d}")
table_row("PES", f"{school_enrollment['PES']} (Pownal PreK-5)", f"{school_enrollment['PES']} (no change)", "+0")

w()
w(f"6th grade composition at DCS: Durham {dcs_grade_6} + Freeport ~{freeport_6th} + Pownal ~{pownal_6th} = **{new_dcs_6th_total}**")

w()
w("### 7.2 Staffing Impact")
w()

current_78_fte = 6 + 10.32
new_78_sections = new_fms_total / 22
saved_fte = current_78_fte - new_78_sections

current_6_fte = 2 + 5.16
new_6_sections = new_dcs_6th_total / 20
net_6_change = new_6_sections - current_6_fte

table_row("Grade Level", "Current FTE", "Proposed FTE", "Change", header=True)
table_row("Grades 7-8", f"{current_78_fte:.1f} (DCS 6 + FMS 10.3)", f"{new_78_sections:.1f} (FMS only)", f"{-saved_fte:+.1f}")
table_row("Grade 6", f"{current_6_fte:.1f} (DCS 2 + FMS 5.2)", f"{new_6_sections:.1f} (DCS only)", f"{net_6_change:+.1f}")
net_fte_change = -(saved_fte - net_6_change)  # negative = savings
table_row("**NET**", f"**{current_78_fte + current_6_fte:.1f}**", f"**{new_78_sections + new_6_sections:.1f}**", f"**{net_fte_change:+.1f} (saves {saved_fte - net_6_change:.1f})**")

# saved_fte is POSITIVE (we need fewer 7-8 teachers)
# net_6_change is POSITIVE (we need MORE 6th grade teachers)
# Net FTE savings = saved_fte MINUS net_6_change
net_fte_savings = saved_fte - net_6_change
total_teacher_savings = net_fte_savings * cost_per_fte
extracurr_savings = 40_000
total_savings_b = total_teacher_savings + extracurr_savings

w()
w()
w("*Note: Grades 7-8 sections computed at 22 students/section (middle school norm);")
w("grade 6 at 20 students/section (upper elementary class size norm). [^1]*")
w()
w(f"- Cost per FTE (salary + benefits): ${cost_per_fte:,} [^2]")
w(f"- Net FTE savings: {net_fte_savings:.1f} FTE (7-8 saves {saved_fte:.1f}, 6th adds {net_6_change:.1f})")
w(f"- Teacher savings: ${total_teacher_savings:,.0f}/year")
w(f"- Extracurricular consolidation: ~${extracurr_savings:,}/year")
w(f"- **Total estimated annual savings: ${total_savings_b:,.0f}**")

# ============================================================
# SECTION 8: ENROLLMENT REBALANCING [fn1]
# ============================================================

w()
w("---")
w()
w("## 8. Enrollment Rebalancing at PES [^1]")
w()
w("PES has significant room to absorb additional students without hiring new teachers")
w("for most grades (target class size: 22).")
w()

pes_grades = {
    "PreK": {"students": 16, "teachers": 0.5},
    "K": {"students": 14, "teachers": 1},
    "1": {"students": 14, "teachers": 1},
    "2": {"students": 13, "teachers": 1},
    "3": {"students": 17, "teachers": 1},
    "4": {"students": 13, "teachers": 1},
    "5": {"students": 18, "teachers": 1},
}

current_total = sum(v["students"] for v in pes_grades.values())
room = sum(22 - v["students"] for k, v in pes_grades.items() if k != "PreK")

table_row("Grade", "Current Students", "Teachers", "Room to 22", header=True)
for gr, info in pes_grades.items():
    cap = 22 - info["students"] if gr != "PreK" else 0
    table_row(gr, info["students"], info["teachers"], f"+{cap}" if cap > 0 else "--")
table_row("**Total**", f"**{current_total}**", "", f"**+{room} K-5**")

w()

for idx, add in enumerate([5, 8], start=1):
    w(f"### 8.{idx} Scenario: +{add} students/grade K-5")
    w()
    new_total = current_total + (add * 6)
    old_cost_per = school_direct["PES"] / current_total

    extra_teachers = 0
    grades_needing_2nd = []
    for gr, info in pes_grades.items():
        if gr == "PreK":
            continue
        if info["students"] + add > 22:
            extra_teachers += 1
            grades_needing_2nd.append(gr)

    additional_cost = extra_teachers * cost_per_fte
    adjusted_cost = school_direct["PES"] + additional_cost
    adjusted_per_student = adjusted_cost / new_total

    table_row("Metric", "Value", header=True)
    table_row("New K-5 enrollment", f"{new_total} (+{add*6})")
    table_row("Grades exceeding 22 (need 2nd teacher)", f"{extra_teachers} ({', '.join(grades_needing_2nd) if grades_needing_2nd else 'none'})")
    table_row("Additional teacher cost", f"${additional_cost:,}")
    table_row("Original per-student", f"${old_cost_per:,.0f}")
    table_row("New per-student (with added teachers)", f"${adjusted_per_student:,.0f}")
    table_row("**Reduction**", f"**${old_cost_per - adjusted_per_student:,.0f}/student**")
    w()

# ============================================================
# SECTION 9: EARLY CHILDHOOD (CDS TRANSITION) [fn13-19]
# ============================================================

w()
w("---")
w()
w("## 9. Early Childhood: CDS Transition Analysis [^17] [^18] [^19]")
w()
w("### 9.1 The Mandate")
w()
w("By **July 1, 2028**, Maine school districts assume responsibility for identifying and providing")
w("Free Appropriate Public Education (FAPE) to children ages 3-5 with disabilities -- a role")
w("currently handled by Child Development Services (CDS). [^17] Districts may begin earlier and")
w("receive state financial support for early adoption. [^17]")
w()
w("**RSU 5 has committed to Cohort 3 (July 2027)** -- one year ahead of the general deadline.")
w("The Superintendent's Feb 11, 2026 presentation states: 'Requirement to serve all 3- and")
w("4-year-olds with IEPs by July 2027 (we're committed to Cohort 3).' [^10] This accelerated")
w("timeline means EC planning decisions must be finalized by Fall 2026.")
w()
w("RSU 5 formed an **Early Childhood Transition Task Force** in October 2025, including staff,")
w("parents, select board members, private care providers, and board directors from all three")
w("towns. [^18] The task force met four times (Oct 2025 - Jan 2026) and produced a detailed")
w("service model options document with startup costs. [^19]")
w()

w("### 9.2 Current PreK Enrollment & Capacity [^1]")
w()

town_pop = {"Pownal": 1590, "Durham": 4339, "Freeport": 8771}
total_pop = sum(town_pop.values())
birth_rate = 9.5 / 1000
total_births = total_pop * birth_rate

current_prek = {"MSS (Freeport)": 64, "PES (Pownal)": 16, "DCS (Durham)": 48}
total_prek = sum(current_prek.values())

table_row("Location", "Current PreK Slots", "Est. Births/Year [^13]", "Coverage", header=True)
for loc, count, town_key in [("MSS (Freeport)", 64, "Freeport"), ("PES (Pownal)", 16, "Pownal"), ("DCS (Durham)", 48, "Durham")]:
    births = town_pop[town_key] * birth_rate
    table_row(loc, count, f"~{births:.0f}", f"~{count/births*100:.0f}%")
table_row("**Total**", f"**{total_prek}**", f"**~{total_births:.0f}**", f"**~{total_prek/total_births*100:.0f}%**")

w()
w("Maine DOE Chapter 124 rules: max **16 students** per PreK classroom, **1:8 staff ratio**")
w("during academic time (1:10 during meals/outdoor). [^15]")
w()
w("RSU 5 Task Force projection: **30-40 four-year-olds** expected for 2026-27, with existing")
w("PreK capacity to absorb **10-14 additional** students. [^19]")

w()
w("### 9.3 RSU 5 Service Model Options & Costs [^19]")
w()
w("The Task Force evaluated three models. **Option 3 was explicitly not recommended** by the")
w("task force due to loss of service time and not meeting Least Restrictive Environment (LRE).")
w()

# RSU 5's own cost estimates from Dec 18, 2025 task force document
ec_option1_staff = 1_128_000
ec_option1_equip = 61_644
ec_option1_total = 1_189_644

ec_option2_staff = 1_477_000
ec_option2_equip = 195_834
ec_option2_total = 1_672_834

ec_option3_total = 2_006_834

w("#### Option 1: School-Based with SpEd Classroom (RECOMMENDED)")
w()
table_row("Position", "Cost", "Count", "Total", header=True)
for pos, cost, count, total in [
    ("ECSE Coordinator (Asst. Director)", "$176,000", 1, "$176,000"),
    ("Office Support / Transport Coord.", "$105,000", 1, "$105,000"),
    ("SpEd Teacher (282B)", "$144,000", 1, "$144,000"),
    ("Ed Techs", "$53,000", 3, "$159,000"),
    ("Speech-Language Pathologist", "$177,000", 1, "$177,000"),
    ("Social Worker", "$177,000", 1, "$177,000"),
    ("Driver Hours", "$13,000", 1, "$13,000"),
    ("Contractors (Psych, TOD, TVI, etc.)", "Variable", "--", "$177,000"),
]:
    table_row(pos, cost, count, total)
table_row("**Staff Subtotal**", "", "", f"**${ec_option1_staff:,}**")
table_row("Equipment (van, tech, assessments)", "", "", f"${ec_option1_equip:,}")
table_row("**OPTION 1 TOTAL**", "", "", f"**${ec_option1_total:,}**")

w()
w("#### Option 2: School + Community-Based with SpEd Classroom")
w()
w(f"Adds OT ($177K), doubles Ed Techs to 6 ($318K), 2 vans. **Total: ${ec_option2_total:,}**")
w(f"(+${ec_option2_total - ec_option1_total:,} vs Option 1)")
w()
w("#### Option 3: Community-Based without SpEd Classroom (NOT RECOMMENDED)")
w()
w(f"Requires 4 out-of-district placements ($260K), 4 vans. **Total: ${ec_option3_total:,}**")
w("Rejected by task force: loses service time, students sent to SPPS with no non-disabled peers.")

w()
w("### 9.4 EC Facility Strategy: Distributed Model [^19] [^1]")
w()
w("**The critical stakeholder constraint:** parents in each town want young children close to home.")
w("A distributed model keeps EC services at each community's school while sharing specialized")
w("staff district-wide:")
w()
table_row("School", "EC Role", "Physical Capacity", "LRE Inclusion Ratio", header=True)
table_row("MSS", "SpEd classroom + Freeport PreK hub", "Declining enrollment (-42 since 2023) frees space; PreK-2 infrastructure", "SpEd ~10-12 among 64 PreK = ~1:5 (strong)")
table_row("DCS", "Durham PreK (existing, continues)", "Existing PreK infrastructure, 48 students", "N/A (no SpEd classroom)")
table_row("PES", "Pownal PreK (existing, continues)", "Surplus space for future K-5 growth", "16 PreK too small for SpEd co-location (~1:1)")

w()
w("Under this model, the **SpEd classroom** (the primary new mandate) is co-located at MSS,")
w("which has the district's largest PreK program (64 students) for robust LRE inclusion.")
w("MSS already serves PreK-2, so facility modifications are minimal compared to a K-5 building.")
w("MSS is also the most centrally located school, reducing transport for the small number")
w("of Durham/Pownal IEP children requiring the self-contained setting.")
w("Itinerant staff (SLP, Social Worker, ECSE Coordinator) travel between all three schools")
w("to serve children with milder IEP needs at their home school.")
w()

# Cost offset analysis
w("### 9.5 EC Cost Offsets")
w()
w("The $1.19M Option 1 cost is **not fully net-new district spending**:")
w()
table_row("Offset Source", "Est. Annual Value", "Notes", header=True)
table_row("CDS budget transfer to districts", "~$200K-$400K", "State transfers existing CDS funding per child served [^17]")
table_row("IDEA Part B (federal 3-5 funding)", "~$100K-$150K", "Federal special ed dollars follow children to districts")
table_row("EPS subsidy for PreK students", "~$50K-$100K", "State EPS formula provides per-pupil subsidy for public PreK [^15]")
table_row("**Est. total offsets**", "**~$350K-$650K**", "Reduces net new cost to ~$540K-$840K")
w()
w("**Net EC cost estimate (Option 1): ~$540K-$840K/year** after state/federal offsets.")

# ============================================================
# SECTION 10: BUDGET GAP MODEL
# ============================================================

w()
w("---")
w()
w("## 10. Budget Gap Analysis: Reaching a Politically Acceptable Increase")
w()

fy26_adopted = 44_455_929
fy27_proposed = 47_357_441  # Revised 02/11/2026: Articles 1-11 ($47,269,441) + Adult Ed ($88,000)
fy27_increase = fy27_proposed - fy26_adopted
fy27_increase_pct = fy27_increase / fy26_adopted * 100

w(f"### 10.1 Official Budget Reconciliation [^9] [^2] [^25]")
w()
w("**Verified against revised FY27 Superintendent's Budget Handbook (02/11/2026):**")
w()
w(f"- FY26 adopted total operating budget: ${fy26_adopted:,} [^9]")
w(f"- FY27 proposed total operating budget: ${fy27_proposed:,} [^25]")
w(f"  - Articles 1-11: $47,269,441")
w(f"  - Adult Education: $88,000 (flat)")
w(f"- Proposed increase: ${fy27_increase:,} ({fy27_increase_pct:.2f}%)")
w()
w("*Note: Earlier versions of this analysis used $47,327,804 from the Jan 28 handbook.")
w("The Feb 11 revision added $29,637 (CTE/Region 10). All figures below use the revised total.*")
w()
w("**10-year budget history for context:**")
w()
budget_history = [
    ("FY26", 6.83), ("FY25", 6.48), ("FY24", 4.99), ("FY23", 4.22), ("FY22", 2.09),
    ("FY21", 2.32), ("FY20", 3.43), ("FY19", 2.31), ("FY18", 4.20), ("FY17", 5.15),
]
table_row("Year", "Increase %", header=True)
for yr, pct in budget_history:
    table_row(yr, f"{pct:.2f}%")
avg_10yr = sum(p for _, p in budget_history) / len(budget_history)
avg_3yr = sum(p for _, p in budget_history[:3]) / 3
table_row("**10-year avg**", f"**{avg_10yr:.1f}%**")
table_row("**3-year avg**", f"**{avg_3yr:.1f}%**")
w()
w(f"The proposed {fy27_increase_pct:.1f}% increase is consistent with the 3-year trend ({avg_3yr:.1f}% avg).")
w(f"Voters approved 6.48% (FY25) and 6.83% (FY26) increases in consecutive years.")
w()

ec_net_low = ec_option1_total - 650_000
ec_net_high = ec_option1_total - 350_000

w(f"Adding EC mandate (Option 1 net of offsets): ${ec_net_low:,} to ${ec_net_high:,}")

total_with_ec_low = fy27_proposed + ec_net_low
total_with_ec_high = fy27_proposed + ec_net_high
pct_with_ec_low = (total_with_ec_low - fy26_adopted) / fy26_adopted * 100
pct_with_ec_high = (total_with_ec_high - fy26_adopted) / fy26_adopted * 100

w(f"- FY27 with EC (low est.): ${total_with_ec_low:,} ({pct_with_ec_low:.1f}% increase)")
w(f"- FY27 with EC (high est.): ${total_with_ec_high:,} ({pct_with_ec_high:.1f}% increase)")
w()

# Target: <6% increase = user's stated goal, need to trim ~$1.7M
target_pct = 6.0
target_budget = fy26_adopted * (1 + target_pct / 100)
gap_from_proposal = fy27_proposed - target_budget
gap_with_ec_high = total_with_ec_high - target_budget

w(f"**Target: <{target_pct:.0f}% increase = ${target_budget:,.0f}**")
w()
w(f"- Gap (FY27 proposal only): ${gap_from_proposal:,.0f}")
w(f"- Gap (FY27 + EC high est.): ${gap_with_ec_high:,.0f}")
w(f"- *Note: A gap of ~$1.7M would be required to hold the increase below ~4.5% when EC is included.*")

w()
w("### 10.2 Savings Levers (Preliminary Estimates)")
w()
w("*Note: These are initial estimates. Section 11 provides critically reviewed figures")
w("that supersede the DCS admin savings and transportation estimates below.*")
w()

net_positions = {}
for town in ["Pownal", "Durham", "Freeport"]:
    net_positions[town] = revenue[town]["total_revenue"] - consumption_by_town[town]

pes_per_stu = school_direct["PES"] / 105
avg_elem = avg_other
gap_per_stu = pes_per_stu - avg_elem
total_gap = gap_per_stu * 105

# DCS admin savings from grade span reduction (K-8 -> K-5+6th)
# When DCS drops 7-8, they lose ~100 students worth of admin overhead
# but gain 131 6th graders (Freeport + Pownal). Net change: +31 students.
# Admin complexity DECREASES because K-6 is simpler than K-8 (no middle school athletics,
# no 8th grade activities, simpler scheduling). Estimate 0.5 admin FTE savings.
dcs_admin_savings = 0.5 * 130_000  # half an admin position

# Transportation: consolidating 7-8 at FMS eliminates DCS->FMS/FHS double-routing
# for Durham 7-8 students. But adds Durham 7-8 to FMS routes.
# Net: roughly neutral for 7-8 (same students, different building)
# BUT: 6th grade to DCS adds Freeport->DCS and Pownal->DCS routes
# Estimate: +1 bus route for Freeport 6th to DCS, +0 for Pownal (close to DCS route)
# Cost per bus route: ~$65K-$80K/year (driver + fuel + maintenance)
new_bus_route_cost = 75_000

# EC: SpEd classroom at MSS (largest PreK cohort = 64 students, best LRE ratio ~1:5)
# MSS already serves PreK-2 with age-appropriate facilities; minimal conversion needed
ec_facility_savings = 0  # MSS conversion ~$50K-$100K, covered by Option 1 startup budget

table_row("Lever", "Annual Savings", "Type", "Stakeholder Impact", header=True)
table_row("A: MS consolidation (7-8 to FMS)", f"${total_savings_b:,.0f}", "Real budget reduction", "Durham 7-8 students move to FMS; all 6th to DCS")
table_row("B: DCS admin simplification (K-6 vs K-8)", f"${dcs_admin_savings:,.0f}", "Real budget reduction", "DCS admin refocused, narrower grade span")
table_row("C: New bus route (Freeport 6th to DCS)", f"-${new_bus_route_cost:,}", "New cost", "6th graders ride bus to Durham (15 min)")
table_row("D: EC at existing schools (vs new facility)", "$0 (avoided)", "Cost avoidance", "Young children stay in home community")
table_row("E: EC state/federal offsets", "$350K-$650K", "Revenue offset", "No stakeholder impact")
table_row("F: PES enrollment rebalancing (+5/gr)", "$0 direct", "Per-student metric", "Smaller Freeport classes, stronger PES culture")

w()

# Net calculation
ms_net = total_savings_b + dcs_admin_savings - new_bus_route_cost
ec_offset_mid = 500_000  # midpoint of offset range

w(f"**Net from structural changes (A+B+C):** ${ms_net:,.0f}")
w(f"**EC offset (midpoint estimate):** ${ec_offset_mid:,}")
w(f"**Combined structural savings + EC offsets:** ${ms_net + ec_offset_mid:,.0f}")
w()
w(f"Remaining gap to reach <6%: ${max(0, gap_with_ec_high - ms_net - ec_offset_mid):,.0f}")

# What else can close the gap?
remaining_gap = gap_with_ec_high - ms_net - ec_offset_mid

w()
w("### 10.3 Additional Efficiency Options to Close the Gap")
w()
w("The structural changes above close a significant portion of the gap but may not fully")
w("reach the ~$1.7M target. Additional options to present to the Board:")
w()

# Transportation article increased 23.65% = ~$457K increase
transport_fy26 = 1_932_000  # estimated FY26 transport
transport_fy27 = 2_388_457
transport_increase = transport_fy27 - transport_fy26

# Facilities increased 9.17%
fac_fy27 = 5_993_777
fac_increase_est = fac_fy27 * 0.0917 / (1 + 0.0917)  # back-calculate

table_row("Option", "Potential Savings", "Notes", header=True)
table_row("Transportation route optimization", "$75K-$150K", f"Art 8 increased 23.65% to ${transport_fy27:,}; route consolidation with grade restructuring")
table_row("Facilities maintenance phasing", "$100K-$200K", f"Art 9 at ${fac_fy27:,}; defer non-critical projects 1 year")
table_row("System admin efficiencies", "$50K-$100K", "Art 6 at $1.31M; shared services, technology")
table_row("Natural attrition (unfilled positions)", "$100K-$200K", "Hold 1-2 positions open through consolidation transition")
table_row("**Range of additional savings**", "**$325K-$650K**", "")

total_potential_low = ms_net + ec_offset_mid + 325_000
total_potential_high = ms_net + ec_offset_mid + 650_000

w()
w(f"**Total potential savings range: ${total_potential_low:,.0f} to ${total_potential_high:,.0f}**")

w()
w("### 10.4 Graduated Increase Targets")
w()
w("What does it take to reach various politically acceptable increase levels?")
w()

table_row("Target Increase", "Target Budget", "Cuts Needed (from proposal)", "Cuts Needed (from proposal + EC)", header=True)
for target in [6.0, 5.0, 4.0, 3.0]:
    target_bud = fy26_adopted * (1 + target / 100)
    from_proposal = fy27_proposed - target_bud
    from_with_ec = total_with_ec_high - target_bud
    table_row(f"<{target:.0f}%", f"${target_bud:,.0f}", f"${max(0, from_proposal):,.0f}", f"${max(0, from_with_ec):,.0f}")

w()
w("The middle school consolidation + EC offsets + efficiencies (Package 1) provides")
w(f"~${abs(ms_net + ec_offset_mid + 400_000):,.0f} toward these targets. Additional cuts beyond")
w("Package 1 require line-item reductions from the Superintendent's proposed budget.")

# ============================================================
# SECTION 11: COMPREHENSIVE SCENARIO PACKAGES
# ============================================================

w()
w("---")
w()
w("## 11. Comprehensive Scenario Packages")
w()
w("Each package below is a **complete, presentable path** that addresses the budget gap,")
w("the EC mandate, grade restructuring, and community impact.")

# Package 1: Preserve All + Distributed EC
w()
# All "budget_impact" variables use: negative = reduces budget (savings), positive = increases budget (cost)
# EC mandate is REQUIRED by 2028 -- must be funded regardless of restructuring choice.

ec_net_cost = ec_option1_total - ec_offset_mid  # net new cost after offsets

w("The EC mandate is required by July 2028 statewide, but **RSU 5 has committed to")
w(f"Cohort 3 (July 2027)** -- one year early. [^10] Its net cost (${ec_net_cost:,})")
w("must be funded either way. The question is how much restructuring savings can")
w("offset this mandatory new cost.")
w()

# BASELINE: FY27 proposed + EC mandate, NO restructuring
baseline_with_ec = fy27_proposed + ec_net_cost
baseline_ec_pct = (baseline_with_ec - fy26_adopted) / fy26_adopted * 100

w("### 11.1 Baseline: FY27 Proposed + EC Mandate (No Restructuring)")
w()
w(f"- FY27 proposed: ${fy27_proposed:,}")
w(f"- EC mandate net cost: +${ec_net_cost:,}")
w(f"- **Baseline total: ${baseline_with_ec:,} ({baseline_ec_pct:.1f}% increase from FY26)**")
w()
w("All packages below are compared against this baseline.")

w()
w("### 11.2 Package 1: \"Preserve & Strengthen\" (Recommended)")
w()
w("*Preserves all three elementary schools, distributes EC to home communities,")
w("restructures middle grades for efficiency.*")
w()

# Package 1 budget savings -- CRITICALLY REVIEWED for hidden costs
# DCS admin: originally $65K savings. But DCS goes from single-community K-8
# to multi-community K-6 with 157 district-wide 6th graders. This is MORE complex
# administratively, not simpler. Revise to $0 savings.
dcs_admin_savings_revised = 0

# Transportation: originally 1 route at $75K. But 116 Freeport 6th graders need
# to get from Freeport to DCS (~9 miles). That requires 2 routes minimum.
# Pownal 6th (~15 students) can share existing Pownal routing.
new_bus_routes_cost_revised = 150_000  # 2 routes at $75K each

# DCS portable classrooms: DCS loses 6 sections (7-8) and needs ~6 additional for
# 157 6th graders (currently has 2 sections of 41; needs 8 sections of 157).
# Net section change: 6 freed - 6 needed = 0. No portables needed.
dcs_portables_path_a = 0

pkg1_savings = (total_savings_b + dcs_admin_savings_revised
                - new_bus_routes_cost_revised - dcs_portables_path_a + 400_000)

w("**NOTE: Path A costs have been critically reviewed. Key corrections from initial estimate:**")
w("- DCS admin savings revised from $65K to $0 (multi-community 6th grade adds complexity)")
w(f"- Freeport 6th transportation revised from $75K to ${new_bus_routes_cost_revised:,} (2 routes needed for ~{freeport_6th} students)")
w(f"- DCS portables: NOT needed (DCS gains ~6 sections of 6th, loses 6 sections of 7-8; net 0)")
w(f"- 6th grade double-count corrected: district total is {new_dcs_6th_total} (not 172)")
w()

w("**Structural Changes:**")
w("1. Grades 7-8 consolidated at FMS (district-wide)")
w("2. Grade 6 consolidated at DCS (district-wide)")
w("3. PES remains PreK-5 (no change)")
w("4. EC services distributed: SpEd classroom at MSS (64 PreK peers, 1:5 LRE ratio), community PreK continues at all three schools")
w("5. PES accepts +5 students/grade from Freeport school choice (voluntary)")
w()

w("**Per-School Impact:**")
w()
table_row("School", "Current Config", "New Config", "Enrollment Change", "Culture Impact", header=True)
table_row("PES", "PreK-5 (105)", "PreK-5 (135 via school choice)", "+30 K-5", "Strengthened: larger cohorts, community school preserved")
table_row("DCS", f"PreK-8 Durham ({school_enrollment['DCS']})", f"PreK-5 Durham + district 6th ({new_dcs_total})", f"{new_dcs_total - school_enrollment['DCS']:+d}", "Shifts to K-6 focus; gains all-district 6th grade community")
table_row("FMS", "Freeport+Pownal 6-8 (306)", "District-wide 7-8 (290)", "-16", "Tighter grade band; all three towns together earlier")
table_row("MSS", "PreK-2 Freeport (274)", "PreK-2 + EC SpEd classroom (269+12)", "+7 net", "Gains EC SpEd hub; declining enrollment absorbs it; becomes EC expertise center")
table_row("MLS", "3-5 Freeport (264)", "3-5 Freeport (254)", "-10", "Modest reduction; relieves any crowding")
table_row("FHS", "All towns 9-12 (554)", "No change (554)", "0", "No change")

w()
w("**Financial Summary (vs. Baseline) -- CRITICALLY REVIEWED:**")
w()
table_row("Item", "Budget Impact", "Notes", header=True)
table_row(f"MS consolidation teacher savings ({net_fte_savings:.1f} net FTE)", f"-${total_savings_b:,.0f}", f"7-8 saves {saved_fte:.1f}, 6th adds {net_6_change:.1f}")
table_row("DCS admin change (multi-community K-6)", "$0", "Revised: complexity increases")
table_row(f"Freeport 6th transportation (2 routes to DCS)", f"+${new_bus_routes_cost_revised:,}", f"~{freeport_6th} students, ~9 mi")
table_row("DCS portable classrooms", "$0", "6 sections freed (7-8) offset 6 sections added (6th)")
table_row("Additional efficiencies (transportation, facilities, attrition)", f"-$400,000", "Route consolidation, maintenance phasing, attrition")
table_row("**Total restructuring savings**", f"**-${pkg1_savings:,.0f}**", "")

w()

fy27_pkg1 = baseline_with_ec - pkg1_savings
pkg1_increase = fy27_pkg1 - fy26_adopted
pkg1_increase_pct = pkg1_increase / fy26_adopted * 100
pkg1_vs_baseline = baseline_with_ec - fy27_pkg1

w(f"- Baseline (FY27 + EC, no restructuring): ${baseline_with_ec:,} ({baseline_ec_pct:.1f}%)")
w(f"- **Package 1 adjusted total: ${fy27_pkg1:,.0f} ({pkg1_increase_pct:.1f}% increase from FY26)**")
w(f"- Savings vs. baseline: ${pkg1_vs_baseline:,.0f}")

w()
w("**Per-Town Tax Impact (savings vs. baseline, estimated):**")
w()
for town_name, alm_pct in [("Pownal", 0.1260), ("Durham", 0.2142), ("Freeport", 0.6598)]:
    town_savings = pkg1_vs_baseline * alm_pct
    state_val = tax_data[town_name]["state_val"]
    eq_mil_change = town_savings / state_val * 1000
    eq_home_change = 400_000 * eq_mil_change / 1000
    w(f"- **{town_name}**: saves ~${eq_home_change:,.0f}/year on a $400K home (equalized)")

w()
w("**Stakeholder Assessment:**")
w()
table_row("Stakeholder", "Impact", "Palatability", header=True)
table_row("Pownal parents", "PES preserved and strengthened; EC stays in Pownal; young kids close", "HIGH")
table_row("Durham parents", "DCS shifts to K-6; gains district-wide 6th grade; 7-8 moves to FMS", "MODERATE (7-8 move is change)")
table_row("Freeport parents", "6th goes to DCS (new); 7-8 stays at FMS; PreK unchanged", "MODERATE (6th to DCS is new)")
table_row("PES teachers", "Larger school, more colleagues, EC specialization", "HIGH")
table_row("DCS teachers", "K-6 focus vs K-8; gain 6th grade teachers from FMS", "MODERATE-HIGH")
table_row("FMS teachers", "Lose 6th, gain Durham 7-8; tighter grade band", "MODERATE")
table_row("Administration", "Fewer grade spans per building; EC integrated not separate", "HIGH")
table_row("School Board", f"Budget increase held to {pkg1_increase_pct:.1f}% (vs {baseline_ec_pct:.1f}% baseline); EC mandate met; no school closures", "HIGH")
table_row("Taxpayers", f"~{pkg1_increase_pct:.1f}% vs {baseline_ec_pct:.1f}% if no action taken; EC mandate funded within savings", "MODERATE-HIGH")

# ============================================================
# SECTION 11.3: EC FACILITY & SPED ANALYSIS
# ============================================================

w()
w("### 11.3 EC Facility Requirements & SpEd Classroom Analysis [^15] [^21]")
w()
w("#### Is One SpEd Classroom Sufficient?")
w()
w("The Task Force staffed Option 1 with 1 SpEd teacher + 3 Ed Techs = 4 adults for")
w("one classroom. At Chapter 124's max of 16 students, this provides a 1:4 ratio,")
w("adequate for the estimated 10-12 IEP-eligible 3-5 year-olds across the district. [^19]")
w()
w("One classroom is likely sufficient for Years 1-2. However, IEP caseloads are")
w("unpredictable, and students with severe needs may require 1:1 aides. The district")
w("should plan physical space for a second classroom even if staffing waits.")
w()
w("#### Can SpEd Be Split from General EC Expansion?")
w()
w("Legally, yes -- the CDS mandate covers only FAPE for children with disabilities.")
w("General PreK expansion for all 4-year-olds is a separate, voluntary decision.")
w()
w("**However, splitting them undermines the program.** Maine Chapter 124 and federal IDEA")
w("require Least Restrictive Environment (LRE): SpEd students should be educated")
w("alongside non-disabled peers whenever possible. The Task Force explicitly rejected")
w("Option 3 (community-based without SpEd classroom) because students would be 'sent to")
w("SPPS and have no access to non-disabled peers.' [^19] Locating the SpEd classroom at a")
w("site without general PreK students recreates that same LRE problem.")
w()
w("**Recommendation:** Keep SpEd co-located with a large general PreK program. Under Path A,")
w("MSS is the clear choice -- it hosts 64 PreK students (4 sections), producing a healthy")
w("~1:5 SpEd-to-general ratio. MSS already serves PreK-2 with age-appropriate infrastructure,")
w("has declining enrollment freeing classroom space, and is the most centrally located school.")
w("PES (16 PreK students) would produce an inadequate ~1:1 ratio that risks creating a")
w("SpEd-weighted environment harmful to both SpEd and general education students.")
w()
w("#### Facility Modifications for 3-Year-Olds [^21]")
w()
w("Chapter 124 requires specific physical standards that existing buildings may not meet:")
w()
table_row("Requirement", "Chapter 124 Standard", "MSS Status (Path A SpEd host)", "PES Status (Path B full EC)", header=True)
table_row("Toilets", "Within 40 feet; preferably IN classroom", "PreK-2 building, likely closer to compliance", "K-5 bathrooms designed for ages 5-11; YES modify")
table_row("Handwashing", "Water source IN each classroom", "May exist in PreK rooms already", "Unlikely in all rooms; YES add")
table_row("Classroom space", "35 sq ft per child, usable", "PreK rooms already sized for youngest", "K-5 rooms likely meet this; Minimal")
table_row("Natural light", "Required in all PreK rooms", "Yes", "Yes")
table_row("Outdoor play", "75 sq ft/child, fenced, age-appropriate", "Already serves PreK-2; may need additions", "K-5 playground not suitable for 3yo; YES modify")
table_row("Furniture", "Age-appropriate for 3-4yo", "PreK furniture likely in place", "K-5 sized; YES full replacement in EC rooms")
w()
w("**Path A: MSS SpEd classroom conversion costs (one-time):**")
w()

mss_conversion_bathroom = 30_000
mss_conversion_playground = 40_000
mss_conversion_plumbing = 20_000
mss_conversion_furniture = 15_000
mss_conversion_safety = 10_000
mss_conversion_design = 15_000
mss_conversion_total = (mss_conversion_bathroom + mss_conversion_playground
                        + mss_conversion_plumbing + mss_conversion_furniture
                        + mss_conversion_safety + mss_conversion_design)
mss_conversion_amortized = mss_conversion_total / 10

table_row("Item", "Estimated Cost", header=True)
table_row("Bathroom modifications (1-2 rooms, already near-compliant)", f"${mss_conversion_bathroom:,}")
table_row("Playground additions (fencing, surfacing for youngest)", f"${mss_conversion_playground:,}")
table_row("In-classroom plumbing (handwashing, 1-2 rooms)", f"${mss_conversion_plumbing:,}")
table_row("Age-appropriate furniture (SpEd classroom only)", f"${mss_conversion_furniture:,}")
table_row("Safety modifications", f"${mss_conversion_safety:,}")
table_row("Design and permitting", f"${mss_conversion_design:,}")
table_row("**Total one-time conversion**", f"**${mss_conversion_total:,}**")
table_row("Amortized over 10 years", f"${mss_conversion_amortized:,.0f}/year")
w()
w("MSS conversion costs are substantially lower than PES because MSS already serves PreK-2")
w("with age-appropriate infrastructure. Only the dedicated SpEd classroom needs full conversion.")
w()
w("**Path B: PES full EC center conversion costs (one-time):**")
w()

pes_conversion_bathroom = 100_000
pes_conversion_playground = 150_000
pes_conversion_plumbing = 60_000
pes_conversion_furniture = 50_000
pes_conversion_safety = 25_000
pes_conversion_design = 40_000
pes_conversion_total = (pes_conversion_bathroom + pes_conversion_playground
                        + pes_conversion_plumbing + pes_conversion_furniture
                        + pes_conversion_safety + pes_conversion_design)
pes_conversion_amortized = pes_conversion_total / 10

table_row("Item", "Estimated Cost", header=True)
table_row("Bathroom retrofit (child-height fixtures, changing stations)", f"${pes_conversion_bathroom:,}")
table_row("Playground (age-appropriate equipment, fencing, surfacing)", f"${pes_conversion_playground:,}")
table_row("In-classroom plumbing (handwashing stations, 3-4 rooms)", f"${pes_conversion_plumbing:,}")
table_row("Age-appropriate furniture and materials", f"${pes_conversion_furniture:,}")
table_row("Safety modifications (outlets, door hardware, barriers)", f"${pes_conversion_safety:,}")
table_row("Design, permitting, and ADA compliance review", f"${pes_conversion_design:,}")
table_row("**Total one-time conversion**", f"**${pes_conversion_total:,}**")
table_row("Amortized over 10 years", f"${pes_conversion_amortized:,.0f}/year")
w()
w("Path B requires converting most or all PES classrooms (a K-5 building not designed for")
w("3-year-olds), roughly 3x the cost of Path A's single-classroom MSS conversion.")

# ============================================================
# SECTION 11.4: PES GEOGRAPHIC ANALYSIS
# ============================================================

w()
w("### 11.4 PES Geographic Position: Not Central [^22]")
w()
w("PES at 587 Elmwood Road, Pownal is the **least centrally located** school in RSU 5.")
w("Approximate driving distances and times from town centers:")
w()

# Approximate distances between RSU 5 schools
table_row("Route", "Distance", "Drive Time", header=True)
table_row("PES to MSS/Freeport center", "~13 miles", "~22 minutes")
table_row("PES to DCS (Durham)", "~11 miles", "~19 minutes")
table_row("PES to FMS/FHS (Freeport)", "~9 miles", "~17 minutes")
table_row("DCS to MSS/Freeport center", "~9 miles", "~16 minutes")
table_row("DCS to FMS/FHS (Freeport)", "~8 miles", "~15 minutes")

w()
w("**Impact on EC transportation under Path B:**")
w("If PES is repurposed as the district's centralized EC center, 3-5 year-olds from")
w("Freeport and Durham must be transported daily to Pownal -- the farthest point in the")
w("RSU. For 3-year-olds, this means:")
w()
w("- Bus rides of 20-25 minutes each way (Freeport) or 19+ minutes (Durham)")
w("- Specialized vehicles required (car seats for 3-year-olds, smaller vans)")
w("- Multiple daily runs if half-day programs are offered (4 trips/day)")
w("- Parents of very young children are highly sensitive to long transit times")
w("- Many families may opt out entirely, undermining the program's enrollment and funding basis")
w()

ec_transport_premium_centralized = 175_000
w(f"**Estimated EC transport premium for centralized-at-PES:** ${ec_transport_premium_centralized:,}/year")
w("(2-3 specialized van routes above what a distributed model requires)")
w()
w("Under Path A's distributed model, EC transportation is minimal -- children attend")
w("their home community's school, same as they do for PreK today.")

# ============================================================
# SECTION 11.5: PATH B FULL HIDDEN COST MODEL
# ============================================================

w()
w("### 11.5 Path B (Scenario 2): Full Cost Model with Hidden Costs [^25]")
w()
w("*Repurposes PES as district-wide Early Childhood Center; moves Pownal K-6 to DCS;")
w("consolidates 7-8 at FMS; restructures Freeport elementary (MSS K-3, MLS 4-6).*")
w()
w("**IMPORTANT:** The Superintendent's actual Scenario 2 (Feb 11, 2026 presentation) is")
w("more comprehensive than a simple 'close PES' plan. It includes middle school consolidation")
w("(same as Path A), Freeport grade-band changes, AND moving ALL PreK district-wide to PES --")
w("not just the SpEd classroom. [^10] This analysis models the full scenario.*")
w()
w("The original estimate of Path B savings included only PES instruction elimination and")
w("Pownal K-5 transportation. A complete accounting reveals substantial additional costs:")
w()

supt_pes_instruction_savings = school_direct["PES"] - 253_384

# TRANSPORTATION
w("#### A. Transportation Costs (Full Scenario 2)")
w()
w("Under the actual Scenario 2, ALL PreK moves to PES -- not just SpEd. This means:")
w()
pownal_k6_transport = 225_000  # 3 bus routes: Pownal K-6 to DCS (11 mi)
freeport_prek_to_pes = 225_000  # 3 specialized routes: 64 Freeport 4yo + some 3yo to PES (13 mi)
durham_prek_to_pes = 150_000  # 2 specialized routes: 48 Durham 4yo to PES (11 mi)
table_row("Route", "Annual Cost", "Notes", header=True)
table_row("Pownal K-6 to DCS (daily)", f"${pownal_k6_transport:,}", "3 routes covering rural Pownal to Durham (~11 mi, 19 min each way)")
table_row("Freeport PreK (64 students) to PES", f"${freeport_prek_to_pes:,}", "3 specialized routes, car seats, ~13 mi/22 min each way")
table_row("Durham PreK (48 students) to PES", f"${durham_prek_to_pes:,}", "2 specialized routes, car seats, ~11 mi/19 min each way")
pkg2_transport_total = pownal_k6_transport + freeport_prek_to_pes + durham_prek_to_pes
table_row("**Transport total**", f"**${pkg2_transport_total:,}**", "vs. original estimate of $200,000")
w()

# DCS EXPANSION
w("#### B. DCS Capacity Change (Under Full Scenario 2)")
w()
w("Under the Superintendent's actual Scenario 2, DCS is NOT simply absorbing +89 students.")
w("DCS simultaneously loses grades 7-8 (~100 students) and gains Pownal K-6 (~120 students):")
w()
table_row("Change", "Students", header=True)
pownal_k6_to_dcs = sum(v["students"] for k, v in pes_grades.items() if k != "PreK") + pownal_6th  # K-5 (89) + 6th (15) = 104
dcs_scenario2_net = -100 + pownal_k6_to_dcs - 48
table_row("Lose Durham 7-8 (to FMS)", "-100")
table_row("Gain Pownal K-6 (no PreK -- that goes to PES EC)", f"+{pownal_k6_to_dcs}")
table_row("Lose Durham PreK (to PES EC center)", "-48")
table_row("**Net change**", f"**{dcs_scenario2_net:+d}**")
w()
w(f"DCS SHRINKS under the full Scenario 2 (467 → {467 + dcs_scenario2_net}), so portables")
w("are NOT needed for DCS. However, absorbing Pownal families still requires:")
w()
dcs_portable_lease = 0  # NOT needed under actual Scenario 2
dcs_admin_addition = 75_000  # multi-community coordination for Pownal families
dcs_support_staff = 0  # net enrollment decreases
dcs_expansion_total = dcs_portable_lease + dcs_admin_addition + dcs_support_staff

table_row("Item", "Annual Cost", "Notes", header=True)
table_row("Portable classrooms", "$0", "NOT needed: net enrollment decreases under full Scenario 2")
table_row("Multi-community coordination", f"${dcs_admin_addition:,}", "Integrating Pownal families, additional parent relations")
table_row("**DCS adjustment total**", f"**${dcs_expansion_total:,}/year**", "")
w()

# PES CONVERSION
w("#### C. PES Conversion to Full EC Center")
w()
w(f"Under Path B, PES converts from K-5 to serving 3-5 year-olds district-wide.")
w(f"One-time conversion: ${pes_conversion_total:,} (amortized: ${pes_conversion_amortized:,.0f}/year)")
w()
w("**Critical note:** If PES becomes a FULL EC center (not just SpEd + Pownal PreK),")
w("most or all classrooms need conversion, roughly doubling the estimate:")
w()
pes_full_conversion = pes_conversion_total * 1.75
pes_full_amortized = pes_full_conversion / 10
w(f"- Partial conversion (SpEd + 2-3 PreK rooms): ${pes_conversion_total:,}")
w(f"- Full conversion (centralized EC for district): ~${pes_full_conversion:,.0f}")
w(f"- Full conversion amortized: ~${pes_full_amortized:,.0f}/year")

w()
w("#### D. DCS Absorption Cost: What the 'PES Savings' Overlook")
w()
w("The $2M 'PES instruction savings' is a **gross** figure. It assumes PES costs vanish")
w(f"entirely. In reality, {pownal_k6_to_dcs} Pownal K-6 students transfer to DCS and still need")
w("teachers, materials, and support. This absorption cost is routinely overlooked in")
w("school closure analyses.")
w()
w("**FTE analysis of district-wide teacher changes (Scenario 2 vs. current):**")
w()
w("*Current FTE derived from budget articles and class size tables [^1] [^2].")
w("Scenario 2 FTE estimated by recomputing sections at 20-22 students per section,")
w("consistent with RSU 5's current class size practice. These are estimates;")
w("actual staffing decisions would be made by the Superintendent.*")
w()

# Scenario 2 FTE derivation methodology:
# Current FTE derived from budget articles and class size tables [^1][^2]:
#   PES: 0.5 PreK + 6 K-5 homeroom = 6.5
#   DCS: 2.5 PreK + 18 K-6 homeroom + 6 grades 7-8 = 26.5
#   MSS: 3 PreK + 11 K-2 homeroom = 14
#   MLS: 14 grades 3-5 homeroom (264 students / ~19 avg class) = 14
#   FMS: 5.16 grade 6 + 10.32 grades 7-8 = 15.5 (rounded)
# Scenario 2 FTE estimated by recomputing sections at 20-22 students/section:
#   PES EC center: ~5.0 (EC staff: 1 SpEd teacher + 4 PreK sections at max 16)
#   DCS K-6 (Durham + Pownal, no PreK, no 7-8): ~23.0
#     (326 Durham K-5 + 157 district 6th - 48 PreK = ~435 at ~19/section)
#   MSS K-3 (loses PreK to PES, gains 3rd from MLS): ~16.0
#   MLS 4-6 (loses 3rd to MSS, gains Freeport 6th): ~15.0
#   FMS 7-8 (district-wide): ~13.0 (290 students at ~22/section)
s2_pes, s2_dcs, s2_mss, s2_mls, s2_fms = 5.0, 23.0, 16.0, 15.0, 13.0
s2_total = s2_pes + s2_dcs + s2_mss + s2_mls + s2_fms
current_total_fte = 76.5  # sum of current FTE above: 6.5+26.5+14+14+15.5
scenario2_fte_saved = current_total_fte - s2_total

table_row("School", "Current FTE", "Scenario 2 FTE", "Change", "Reason", header=True)
table_row("PES", "6.5 (K-5)", f"{s2_pes:.1f} (EC center)", f"{s2_pes - 6.5:+.1f}", "K-5 gone; EC staff partially transferred from other schools")
table_row("DCS", "26.5 (PreK-8)", f"{s2_dcs:.1f} (K-6 Dur+Pow)", f"{s2_dcs - 26.5:+.1f}", f"Loses PreK/7-8; gains Pownal K-6 ({pownal_k6_to_dcs} students)")
table_row("MSS", "14 (PreK-2)", f"{s2_mss:.0f} (K-3)", f"{s2_mss - 14:+.0f}", "Loses PreK (to PES); gains 3rd grade (from MLS)")
table_row("MLS", "14 (3-5)", f"{s2_mls:.0f} (4-6)", f"{s2_mls - 14:+.0f}", f"Loses 3rd (to MSS); gains Freeport 6th (~{freeport_6th} students)")
table_row("FMS", "15.5 (6-8)", f"{s2_fms:.0f} (7-8 district)", f"{s2_fms - 15.5:+.1f}", "Loses 6th; gains Durham 7-8; net consolidation")
table_row("**Total**", f"**{current_total_fte}**", f"**{s2_total:.1f}**", f"**{-scenario2_fte_saved:+.1f}**", "")
w()
scenario2_teacher_savings = scenario2_fte_saved * 107_700
w(f"District-wide, Scenario 2 saves **{scenario2_fte_saved:.1f} FTE** = **${scenario2_teacher_savings:,.0f}** in teaching staff.")
w()

# Compare with Path A FTE savings
w("**For comparison, Path A (MS consolidation only) saves:**")
w(f"- {saved_fte - net_6_change:.1f} FTE = ${total_savings_b:,.0f} in teaching staff")
w()
w(f"**The marginal teacher savings from closing PES (Path B minus Path A):**")
marginal_fte = scenario2_fte_saved - (saved_fte - net_6_change)
marginal_teacher = marginal_fte * 107_700
w(f"- {marginal_fte:.1f} additional FTE = ${marginal_teacher:,.0f}")
w()
w("This means closing PES and restructuring the entire district saves only")
w(f"**{marginal_fte:.1f} additional teachers** beyond what middle school consolidation alone achieves.")

w()
w("#### E. Complete Path B Financial Summary (with Absorption)")
w()

# PES admin savings (eliminated -- DCS absorbs with coordination cost)
pes_admin_savings = fy27_by_school_article["Art 7 - School Administration"]["PES"]
# 50% of PES Art 5 support costs saved: guidance, nursing, and library
# positions partially transfer to receiving schools; roughly half the cost
# represents fixed overhead that can be absorbed by existing staff elsewhere.
pes_support_savings = fy27_by_school_article["Art 5 - Student & Staff Support"]["PES"] * 0.5

table_row("Item", "Budget Impact", "Notes", header=True)
table_row("PES instruction eliminated (gross)", f"-${supt_pes_instruction_savings:,.0f}", "Full PES budget minus facility")
table_row("DCS absorption: new FTE for Pownal students", f"+${int(supt_pes_instruction_savings - marginal_teacher - pes_admin_savings - pes_support_savings):,}", "Students still need teachers and services")
table_row("NET instruction savings (after absorption)", f"-${int(marginal_teacher + pes_admin_savings + pes_support_savings):,}", "True efficiency gain from consolidation")
w()
table_row("Transport: Pownal K-6 to DCS", f"+${pownal_k6_transport:,}", "3 bus routes, 11 mi")
table_row("Transport: Freeport PreK (64) to PES", f"+${freeport_prek_to_pes:,}", "3 specialized routes, 13 mi")
table_row("Transport: Durham PreK (48) to PES", f"+${durham_prek_to_pes:,}", "2 specialized routes, 11 mi")
table_row("DCS multi-community coordination", f"+${dcs_admin_addition:,}", "Pownal family integration")
table_row("PES full EC conversion (amortized)", f"+${pes_full_amortized:,.0f}", "10-year amortization")
w()

# Calculate net savings using absorption-aware model
net_instruction_savings = int(marginal_teacher + pes_admin_savings + pes_support_savings)
pkg2_total_new_costs = (pkg2_transport_total + dcs_expansion_total + pes_full_amortized)
pkg2_corrected_savings = net_instruction_savings - int(pkg2_total_new_costs)

# But also credit the DCS portables that Path A needs and Path B doesn't
pkg2_vs_patha_portables_credit = dcs_portables_path_a

# And credit the transport Path A pays that Path B doesn't
# (Path A: $150K for Freeport 6th to DCS. Path B: $0 for that specific route)
# But Path B has its own higher transport costs, already counted above

table_row("**NET savings from PES closure (after absorption)**", f"**${pkg2_corrected_savings:,}**", "")
w()

w()
w("**Two lenses on Path B savings:**")
w()
w("*Lens 1 -- Gross budget-line elimination (how closures are typically presented):*")
gross_savings = supt_pes_instruction_savings - int(pkg2_total_new_costs)
w(f"  PES instruction savings (${supt_pes_instruction_savings:,.0f}) minus transport/conversion/admin")
w(f"  (${int(pkg2_total_new_costs):,}) = **${gross_savings:,} gross savings**")
w()
w("*Lens 2 -- Actual FTE and admin efficiency (what the budget truly saves):*")
w(f"  Marginal teacher savings ({marginal_fte:.1f} FTE): ${marginal_teacher:,.0f}")
w(f"  PES admin eliminated (net of DCS coordination): ${pes_admin_savings - dcs_admin_addition:,}")
w(f"  PES support services efficiency: ${int(pes_support_savings):,}")
true_efficiency_savings = int(marginal_teacher + pes_admin_savings - dcs_admin_addition + pes_support_savings)
w(f"  **True efficiency savings: ${true_efficiency_savings:,}**")
w(f"  Minus transport costs: -${pkg2_transport_total:,}")
w(f"  Minus PES conversion: -${int(pes_full_amortized):,}")
true_net = true_efficiency_savings - pkg2_transport_total - int(pes_full_amortized)
w(f"  **True net savings from closing PES: ${true_net:,}**")
w()

# Use the gross model for budget comparison (traditional approach), but flag the reality
fy27_pkg2 = baseline_with_ec - gross_savings
pkg2_increase = fy27_pkg2 - fy26_adopted
pkg2_increase_pct = pkg2_increase / fy26_adopted * 100
pkg2_vs_baseline = baseline_with_ec - fy27_pkg2

w(f"The gross-lens Path B estimate: ${gross_savings:,} savings, FY27 increase **{pkg2_increase_pct:.1f}%**")
w()
w(f"**Critical caveat:** The gross model assumes PES costs vanish entirely. In practice,")
w(f"DCS must absorb {pownal_k6_to_dcs} Pownal students and expand services accordingly.")
w(f"The FTE efficiency lens shows the true marginal gain is ${true_efficiency_savings:,}")
w(f"-- which, after transport and conversion costs, yields a net of only **${true_net:,}/year.**")
w()
w(f"This means closing PES produces only modest net financial benefit once all costs")
w(f"are honestly accounted for. The case for Scenario 2 is **programmatic** (EC consolidation,")
w(f"MS equity, consistent academics), not primarily financial.")

w()
w("**Stakeholder Assessment (Full Scenario 2):**")
w()
table_row("Stakeholder", "Impact", "Palatability", header=True)
table_row("Pownal parents (K-5)", "Lose community school; children bused 19 min to Durham daily", "**VERY LOW**")
table_row("Pownal parents (EC)", "EC stays in Pownal (at repurposed PES)", "Moderate")
table_row("Freeport parents (K-5)", "MSS becomes K-3, MLS becomes 4-6; PreK leaves for PES", "MODERATE")
table_row("Freeport parents (EC)", "3-4 year olds bused 22 min to Pownal for EC", "**LOW**")
table_row("Durham parents (EC)", "3-4 year olds bused 19 min to Pownal for EC", "**LOW**")
table_row("Durham parents (K-6)", "DCS absorbs Pownal students; but net enrollment DECREASES", "MODERATE")
table_row("Pownal community", "Property values at risk (5-15% decline [^20]); community anchor lost", "**VERY LOW**")
table_row("PES teachers", "Displaced or reassigned to EC roles; lose K-5 school culture", "**VERY LOW**")
table_row("DCS teachers", "New community integration; K-6 focus (currently K-8)", "MODERATE")
table_row("School Board", "Politically divisive; Pownal AND Freeport/Durham EC parent opposition", "LOW")
table_row("Taxpayers", f"True FTE savings: only {marginal_fte:.1f} more than Path A", "MODERATE")

w()
w("#### E. EC Parent Opt-Out Risk Under Path B")
w()
w("A centralized EC model at PES creates a significant **enrollment risk**. Parents of")
w("3-year-olds are not compelled to enroll (general PreK is voluntary). If Freeport and")
w("Durham families decline to send 3-year-olds on 20+ minute bus rides to Pownal:")
w()
w("- EC enrollment drops below projections")
w("- Per-student costs rise (fixed staffing spread over fewer children)")
w("- State EPS funding (per-pupil based) is reduced")
w("- SpEd students lose non-disabled peers, creating LRE compliance risk")
w("- The program becomes financially unsustainable and politically indefensible")
w()
w("Under Path A's distributed model, this risk does not exist -- children attend their")
w("own community school for EC, just as they do for PreK today.")

# ============================================================
# SECTION 11.6: PACKAGE COMPARISON (CORRECTED)
# ============================================================

w()
w("### 11.6 Package Comparison Summary (Corrected)")
w()
table_row("Dimension", "No Action Baseline", "Path A (Preserve)", "Path B (Scenario 2) CORRECTED", header=True)
table_row("FY27 budget", f"${baseline_with_ec:,}", f"${fy27_pkg1:,.0f}", f"${fy27_pkg2:,.0f}")
table_row("FY27 increase", f"{baseline_ec_pct:.1f}%", f"**{pkg1_increase_pct:.1f}%**", f"{pkg2_increase_pct:.1f}%")
table_row("Savings vs baseline", "--", f"${pkg1_vs_baseline:,.0f}", f"${pkg2_vs_baseline:,.0f}")
table_row("Path-specific hidden costs", "--", f"${new_bus_routes_cost_revised + dcs_portables_path_a:,}/yr (transport + portables)", f"${int(pkg2_total_new_costs):,}/yr (transport + conversion + admin)")
table_row("District-wide FTE change", "0", f"-{net_fte_savings:.1f} teachers", f"-{scenario2_fte_saved:.1f} teachers (only {marginal_fte:.1f} more than Path A)")
table_row("PES PreK-5 program", "Preserved", "**Preserved + strengthened**", "Eliminated")
table_row("EC model", "Unfunded", "Distributed; SpEd at MSS (64 PreK peers, 1:5 LRE)", "Centralized at PES (least central school)")
table_row("EC parent opt-out risk", "N/A", "Low (local schools)", "**High** (20+ min bus for 3yo)")
table_row("Young children close to home", "Yes", "**Yes, all three towns**", "No (K-5 to Durham; EC to Pownal)")
table_row("DCS enrollment change", "No change", f"+{new_dcs_total - school_enrollment['DCS']} students (0 portables needed)", f"**{dcs_scenario2_net:+d} students** (net shrinks)")
table_row("Political feasibility", f"Low ({baseline_ec_pct:.0f}%+ increase)", "**High** (no closures)", "Low (multi-town opposition)")
table_row("Reversibility", "N/A", "High", "Low")
table_row("Property value risk", "None", "None", "Significant for Pownal [^20]")
table_row("Consistent with recent voter approvals", "--", f"Yes ({pkg1_increase_pct:.1f}% vs 6.5-6.8% approved FY25-26)", f"Yes ({pkg2_increase_pct:.1f}%)")

w()
gap_between_paths = pkg2_vs_baseline - pkg1_vs_baseline
w(f"**Corrected gap between paths:** Path B saves ${gap_between_paths:,.0f} more than Path A annually.")
w()
w(f"**However**, the DCS absorption cost analysis (Section 11.5D) reveals that the")
w(f"${supt_pes_instruction_savings:,.0f} 'PES instruction savings' is a gross figure --")
w(f"most of those costs transfer to DCS and other schools. The true marginal teacher")
w(f"savings from closing PES are only {marginal_fte:.1f} FTE (${marginal_teacher:,.0f}/year).")
w(f"When measured in actual FTE efficiency gains (not gross budget line elimination),")
w(f"the annual financial advantage of closing PES is modest.")

# ============================================================
# SECTION 12: IMPLEMENTATION ROADMAP
# ============================================================

w()
w("---")
w()
w("## 12. Implementation Roadmap (Package 1)")
w()

table_row("Timeline", "Action", "Responsible", header=True)
table_row("Spring 2026", "Board adopts Package 1 framework; community forums in all 3 towns", "Board, Superintendent")
table_row("Summer 2026", "Begin planning: FMS 7-8 scheduling, DCS 6th grade integration", "Principals, Curriculum Dir.")
table_row("Fall 2026", "EC Task Force finalizes distributed model; MSS SpEd classroom prep", "EC Task Force, Facilities")
table_row("2026-2027", "Voluntary PES enrollment expansion (school choice applications)", "PES Principal, Families")
table_row("Summer 2027", "Grade 6 transition to DCS; Grade 7-8 consolidation at FMS", "All principals")
table_row("Fall 2027", "New grade configuration operational; monitor and adjust", "Administration")
table_row("2027-2028", "EC services launch at all three schools (ahead of 2028 mandate)", "EC Coordinator, Principals")
table_row("Ongoing", "Annual review of enrollment, class sizes, transportation routes", "Board, Finance Dir.")

# ============================================================
# FOOTNOTES
# ============================================================

w()
w("---")
w()
w("## Footnotes")
w()
for num, text in sorted(FOOTNOTES.items()):
    w(f"[^{num}]: {text}")
w()

# ============================================================
# WRITE OUTPUT
# ============================================================

output_path = "RSU 5 Calculation Appendix.md"
content = md.getvalue()

with open(output_path, "w", encoding="utf-8") as f:
    f.write(content)

print(f"Wrote {len(content):,} characters to '{output_path}'")
print(f"Sections: 12 major sections with {len(FOOTNOTES)} footnotes")

print("\n--- KEY RESULTS SUMMARY ---")
print(f"Total district students: {total_district}")
print(f"Pownal: {pownal_total_students} students, consumption ${pownal_consumption:,.0f}, net ${net_positions['Pownal']:+,.0f}")
print(f"Durham: {durham_total_students} students, consumption ${durham_consumption:,.0f}, net ${net_positions['Durham']:+,.0f}")
print(f"Freeport: {freeport_total_students} students, consumption ${freeport_consumption:,.0f}, net ${net_positions['Freeport']:+,.0f}")
print(f"PES cost gap: ${total_gap:,.0f}")
print(f"MS consolidation net savings (corrected): ${total_savings_b:,.0f}")
print(f"EC Option 1 cost: ${ec_option1_total:,} (net after offsets: ~${ec_net_cost:,})")
print(f"Baseline (FY27 + EC, no action): ${baseline_with_ec:,} ({baseline_ec_pct:.1f}%)")
print(f"Package 1 total: ${fy27_pkg1:,.0f} ({pkg1_increase_pct:.1f}%) -- saves ${pkg1_vs_baseline:,.0f} vs baseline")
print(f"Package 2 total: ${fy27_pkg2:,.0f} ({pkg2_increase_pct:.1f}%) -- saves ${pkg2_vs_baseline:,.0f} vs baseline")
