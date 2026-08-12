import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

NAVY = "0B2E4A"
header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor=NAVY)
title_font = Font(name="Arial", size=14, bold=True, color=NAVY)
sub_font = Font(name="Arial", size=10, italic=True, color="595959")
normal_font = Font(name="Arial", size=10)
thin = Side(style="thin", color="B7B7B7")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

def write_df(ws, df, start_row=1):
    for j, col in enumerate(df.columns, start=1):
        ws.cell(row=start_row, column=j, value=col)
    style_header(ws, start_row, len(df.columns))
    for i, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for j, val in enumerate(row, start=1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = normal_font
    for j, col in enumerate(df.columns, start=1):
        maxlen = max(len(str(col)), df[col].map(lambda x: len(str(x))).max() if len(df) else 0)
        ws.column_dimensions[get_column_letter(j)].width = min(max(maxlen + 2, 10), 42)

wb = Workbook()

ws = wb.active
ws.title = "ReadMe"
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 110
ws["B2"] = "IFC Investment Services Projects — Data Model"
ws["B2"].font = title_font
ws["B3"] = "Star-schema data model for Power BI, cleaned from the IFC's public Investment Services Projects disclosure data (World Bank Group Finances One)."
ws["B3"].font = sub_font
ws["B3"].alignment = Alignment(wrap_text=True)
ws["B5"] = "Source"
ws["B5"].font = Font(bold=True)
ws["B6"] = "financesone.worldbank.org/ifc-investment-services-projects/DS00499 — real, public IFC project disclosure data (as of 8/11/2026)."
ws["B6"].alignment = Alignment(wrap_text=True)
ws["B8"] = "Tabs"
ws["B8"].font = Font(bold=True)
notes = [
    "fact_projects — one row per IFC investment disclosure (7,083 rows): amounts by instrument, dates, status/country/industry keys.",
    "dim_country, dim_industry, dim_status, dim_product_line — lookup dimensions for the star schema.",
    "Summary — headline portfolio KPIs.",
    "DAX_Measures — measure definitions to paste into Power BI after loading the tables.",
]
for i, n in enumerate(notes, start=9):
    ws.cell(row=i, column=2, value="• " + n).font = normal_font
    ws.cell(row=i, column=2).alignment = Alignment(wrap_text=True)
ws["B14"] = ("Cleaning notes: column names standardized to snake_case; dates parsed; missing product_line (382 rows, "
             "instrument not disclosed) set to 'Not Specified'; missing company_name (25 rows) backfilled; investment "
             "amount columns' NaNs mean 'not applicable to this instrument' and were set to 0; total_investment derived "
             "for 295 rows from instrument components where the board-approved total was blank. 17 project_number values "
             "appear twice — each pair is a 'Proposed Investment' and later 'Investment Information' disclosure for the "
             "same project (not a data error), left as-is to preserve both disclosure stages.")
ws["B14"].font = sub_font
ws["B14"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[14].height = 75

fact = pd.read_csv("data/fact_projects.csv")
dim_country = pd.read_csv("data/dim_country.csv")
dim_industry = pd.read_csv("data/dim_industry.csv")
dim_status = pd.read_csv("data/dim_status.csv")
dim_product_line = pd.read_csv("data/dim_product_line.csv")

for name, d in [("fact_projects", fact), ("dim_country", dim_country), ("dim_industry", dim_industry),
                 ("dim_status", dim_status), ("dim_product_line", dim_product_line)]:
    ws = wb.create_sheet(name)
    write_df(ws, d)

ws = wb.create_sheet("Summary")
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 20
ws["B2"] = "Portfolio Summary"
ws["B2"].font = title_font

total_investment = fact["total_investment_usdm"].sum()
active_count = int(fact["is_active"].sum())
committed_count = int(fact["is_committed"].sum())

stats = [
    ("Total Projects/Disclosures", f"{len(fact):,}"),
    ("Countries Covered", f"{dim_country.shape[0]}"),
    ("Total IFC Investment (Board-approved, US$M)", f"${total_investment:,.0f}M"),
    ("Active Projects", f"{active_count:,}"),
    ("Committed Projects (Active/Completed/Pending Disb.)", f"{committed_count:,}"),
]
r = 4
for label, val in stats:
    ws.cell(row=r, column=2, value=label).font = normal_font
    ws.cell(row=r, column=3, value=val).font = Font(bold=True, size=10)
    r += 1

r += 2
ws.cell(row=r, column=2, value="By Industry").font = Font(bold=True, color=NAVY)
r += 1
by_industry = (
    fact.merge(dim_industry, on="industry_id")
    .groupby("industry")
    .agg(total_investment=("total_investment_usdm", "sum"), count=("project_number", "count"))
    .reset_index()
    .sort_values("total_investment", ascending=False)
)
write_df(ws, by_industry, start_row=r)

ws = wb.create_sheet("DAX_Measures")
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 95
ws["B2"] = "DAX Measures — paste into Power BI after loading the star schema"
ws["B2"].font = title_font
ws.cell(row=4, column=2, value="Measure Name")
ws.cell(row=4, column=3, value="DAX Formula")
style_header(ws, 4, 2)

measures = [
    ("Total Investment (US$M)", "Total Investment (US$M) = SUM(fact_projects[total_investment_usdm])"),
    ("Total Loan Investment", "Total Loan Investment = SUM(fact_projects[investment_loan_usdm])"),
    ("Total Equity Investment", "Total Equity Investment = SUM(fact_projects[investment_equity_usdm])"),
    ("Total Guarantee Investment", "Total Guarantee Investment = SUM(fact_projects[investment_guarantee_usdm])"),
    ("Active Project Count", "Active Project Count = CALCULATE(COUNTROWS(fact_projects), fact_projects[is_active] = TRUE)"),
    ("Avg Investment Size", "Avg Investment Size = DIVIDE([Total Investment (US$M)], COUNTROWS(fact_projects), 0)"),
    ("Equity Share of Portfolio", "Equity Share of Portfolio = DIVIDE([Total Equity Investment], [Total Investment (US$M)], 0)"),
    ("Rank Country by Investment", "Rank Country by Investment = RANKX(ALL(dim_country[country]), [Total Investment (US$M)])"),
]
for name, formula in measures:
    r = ws.max_row + 1
    ws.cell(row=r, column=2, value=name).font = Font(bold=True, size=10)
    ws.cell(row=r, column=2).alignment = Alignment(vertical="top", wrap_text=True)
    c = ws.cell(row=r, column=3, value=formula)
    c.font = Font(name="Consolas", size=9)
    c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[r].height = 32

order = ["ReadMe", "fact_projects", "dim_country", "dim_industry", "dim_status", "dim_product_line", "Summary", "DAX_Measures"]
wb._sheets = [wb[n] for n in order]

wb.save("data/IFC_PowerBI_Data_Model.xlsx")
print("Saved.")
