"""
Clean the IFC (International Finance Corporation) Investment Services Projects dataset
and export a Power BI-ready star schema (fact + dimension tables).

Source: World Bank Group Finances One Open Data
https://financesone.worldbank.org/ifc-investment-services-projects/DS00499
"""
import pandas as pd
import numpy as np

RAW_PATH = "data/ifc_raw.csv"

df = pd.read_csv(RAW_PATH)
print(f"Loaded {len(df):,} rows, {df.shape[1]} columns.")

# ---------------------------------------------------------------
# 1. Standardize column names (snake_case)
# ---------------------------------------------------------------
rename_map = {
    "Date Disclosed": "date_disclosed",
    "Project Name": "project_name",
    "Document Type": "document_type",
    "Project Number": "project_number",
    "Project Url": "project_url",
    "Product Line": "product_line",
    "Company Name": "company_name",
    "Country": "country",
    "IFC Country Code": "ifc_country_code",
    "Industry": "industry",
    "Environmental Category": "environmental_category",
    "Department": "department",
    "Status": "status",
    "Projected Board Date": "projected_board_date",
    "IFC Approval Date": "approval_date",
    "IFC Signed Date": "signed_date",
    "IFC Invested Date": "invested_date",
    "IFC investment for Risk Management(Million - USD)": "investment_risk_mgmt_usdm",
    "IFC investment for Guarantee(Million - USD)": "investment_guarantee_usdm",
    "IFC investment for Loan(Million - USD)": "investment_loan_usdm",
    "IFC investment for Equity(Million - USD)": "investment_equity_usdm",
    "Total IFC investment as approved by Board(Million - USD)": "total_investment_usdm",
    "WB Country Code": "wb_country_code",
    "As of Date": "as_of_date",
    "Board approval - Fiscal year": "board_approval_fy",
    "Board approval - Calendar year": "board_approval_cy",
}
df = df.rename(columns=rename_map)

# ---------------------------------------------------------------
# 2. Fix data types
# ---------------------------------------------------------------
date_cols = ["date_disclosed", "projected_board_date", "approval_date",
             "signed_date", "invested_date", "as_of_date"]
for c in date_cols:
    df[c] = pd.to_datetime(df[c], errors="coerce")

for c in ["project_name", "company_name", "country", "industry", "status", "product_line"]:
    df[c] = df[c].astype(str).str.strip().replace("nan", np.nan)

# ---------------------------------------------------------------
# 3. Handle missing values (documented, not silently dropped)
# ---------------------------------------------------------------
missing_product_line = df["product_line"].isna().sum()
df["product_line"] = df["product_line"].fillna("Not Specified")

missing_company = df["company_name"].isna().sum()
df["company_name"] = df["company_name"].fillna("(Unnamed Company)")

# Investment amount columns: NaN means "not applicable to this instrument", not missing data
investment_cols = ["investment_risk_mgmt_usdm", "investment_guarantee_usdm",
                    "investment_loan_usdm", "investment_equity_usdm"]
for c in investment_cols:
    df[c] = df[c].fillna(0.0)

missing_total = df["total_investment_usdm"].isna().sum()
# Where total is missing but component investments exist, derive it
derived_total = df[investment_cols].sum(axis=1)
df["total_investment_usdm"] = df["total_investment_usdm"].fillna(derived_total)

# ---------------------------------------------------------------
# 4. Derived / enrichment columns
# ---------------------------------------------------------------
df["is_active"] = df["status"].isin(["Active", "Pending Disbursement"])
df["is_committed"] = df["status"].isin(["Active", "Completed", "Pending Disbursement"])
df["days_approval_to_signed"] = (df["signed_date"] - df["approval_date"]).dt.days
df["primary_instrument"] = df[investment_cols].idxmax(axis=1).str.replace(
    "investment_", "").str.replace("_usdm", "").where(df[investment_cols].sum(axis=1) > 0, "Not Specified")

# ---------------------------------------------------------------
# 5. Data quality checks
# ---------------------------------------------------------------
dupes = df.duplicated(subset=["project_number"]).sum()
negative_amounts = (df["total_investment_usdm"] < 0).sum()
print(f"Duplicate project numbers: {dupes}")
print(f"Negative investment amounts: {negative_amounts}")
print(f"Product line backfilled ('Not Specified'): {missing_product_line}")
print(f"Company name backfilled: {missing_company}")
print(f"Total investment derived from components: {missing_total}")

# ---------------------------------------------------------------
# 6. Save cleaned flat file
# ---------------------------------------------------------------
df.to_csv("data/ifc_clean.csv", index=False)
print(f"\nSaved cleaned dataset: {df.shape[0]:,} rows, {df.shape[1]} columns.")

# ---------------------------------------------------------------
# 7. Build star schema for Power BI
# ---------------------------------------------------------------
dim_country = df[["ifc_country_code", "country", "wb_country_code"]].drop_duplicates(
    subset=["country"]).reset_index(drop=True)
dim_country.insert(0, "country_id", dim_country.index + 1)

dim_industry = df[["industry"]].drop_duplicates().reset_index(drop=True)
dim_industry.insert(0, "industry_id", dim_industry.index + 1)

dim_status = df[["status"]].drop_duplicates().reset_index(drop=True)
dim_status.insert(0, "status_id", dim_status.index + 1)

dim_product_line = df[["product_line"]].drop_duplicates().reset_index(drop=True)
dim_product_line.insert(0, "product_line_id", dim_product_line.index + 1)

fact = df.merge(dim_country[["country", "country_id"]], on="country", how="left")
fact = fact.merge(dim_industry, on="industry", how="left")
fact = fact.merge(dim_status, on="status", how="left")
fact = fact.merge(dim_product_line, on="product_line", how="left")

fact_cols = [
    "project_number", "project_name", "company_name", "country_id", "industry_id",
    "status_id", "product_line_id", "primary_instrument",
    "investment_risk_mgmt_usdm", "investment_guarantee_usdm",
    "investment_loan_usdm", "investment_equity_usdm", "total_investment_usdm",
    "is_active", "is_committed", "board_approval_fy", "board_approval_cy",
    "date_disclosed", "approval_date", "signed_date", "invested_date",
]
fact_projects = fact[fact_cols]

dim_country.to_csv("data/dim_country.csv", index=False)
dim_industry.to_csv("data/dim_industry.csv", index=False)
dim_status.to_csv("data/dim_status.csv", index=False)
dim_product_line.to_csv("data/dim_product_line.csv", index=False)
fact_projects.to_csv("data/fact_projects.csv", index=False)

print("\nStar schema exported:")
print(f"  fact_projects.csv      {fact_projects.shape}")
print(f"  dim_country.csv        {dim_country.shape}")
print(f"  dim_industry.csv       {dim_industry.shape}")
print(f"  dim_status.csv         {dim_status.shape}")
print(f"  dim_product_line.csv   {dim_product_line.shape}")
