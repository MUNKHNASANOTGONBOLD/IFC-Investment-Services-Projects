# IFC Investment Services Projects — Portfolio Analytics

Python cleaning pipeline + Power BI-ready star schema + interactive dashboard, built on real, public IFC (International Finance Corporation) disclosure data.

## Dataset

**Source:** [World Bank Group Finances One Open Data — IFC Investment Services Projects](https://financesone.worldbank.org/ifc-investment-services-projects/DS00499), snapshot as of 8/11/2026.

**Scale:** 7,083 project disclosures, 172 countries, 26 original columns — investment amounts by instrument (loan, equity, guarantee, risk management), status, industry, and key dates.

## Pipeline

1. **Python cleaning** (`clean_ifc_data.py`) — standardized column names to snake_case, parsed date fields, backfilled missing product line (382 rows, "Not Specified") and company name (25 rows), treated missing instrument-amount columns as "not applicable" (set to 0) rather than data gaps, derived total investment for 295 rows from instrument components. 17 project numbers appear twice — each pair is a "Proposed Investment" and later "Investment Information" disclosure for the same project, not a duplication error.
2. **Star schema export** — `fact_projects` (7,083 rows) + `dim_country`, `dim_industry`, `dim_status`, `dim_product_line` lookup tables.
3. **Power BI data model** (`build_powerbi_model.py`) — Excel workbook with Summary tab and 8 DAX measures ready to paste into Power BI.
4. **Dashboard** (`IFC_Portfolio_Dashboard.html`) — multi-page interactive explorer (Overview / Industry & Instrument / Portfolio Detail) with live country, industry, and status filters, and an instrument-mix visualization (Loan / Equity / Guarantee).

## Repo structure

```
data/
  ifc_raw.csv                    Original uploaded snapshot
  ifc_clean.csv                  Cleaned flat file (30 columns)
  fact_projects.csv              Fact table (star schema)
  dim_country.csv / dim_industry.csv / dim_status.csv / dim_product_line.csv
  IFC_PowerBI_Data_Model.xlsx    Excel workbook: star schema + Summary + DAX measures
clean_ifc_data.py                Cleaning + star-schema export script
build_powerbi_model.py           Excel workbook builder
IFC_Portfolio_Dashboard.html     Multi-page dashboard with sidebar filters
```

## Tools

Python (pandas, NumPy) · Excel · Power BI (DAX) · Chart.js (prototype)
