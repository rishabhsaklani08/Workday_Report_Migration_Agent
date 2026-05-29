# Workday Report Catalog Agent

An AI agent that fetches Workday report metadata via the ISU RaaS API, processes the JSON response, and generates a structured, multi-tab Excel workbook organized by industry — with dependency flagging for reports containing calculated fields.

The codebase is consolidated into a single master file (`workday_agent.py`) for ease of execution and zero-dependency environment configuration.

---

## Features

- **Workday RaaS Integration** — Authenticates via ISU credentials and fetches report metadata as JSON
- **AI-Tag Pre-Filtering** — Parser silently excludes reports without an `AI_*` tag, keeping your catalog output focused
- **Industry-Aware Routing** — Maps `Report_Tag` values to industry tabs (Healthcare, Insurance, Banking, etc.)
- **Dependency Detection** — Flags reports with `Calculated_Fields_for_Report` as having dependencies
- **Multi-Tag Support** — Reports with multiple tags (e.g. `AI_Healthcare; AI_Insurance`) appear in all relevant tabs
- **Smart Tag Detection** — Auto-creates tabs for unknown `AI_*` prefixed tags
- **Professional Excel Output** — Color-coded tabs, frozen headers, auto-filters, alternating rows, dependency highlighting

---

## Quick Start

### 1. Install Dependencies
Ensure you have Python 3.10+ installed, then run:
```bash
pip install -r requirements.txt
```

### 2. Configure Credentials
Copy `.env.example` to `.env` and fill in your details:
```bash
cp .env.example .env
```
Provide the full RaaS URL (customreport2 or Reporting_UI format) and your ISU credentials.

### 3. Run Demo Mode (Test Offline)
Uses built-in sample data to generate an Excel sheet:
```bash
python workday_agent.py --sample
```

### 4. Run Live Mode (Production)
Calls the real Workday endpoint:
```bash
python workday_agent.py
```

---

## Clean Project Structure

```
workday-report-agent/
├── workday_agent.py               # Master agent script containing all logic
├── output/                        # Directory where generated Excel files are saved
├── .env                           # Local credentials (ignored by git)
├── .env.example                   # Template env file
├── .gitignore
├── requirements.txt
└── README.md
```

---

## Excel Output Design

The generated workbook contains:

| Tab | Description |
|-----|-------------|
| **Summary** | Bird's-eye view: report counts, dependency %, categories list, unique data sources |
| **Healthcare** | Reports tagged `AI_Healthcare` |
| **Insurance** | Reports tagged `AI_Insurance` |
| **Banking** | Reports tagged `AI_Banking` |
| **Financial Services** | Reports tagged `AI_FinancialServices` / `AI_Finance` |
| **Retail** | Reports tagged `AI_Retail` |
| **Manufacturing** | Reports tagged `AI_Manufacturing` |
| **Technology** | Reports tagged `AI_Technology` / `AI_Tech` |
| **Government** | Reports tagged `AI_Government` / `AI_PublicSector` |
| **Education** | Reports tagged `AI_Education` |
| **Real Estate** | Reports tagged `AI_RealEstate` |

* Dependency rows are highlighted in **soft yellow** with green badges in the "Has Dependencies" column.
* Column structure matches real Workday JSON: *No., Report Name, Type, Output Type, Category, Data Source, Owner, Tag, Fields, Calc Fields, Dependencies, Details, Created On, Last Updated*.

---

## Custom Industry Tag Mappings
To modify mapping rules, edit the `INDUSTRY_CONFIG` dictionary at the top of `workday_agent.py`:
```python
"custom_mappings": {
    "AI_Pharma": "Pharmaceutical",
    "AI_Telecom": "Telecommunications",
    # ...
}
```

---

## Requirements
- Python 3.10+
- `requests` (for RaaS calls)
- `openpyxl` (for Excel generation)
