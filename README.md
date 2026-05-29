# Workday Report Catalog Agent

An AI agent that fetches Workday report metadata via ISU RaaS API, processes the JSON response, and generates a structured multi-tab Excel workbook organized by industry — with dependency flagging for reports containing calculated fields.

## Features

- **Workday RaaS Integration** — Authenticates via ISU credentials and fetches report metadata as JSON
- **Industry-Aware Routing** — Maps `Report_Tag` values to industry tabs (Healthcare, Insurance, Banking, etc.)
- **Dependency Detection** — Flags reports with `Calculated_Fields_for_Report` as having dependencies
- **Multi-Tag Support** — Reports with multiple tags (e.g. `AI_Healthcare; AI_Insurance`) appear in all relevant tabs
- **Smart Tag Detection** — Auto-creates tabs for unknown `AI_*` prefixed tags
- **Professional Excel Output** — Color-coded tabs, frozen headers, auto-filters, alternating rows, dependency highlighting

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Configure credentials (edit .env)
#    Set your Workday tenant URL, ISU username/password, and report name

# 3. Run with sample data (demo/testing)
python main.py --sample

# 4. Run against live Workday RaaS
python main.py
```

## Project Structure

```
workday-report-agent/
├── main.py                        # Entry point & orchestration
├── config.py                      # All configuration & color scheme
├── api/
│   └── raas_client.py             # Workday RaaS API client (retry + auth)
├── processors/
│   ├── json_parser.py             # JSON response parsing
│   ├── industry_router.py         # Report_Tag → Industry mapping
│   └── dependency_checker.py      # Calculated field detection
├── excel/
│   ├── workbook_builder.py        # Multi-tab workbook orchestrator
│   ├── summary_tab.py             # Summary tab creation
│   └── industry_tab.py            # Industry tab creation
├── utils/
│   └── helpers.py                 # Shared utilities (borders, etc.)
├── output/                        # Generated Excel files
├── .env                           # API credentials (never commit)
├── .gitignore
├── requirements.txt
└── README.md
```

## Excel Output

The generated workbook contains:

| Tab | Description |
|-----|-------------|
| **Summary** | Bird's-eye view: report counts, dependency %, data sources per industry |
| **Healthcare** | All reports tagged `AI_Healthcare` |
| **Insurance** | All reports tagged `AI_Insurance` |
| **Banking** | All reports tagged `AI_Banking` |
| **Financial Services** | All reports tagged `AI_FinancialServices` |
| **Retail** | All reports tagged `AI_Retail` |
| **Manufacturing** | All reports tagged `AI_Manufacturing` |
| **Technology** | All reports tagged `AI_Technology` / `AI_Tech` |
| **Government** | All reports tagged `AI_Government` / `AI_PublicSector` |
| **Education** | All reports tagged `AI_Education` |
| **Real Estate** | All reports tagged `AI_RealEstate` |
| **Other / Untagged** | Reports with empty or unrecognized tags |

Dependency rows are highlighted in **soft yellow** with green badges in the "Has Dependencies" column.

## Configuration

### Environment Variables (`.env`)

| Variable | Description |
|----------|-------------|
| `WORKDAY_TENANT_URL` | Your Workday tenant URL |
| `WORKDAY_TENANT_NAME` | Tenant name for the API path |
| `WORKDAY_REPORT_NAME` | Custom report name exposed via RaaS |
| `WORKDAY_ISU_USERNAME` | Integration System User username |
| `WORKDAY_ISU_PASSWORD` | ISU password |
| `WORKDAY_API_VERSION` | API version (default: `v43.0`) |
| `WORKDAY_TIMEOUT` | Request timeout in seconds (default: `60`) |
| `WORKDAY_RETRIES` | Number of retry attempts (default: `3`) |
| `OUTPUT_DIR` | Output directory for Excel files (default: `./output`) |

### Adding Custom Industry Tags

Edit `config.py` → `INDUSTRY_CONFIG["custom_mappings"]` to add new tag-to-industry mappings:

```python
"custom_mappings": {
    "AI_Pharma": "Pharmaceutical",
    "AI_Telecom": "Telecommunications",
    # ...
}
```

## Requirements

- Python 3.10+
- `requests` — HTTP client for RaaS calls
- `openpyxl` — Excel file generation
- `python-dotenv` — Environment variable management
