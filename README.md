# Workday Report Catalog Agent

An AI agent that fetches Workday report metadata via the ISU RaaS API, processes the JSON response, and generates a structured, multi-tab Excel workbook organized by industry — with dependency flagging for reports containing calculated fields.

Now includes a **Browser Automation Agent** powered by `browser-use` + Google Gemini that can log into the Workday UI and perform defined tasks.

---

## Features

### Report Catalog Agent (`workday_agent.py`)
- **Workday RaaS Integration** — Authenticates via ISU credentials and fetches report metadata as JSON
- **AI-Tag Pre-Filtering** — Parser silently excludes reports without an `AI_*` tag, keeping your catalog output focused
- **Industry-Aware Routing** — Maps `Report_Tag` values to industry tabs (Healthcare, Insurance, Banking, etc.)
- **Dependency Detection** — Flags reports with `Calculated_Fields_for_Report` as having dependencies
- **Multi-Tag Support** — Reports with multiple tags (e.g. `AI_Healthcare; AI_Insurance`) appear in all relevant tabs
- **Smart Tag Detection** — Auto-creates tabs for unknown `AI_*` prefixed tags
- **Professional Excel Output** — Color-coded tabs, frozen headers, auto-filters, alternating rows, dependency highlighting

### Browser Agent (`workday_browser_agent.py`)
- **LLM-Driven Browser Automation** — Uses Google Gemini to interpret and interact with Workday's UI
- **Visible Browser Mode** — Watch the agent navigate Workday in real-time
- **Predefined Task Library** — Ready-made tasks for Reports, Finance, Procurement, Admin, Security, and more
- **Pipeline Integration** — Runs only after the report agent completes, receiving the Excel output path as context
- **Custom Tasks** — Send any natural language instruction to interact with the Workday UI
- **Structured Results** — Returns status, result data, and timestamp in a clean dict format

---

## Quick Start

### 1. Install Dependencies
Ensure you have Python 3.11+ installed, then run:
```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Configure Credentials
Copy `.env.example` to `.env` and fill in your details:
```bash
cp .env.example .env
```

You need:
- **RaaS API**: ISU username/password and the report URL
- **Browser Agent**: Workday UI login credentials and a Google Gemini API key

### 3. Run Report Catalog Agent
```bash
# Live mode (calls Workday API)
python workday_agent.py

# Demo mode (offline with sample data)
python workday_agent.py --sample
```

### 4. Run Browser Agent

```bash
# List available predefined tasks
python workday_browser_agent.py --list

# Run a predefined task (browser only)
python workday_browser_agent.py --task navigate_to_reports

# Run a custom task (browser only)
python workday_browser_agent.py --custom "Go to Reports and list all custom reports"

# Full pipeline: Report Agent → Excel → Browser Agent
python workday_browser_agent.py --pipeline --task navigate_to_reports

# Full pipeline with sample data
python workday_browser_agent.py --pipeline --sample --task navigate_to_reports
```

---

## Project Structure

```
Workday_Report_Migration_Agent/
├── workday_agent.py              # Report Catalog Agent (RaaS API → Excel)
├── workday_browser_agent.py      # Browser Automation Agent (UI tasks via browser-use)
├── output/                       # Generated Excel files
├── .env                          # Local credentials (git-ignored)
├── .env.example                  # Template env file
├── .gitignore
├── requirements.txt
└── README.md
```

---

## Pipeline Flow

```
┌─────────────────────────────────────────────────────────┐
│                    FULL PIPELINE                         │
│                                                         │
│  Phase 1: Report Catalog Agent                          │
│  ┌───────────┐    ┌──────────┐    ┌────────────────┐   │
│  │ Workday   │───>│ Parse &  │───>│ Generate Excel │   │
│  │ RaaS API  │    │ Filter   │    │ Workbook       │   │
│  └───────────┘    └──────────┘    └──────┬─────────┘   │
│                                          │              │
│                                    Excel Path           │
│                                          │              │
│  Phase 2: Browser Agent                  │              │
│  ┌───────────┐    ┌──────────┐    ┌──────▼─────────┐   │
│  │ Gemini    │───>│ Browser  │───>│ Execute Task   │   │
│  │ LLM       │    │ (Chrome) │    │ in Workday UI  │   │
│  └───────────┘    └──────────┘    └────────────────┘   │
│                                                         │
└─────────────────────────────────────────────────────────┘
```

---

## Predefined Browser Tasks

| Task Key | Description |
|----------|-------------|
| `navigate_to_reports` | Navigate to Reports and list all visible reports |
| `run_custom_report` | Find and run a specific custom report |
| `check_report_status` | Check status of recently run reports |
| `navigate_to_journals` | Find journals with In Progress status |
| `check_period_close` | Check current period close status |
| `view_account_balances` | Extract top-level account balances |
| `check_pending_pos` | List purchase orders pending approval |
| `check_supplier_invoices` | List supplier invoices awaiting action |
| `check_business_process` | Check a specific BP configuration |
| `navigate_to_integrations` | Check integration system status |
| `check_security_group` | Extract security group members and permissions |
| `check_tenant_settings` | Extract tenant configuration settings |
| `verify_reports_from_catalog` | Verify reports from Excel catalog exist in Workday |

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
- Python 3.11+
- `requests` (for RaaS calls)
- `openpyxl` (for Excel generation)
- `browser-use` (for browser automation)
- `playwright` (browser engine)
- `langchain-google-genai` (Gemini LLM integration)
