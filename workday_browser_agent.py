"""
workday_browser_agent.py — Workday Browser Automation Agent.

Uses the browser-use library + Google Gemini LLM to log into
a Workday tenant via the browser and perform defined UI tasks.

This agent is designed to run AFTER the report catalog agent
(workday_agent.py) has generated its Excel output. It receives
the Excel path and can then perform follow-up browser tasks
like verifying reports, navigating to specific configurations,
or extracting additional data from the Workday UI.

Requirements:
    pip install browser-use playwright langchain-google-genai python-dotenv
    playwright install chromium
"""

import os
import asyncio
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────────────────────────
# Environment Loader (shared with workday_agent.py)
# ─────────────────────────────────────────────────────────────
_env_path = Path(__file__).resolve().parent / ".env"
if _env_path.exists():
    for _line in _env_path.read_text(encoding="utf-8").splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _key, _, _val = _line.partition("=")
            os.environ.setdefault(_key.strip(), _val.strip())

# ─────────────────────────────────────────────────────────────
# Third-party imports (deferred to allow clean error messages)
# ─────────────────────────────────────────────────────────────
try:
    from browser_use import Agent, Browser, BrowserProfile
    from langchain_google_genai import ChatGoogleGenerativeAI
except ImportError as e:
    raise ImportError(
        f"Missing dependency: {e}. "
        "Run: pip install browser-use langchain-google-genai playwright && "
        "playwright install chromium"
    )

# ─────────────────────────────────────────────────────────────
# SECTION 1: Configuration
# ─────────────────────────────────────────────────────────────
BROWSER_CONFIG = {
    "tenant_url": os.getenv(
        "WORKDAY_TENANT_URL",
        "https://wd2-impl-services1.workday.com/accenture_dpt3"
    ),
    "username": os.getenv("WORKDAY_USERNAME", ""),
    "password": os.getenv("WORKDAY_PASSWORD", ""),
    "headless": os.getenv("BROWSER_HEADLESS", "false").lower() == "true",
    "timeout": int(os.getenv("BROWSER_TIMEOUT", "30000")),
}

GEMINI_API_KEY = os.getenv("GOOGLE_API_KEY", "")

# Validate required credentials
def _validate_config():
    """Check that all required configuration is present."""
    missing = []
    if not BROWSER_CONFIG["username"]:
        missing.append("WORKDAY_USERNAME")
    if not BROWSER_CONFIG["password"]:
        missing.append("WORKDAY_PASSWORD")
    if not GEMINI_API_KEY:
        missing.append("GOOGLE_API_KEY")
    if missing:
        raise EnvironmentError(
            f"Missing required environment variables: {', '.join(missing)}. "
            "Add them to your .env file."
        )

# ─────────────────────────────────────────────────────────────
# SECTION 2: Predefined Workday Task Library
# ─────────────────────────────────────────────────────────────
WORKDAY_TASKS = {

    # ── REPORT TASKS ──────────────────────────────────────────
    "navigate_to_reports": """
        Navigate to the Reports section in Workday.
        Look for 'All Reports' or 'Custom Reports'.
        Extract the list of all visible report names and their types.
        Return the data as a structured list.
    """,

    "run_custom_report": """
        Navigate to Workday Reports section.
        Search for the report named: {report_name}
        Run the report with default parameters.
        Wait for results to load completely.
        Extract the data shown on screen.
        Return the results.
    """,

    "check_report_status": """
        Navigate to the Report Monitor or Scheduled Reports in Workday.
        Check the status of recently run reports.
        Return the status of each report found including
        report name, status, start time, and completion time.
    """,

    # ── FINANCE TASKS ─────────────────────────────────────────
    "navigate_to_journals": """
        Navigate to Accounting > Journals in Workday.
        Find all journals with status: In Progress.
        Extract: Journal ID, Date, Amount, Status.
        Return as a structured list.
    """,

    "check_period_close": """
        Navigate to Accounting > Close Periods in Workday.
        Check the current period close status.
        Return: Period name, Status, and count of open tasks.
    """,

    "view_account_balances": """
        Navigate to Accounting > Account Balances in Workday.
        Select the current fiscal period.
        Extract top-level account balances.
        Return as structured data.
    """,

    # ── PROCUREMENT TASKS ─────────────────────────────────────
    "check_pending_pos": """
        Navigate to Procurement > Purchase Orders in Workday.
        Filter by status: Pending Approval.
        Extract: PO Number, Supplier, Amount, Date.
        Return all pending POs found.
    """,

    "check_supplier_invoices": """
        Navigate to Procurement > Supplier Invoices in Workday.
        Filter by status: Awaiting Action.
        Extract: Invoice ID, Supplier, Amount, Due Date.
        Return list of invoices needing action.
    """,

    # ── ADMIN TASKS ───────────────────────────────────────────
    "check_business_process": """
        Navigate to Business Process configuration in Workday.
        Search for business process: {bp_name}
        Check current configuration and active steps.
        Extract all steps with their conditions and security groups.
        Return complete BP configuration.
    """,

    "navigate_to_integrations": """
        Navigate to Integration > Integration Systems in Workday.
        Search for integration named: {integration_name}
        Check the last run status and any errors.
        Return status and error details if any.
    """,

    # ── SECURITY TASKS ────────────────────────────────────────
    "check_security_group": """
        Navigate to Security > Security Groups in Workday.
        Search for security group: {group_name}
        Extract all members and their assigned permissions.
        Return complete security group details.
    """,

    # ── TENANT MANAGEMENT ────────────────────────────────────
    "check_tenant_settings": """
        Navigate to Administration > Tenant Setup in Workday.
        Extract key configuration settings visible on the page.
        Return as structured configuration data.
    """,

    # ── REPORT VERIFICATION (uses Excel output) ──────────────
    "verify_reports_from_catalog": """
        I have an Excel report catalog at: {excel_path}
        The catalog contains AI-tagged reports organized by industry.
        Navigate to the Workday Reports section.
        For each of these report names: {report_names}
        Verify that the report exists in Workday by searching for it.
        Return the verification status for each report
        (Found / Not Found / Access Denied).
    """,
}


# ─────────────────────────────────────────────────────────────
# SECTION 3: Browser & LLM Setup
# ─────────────────────────────────────────────────────────────
def _create_browser():
    """Create a fresh Browser instance."""
    return Browser(
        config=BrowserProfile(
            headless=BROWSER_CONFIG["headless"],
            disable_security=False,
        )
    )


def _create_llm():
    """Create the Gemini LLM instance."""
    return ChatGoogleGenerativeAI(
        model="gemini-2.0-flash",
        google_api_key=GEMINI_API_KEY,
        temperature=0.0,  # Deterministic for reliable browser automation
    )


# ─────────────────────────────────────────────────────────────
# SECTION 4: Core Browser Task Execution
# ─────────────────────────────────────────────────────────────
async def run_workday_browser_task(task: str) -> dict:
    """
    Execute a defined task on the Workday tenant via browser.

    This function:
    1. Opens a Chromium browser (visible mode by default)
    2. Navigates to the Workday tenant URL
    3. Logs in with configured UI credentials
    4. Performs the given task using LLM-driven automation
    5. Returns structured results

    Args:
        task: Natural language description of the task to perform.

    Returns:
        dict with keys: status, result, task, timestamp
    """
    _validate_config()

    # Prepend login steps to every task
    full_task = f"""
    IMPORTANT: Follow these steps carefully and sequentially.

    STEP 1 — NAVIGATE:
    Go to {BROWSER_CONFIG['tenant_url']}

    STEP 2 — LOGIN:
    On the Workday login page:
    - Enter username: {BROWSER_CONFIG['username']}
    - Enter password: {BROWSER_CONFIG['password']}
    - Click the Sign In / Log In button
    - Wait for the Workday home page to fully load
      (look for the search bar or home dashboard)

    STEP 3 — PERFORM TASK:
    Once logged in and the home page is loaded, perform this task:
    {task}

    STEP 4 — REPORT RESULTS:
    After completing the task, provide a clear summary of:
    - What was done
    - Any data that was extracted
    - Any errors encountered
    """

    browser = _create_browser()
    llm = _create_llm()

    agent = Agent(
        task=full_task,
        llm=llm,
        browser=browser,
    )

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        print(f"\n{'='*70}")
        print(f"Browser Agent Starting — {timestamp}")
        print(f"Task: {task[:100]}{'...' if len(task) > 100 else ''}")
        print(f"Tenant: {BROWSER_CONFIG['tenant_url']}")
        print(f"Headless: {BROWSER_CONFIG['headless']}")
        print(f"{'='*70}\n")

        result = await agent.run(max_steps=25)

        print(f"\n{'='*70}")
        print("Browser Agent — Task Completed Successfully")
        print(f"{'='*70}\n")

        return {
            "status": "success",
            "result": str(result),
            "task": task,
            "timestamp": timestamp,
        }

    except Exception as e:
        print(f"\n{'='*70}")
        print(f"Browser Agent — Error: {e}")
        print(f"{'='*70}\n")

        return {
            "status": "error",
            "error": str(e),
            "task": task,
            "timestamp": timestamp,
        }

    finally:
        await browser.close()


# ─────────────────────────────────────────────────────────────
# SECTION 5: Synchronous Wrapper (Tool Interface)
# ─────────────────────────────────────────────────────────────
def workday_browser_agent(task: str) -> dict:
    """
    Synchronous entry point for the browser agent.

    Can be called as a tool from any orchestrator or directly:
        result = workday_browser_agent("Navigate to reports and list them")

    Args:
        task: Natural language task description.

    Returns:
        dict with status, result/error, task, and timestamp.
    """
    return asyncio.run(run_workday_browser_task(task))


def run_predefined_task(task_key: str, **kwargs) -> dict:
    """
    Run a predefined task from the WORKDAY_TASKS library.

    Args:
        task_key: Key from WORKDAY_TASKS dict (e.g. "navigate_to_reports")
        **kwargs: Template variables to fill in (e.g. report_name="My Report")

    Returns:
        dict with status, result/error, task, and timestamp.

    Example:
        run_predefined_task("run_custom_report", report_name="Employee Details")
    """
    if task_key not in WORKDAY_TASKS:
        available = ", ".join(sorted(WORKDAY_TASKS.keys()))
        return {
            "status": "error",
            "error": f"Unknown task key: '{task_key}'. Available: {available}",
            "task": task_key,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

    task_template = WORKDAY_TASKS[task_key]

    # Fill in template variables like {report_name}
    for key, value in kwargs.items():
        task_template = task_template.replace(f"{{{key}}}", str(value))

    return workday_browser_agent(task_template.strip())


# ─────────────────────────────────────────────────────────────
# SECTION 6: Orchestrated Workflow (Report Agent → Browser Agent)
# ─────────────────────────────────────────────────────────────
def run_full_pipeline(browser_task: str = None, use_sample_data: bool = False) -> dict:
    """
    Full pipeline: Run the report catalog agent first, then the browser agent.

    Step 1: Execute workday_agent.run_agent() to fetch reports and generate Excel.
    Step 2: If an Excel file was produced and a browser_task is provided,
            execute the browser agent with the Excel path context.

    Args:
        browser_task: Optional browser task to run after Excel generation.
                      If None, only runs the report agent.
        use_sample_data: If True, uses sample data for the report agent.

    Returns:
        dict with excel_output path and browser_result (if applicable).
    """
    # Step 1: Run the report catalog agent
    print("\n" + "=" * 70)
    print("PHASE 1: Running Report Catalog Agent (RaaS → Excel)")
    print("=" * 70)

    try:
        from workday_agent import run_agent
        excel_path = run_agent(use_sample_data=use_sample_data)
    except Exception as e:
        return {
            "phase_1_status": "error",
            "phase_1_error": str(e),
            "phase_2_status": "skipped",
        }

    if not excel_path:
        return {
            "phase_1_status": "completed",
            "phase_1_result": "No AI-tagged reports found — no Excel generated.",
            "phase_2_status": "skipped",
            "phase_2_reason": "No Excel output to act on.",
        }

    result = {
        "phase_1_status": "success",
        "excel_output": excel_path,
    }

    # Step 2: Run the browser agent (only if Excel was produced)
    if browser_task:
        print("\n" + "=" * 70)
        print("PHASE 2: Running Browser Agent (Workday UI Automation)")
        print("=" * 70)

        # Inject the Excel path into the task context
        enriched_task = (
            f"Context: An Excel report catalog has been generated at: {excel_path}\n\n"
            f"Task: {browser_task}"
        )

        browser_result = workday_browser_agent(enriched_task)
        result["phase_2_status"] = browser_result["status"]
        result["browser_result"] = browser_result
    else:
        result["phase_2_status"] = "skipped"
        result["phase_2_reason"] = "No browser task provided."

    # Final summary
    print("\n" + "=" * 70)
    print("PIPELINE COMPLETE")
    print(f"  Phase 1 (Report Agent): {result['phase_1_status']}")
    if excel_path:
        print(f"  Excel Output: {os.path.abspath(excel_path)}")
    print(f"  Phase 2 (Browser Agent): {result.get('phase_2_status', 'skipped')}")
    print("=" * 70 + "\n")

    return result


# ─────────────────────────────────────────────────────────────
# SECTION 7: CLI Entrypoint
# ─────────────────────────────────────────────────────────────
def _print_available_tasks():
    """Print all predefined tasks."""
    print("\nAvailable predefined tasks:")
    print("-" * 50)
    for key in sorted(WORKDAY_TASKS.keys()):
        # Show first line of task description
        desc = WORKDAY_TASKS[key].strip().split("\n")[0].strip()
        print(f"  {key:<35} {desc}")
    print("-" * 50)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Workday Browser Agent — CLI Usage:")
        print()
        print("  Run a predefined task:")
        print('    python workday_browser_agent.py --task navigate_to_reports')
        print()
        print("  Run a custom task:")
        print('    python workday_browser_agent.py --custom "Go to Reports and list all custom reports"')
        print()
        print("  Run the full pipeline (Report Agent + Browser Agent):")
        print('    python workday_browser_agent.py --pipeline --task navigate_to_reports')
        print('    python workday_browser_agent.py --pipeline --sample --task navigate_to_reports')
        print()
        print("  List available predefined tasks:")
        print('    python workday_browser_agent.py --list')
        _print_available_tasks()
        sys.exit(0)

    args = sys.argv[1:]

    # --list: show available tasks
    if "--list" in args:
        _print_available_tasks()
        sys.exit(0)

    # --pipeline mode: Report Agent first, then Browser Agent
    if "--pipeline" in args:
        use_sample = "--sample" in args
        task = None

        if "--task" in args:
            idx = args.index("--task")
            if idx + 1 < len(args):
                task_key = args[idx + 1]
                task = WORKDAY_TASKS.get(task_key, task_key)

        elif "--custom" in args:
            idx = args.index("--custom")
            if idx + 1 < len(args):
                task = args[idx + 1]

        result = run_full_pipeline(browser_task=task, use_sample_data=use_sample)
        print(f"\nPipeline result: {result}")
        sys.exit(0)

    # --task: run a predefined task (browser only)
    if "--task" in args:
        idx = args.index("--task")
        if idx + 1 < len(args):
            task_key = args[idx + 1]
            result = run_predefined_task(task_key)
            print(f"\nResult: {result}")
        else:
            print("Error: --task requires a task key argument.")
            _print_available_tasks()
        sys.exit(0)

    # --custom: run a custom task (browser only)
    if "--custom" in args:
        idx = args.index("--custom")
        if idx + 1 < len(args):
            task = args[idx + 1]
            result = workday_browser_agent(task)
            print(f"\nResult: {result}")
        else:
            print("Error: --custom requires a task description string.")
        sys.exit(0)

    print(f"Unknown arguments: {args}")
    print("Run with no arguments to see usage.")
