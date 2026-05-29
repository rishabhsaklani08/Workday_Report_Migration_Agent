"""
raas_client.py — Workday RaaS API client.

Handles ISU-authenticated GET requests to a Workday RaaS endpoint
(customreport2 or Reporting_UI) with retry logic, timeout handling,
and structured error messages.
"""

import time
import logging
import requests
from requests.auth import HTTPBasicAuth

from config import WORKDAY_CONFIG

logger = logging.getLogger(__name__)


def fetch_workday_report() -> dict:
    """
    Fetch report data from Workday via ISU RaaS call.

    Uses the direct raas_url from config (supports both customreport2
    and Reporting_UI endpoint patterns).

    Returns:
        dict: Parsed JSON response from Workday.

    Raises:
        Exception: On authentication, permission, network, or unexpected errors.
    """
    url = WORKDAY_CONFIG["raas_url"]
    username = WORKDAY_CONFIG["isu_username"]
    password = WORKDAY_CONFIG["isu_password"]
    timeout = WORKDAY_CONFIG["timeout_seconds"]
    max_retries = WORKDAY_CONFIG["retry_attempts"]
    retry_delay = WORKDAY_CONFIG["retry_delay_seconds"]

    params = {"format": "json"}
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    last_exception = None

    for attempt in range(1, max_retries + 1):
        try:
            logger.info(
                f"📡 RaaS API call attempt {attempt}/{max_retries} → {url}"
            )

            response = requests.get(
                url,
                auth=HTTPBasicAuth(username, password),
                params=params,
                headers=headers,
                timeout=timeout,
                verify=True,  # SSL verification
            )

            # ── Handle HTTP status codes ──────────────────────
            if response.status_code == 200:
                logger.info("✅ RaaS call successful")
                return response.json()

            elif response.status_code == 401:
                raise Exception(
                    "Authentication failed (401). "
                    "Check ISU credentials in .env file."
                )

            elif response.status_code == 403:
                raise Exception(
                    "Access denied (403). "
                    "Check ISU permissions for this RaaS report."
                )

            elif response.status_code == 404:
                raise Exception(
                    "Report not found (404). "
                    "Verify the WORKDAY_RAAS_URL in .env."
                )

            elif response.status_code == 429:
                # Rate-limited — wait and retry
                wait = retry_delay * attempt
                logger.warning(
                    f"⏳ Rate limited (429). Retrying in {wait}s..."
                )
                time.sleep(wait)
                continue

            else:
                raise Exception(
                    f"API Error: {response.status_code} — "
                    f"{response.text[:500]}"
                )

        except requests.exceptions.ConnectionError as exc:
            last_exception = exc
            logger.error(
                f"❌ Connection failed (attempt {attempt}): {exc}"
            )
            if attempt < max_retries:
                time.sleep(retry_delay)

        except requests.exceptions.Timeout as exc:
            last_exception = exc
            logger.error(
                f"❌ Request timed out after {timeout}s (attempt {attempt})"
            )
            if attempt < max_retries:
                time.sleep(retry_delay)

        except Exception:
            # Re-raise non-retryable errors immediately
            raise

    # All retries exhausted
    raise Exception(
        f"All {max_retries} API attempts failed. "
        f"Last error: {last_exception}"
    )
