"""
Quick diagnostic: fetch the live JSON and inspect field names + Report_Tag values.
"""
import sys, json
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Reuse our existing modules
from api.raas_client import fetch_workday_report

data = fetch_workday_report()

# Get the report list
if "Report_Entry" in data:
    reports = data["Report_Entry"]
else:
    reports = list(data.values())[0] if data else []

print(f"\nTotal reports: {len(reports)}")
print(f"\n{'='*60}")
print("FIELD NAMES in first report:")
print("="*60)
if reports:
    for key in sorted(reports[0].keys()):
        val = str(reports[0][key])[:80]
        print(f"  {key:<45} {val}")

# Check for Report_Tag or any tag-like field
print(f"\n{'='*60}")
print("REPORT_TAG ANALYSIS:")
print("="*60)
tag_field_candidates = []
for key in (reports[0].keys() if reports else []):
    if "tag" in key.lower():
        tag_field_candidates.append(key)

if tag_field_candidates:
    print(f"Tag-like fields found: {tag_field_candidates}")
    for field in tag_field_candidates:
        values = set()
        for r in reports:
            v = r.get(field, "")
            if v:
                values.add(str(v)[:60])
        print(f"\n  {field} — {len(values)} unique non-empty values:")
        for v in sorted(values)[:20]:
            print(f"    • {v}")
else:
    print("No field containing 'tag' found in report keys.")
    # Check Report_Name for AI_ prefixed names
    ai_names = [r.get("Report_Name","") for r in reports if "AI_" in str(r.get("Report_Name",""))]
    print(f"\nReports with 'AI_' in Report_Name: {len(ai_names)}")
    for n in ai_names[:10]:
        print(f"  • {n}")

# Show all unique Report_Tag values (if field exists)
if reports and "Report_Tag" in reports[0]:
    all_tags = set(str(r.get("Report_Tag","")) for r in reports if r.get("Report_Tag"))
    print(f"\nAll unique Report_Tag values ({len(all_tags)}):")
    for t in sorted(all_tags)[:30]:
        print(f"  • {t}")
