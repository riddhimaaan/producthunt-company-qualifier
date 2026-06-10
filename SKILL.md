---
name: producthunt-company-qualifier
version: "1.0.0"
description: "Score and label companies from a Product Hunt CSV against a B2B SaaS GTM/outreach ICP. Fetches live websites, scores by keyword patterns, produces shortlist + split CSVs + Excel workbook."
argument-hint: 'producthunt-company-qualifier'
user-invocable: true
allowed-tools: Bash, Read, Write
metadata:
  openclaw:
    emoji: "🏹"
    tags:
      - icp
      - qualification
      - producthunt
      - lead-generation
      - b2b
      - csv
      - outbound
---

# ProductHunt Company Qualifier

## Trigger

Use this skill when the user runs `/producthunt-company-qualifier` or asks to "qualify Product Hunt companies", "score a PH CSV", or "run company qualification on a CSV".

---

## What This Skill Does

- Reads a CSV of companies (Product Hunt export or any CSV with matching schema)
- Fetches each company's live website in parallel (24 workers, 8s timeout)
- Scores by keyword pattern matching against a B2B GTM/outreach ICP
- Applies hardcoded curated labels to known domains
- Writes 8 output files: broad CSV, shortlist CSV, split-by-label CSVs, and an Excel workbook with tabs
- Optionally pushes the workbook to Google Sheets via Composio

**Required input columns:** `url`, `name`, `domain`, `tagline`

**Output labels:** `priority_now`, `good_fit`, `maybe`, `skip`

Read `references/target-profile.md` for the full ICP fit logic.

---

## Step-by-Step Workflow

### STEP 1 — Check dependencies (first run only)

Before the first run, verify Python packages are installed:

```bash
pip show requests beautifulsoup4 openpyxl > /dev/null 2>&1 || pip install requests beautifulsoup4 openpyxl
```

On Windows (PowerShell):
```powershell
pip show requests beautifulsoup4 openpyxl 2>$null; if ($LASTEXITCODE -ne 0) { pip install requests beautifulsoup4 openpyxl }
```

### STEP 2 — Get the input CSV path

Ask the user: **"What is the full path to your CSV file?"**

Confirm the file exists before proceeding. Required columns: `url`, `name`, `domain`, `tagline`.

### STEP 3 — Run the qualification script

Use this Python snippet to invoke the script portably across operating systems:

```python
from pathlib import Path
import subprocess, sys

skill_root = Path.home() / ".claude" / "skills" / "producthunt-company-qualifier"
script = skill_root / "scripts" / "build_company_sheet.py"
out_dir = Path.cwd() / "outputs"
csv_path = Path(r"REPLACE_WITH_USER_CSV_PATH")

subprocess.run(
    [sys.executable, str(script), "--input-csv", str(csv_path), "--output-dir", str(out_dir)],
    check=True
)
```

Replace `REPLACE_WITH_USER_CSV_PATH` with the actual path the user provided.

To customise shortlist size (default 121):

```python
subprocess.run(
    [sys.executable, str(script),
     "--input-csv", str(csv_path),
     "--output-dir", str(out_dir),
     "--shortlist-size", "150"],
    check=True
)
```

### STEP 4 — Confirm output files

After the script finishes, confirm all 8 files exist in the output directory:

- `qualified-products-broad.csv`
- `qualified-products-broad-shortlist.csv`
- `shortlist-review.txt`
- `qualified-products-curated.csv`
- `qualified-products-priority_now.csv`
- `qualified-products-good_fit.csv`
- `qualified-products-maybe.csv`
- `qualified-products-skip.csv`
- `qualified-products-google-upload.xlsx`

Report to the user: how many companies were scored, shortlist size, and a brief label breakdown.

### STEP 5 — Optional: Push to Google Sheets

If the user asks to push the results to Google Sheets, use Composio:

1. Create one spreadsheet
2. Create tabs: `curated_all`, `priority_now`, `good_fit`, `maybe`, `skip`
3. Prefer importing the `.xlsx` workbook if the Composio connection allows it
4. If import is blocked, write rows tab-by-tab from the split CSVs

Only delete local files if the user explicitly asks.

---

## Notes

- The script auto-detects the only CSV in the working directory if `--input-csv` is omitted
- All hardcoded domain labels are embedded in the script; no external config is needed
- Labels may need manual review if the source CSV covers domains outside the original Product Hunt batch
