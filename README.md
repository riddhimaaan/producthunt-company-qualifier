# ProductHunt Company Qualifier — Claude Code Skill

A Claude Code skill that scores and labels companies from a Product Hunt CSV against a B2B SaaS GTM/outreach ICP. Fetches live websites, scores by keyword pattern matching, and produces a shortlist + split CSVs + an Excel workbook ready for Google Sheets.

**Output labels:** `priority_now` · `good_fit` · `maybe` · `skip`

---

## Install

### Mac / Linux

```bash
git clone https://github.com/riddhimaaan/producthunt-company-qualifier \
  ~/.claude/skills/producthunt-company-qualifier

pip install -r ~/.claude/skills/producthunt-company-qualifier/requirements.txt
```

### Windows (PowerShell)

```powershell
git clone https://github.com/riddhimaaan/producthunt-company-qualifier `
  "$env:USERPROFILE\.claude\skills\producthunt-company-qualifier"

pip install -r "$env:USERPROFILE\.claude\skills\producthunt-company-qualifier\requirements.txt"
```

---

## Usage

In Claude Code, type:

```
/producthunt-company-qualifier
```

Claude will ask for your CSV path and run the full qualification pipeline.

---

## Input CSV

Any CSV with these four columns (Product Hunt exports work out of the box):

| Column | Description |
|--------|-------------|
| `url` | Company website URL |
| `name` | Company name |
| `domain` | Domain (e.g. `clay.com`) |
| `tagline` | One-line description |

Additional columns are preserved in the output.

---

## Output Files

All files land in an `outputs/` folder next to where you run the skill:

| File | Contents |
|------|----------|
| `qualified-products-broad.csv` | All scored companies |
| `qualified-products-broad-shortlist.csv` | Top 121 by score |
| `shortlist-review.txt` | Human-readable shortlist preview |
| `qualified-products-curated.csv` | Shortlist with curated labels |
| `qualified-products-priority_now.csv` | Direct GTM/outreach fit |
| `qualified-products-good_fit.csv` | Strong adjacent workflow fit |
| `qualified-products-maybe.csv` | Adjacent but weaker buyer overlap |
| `qualified-products-skip.csv` | Off-ICP or false positive |
| `qualified-products-google-upload.xlsx` | Excel workbook with all tabs |

---

## ICP Profile

The qualification targets B2B SaaS companies in the GTM/outreach stack:

**Tier 1 (worked with):** Smartlead, Instantly, HeyReach, Aimfox, Clay, Amplemarket, Lindy, and more

**Strong fit:** Cold email, LinkedIn automation, outreach sequencing, enrichment, AI SDR, sales intelligence

**Adjacent fit:** Workflow automation, agent platforms, RevOps/CRM sync, browser automation, no-code

See [`references/target-profile.md`](references/target-profile.md) for the full profile.

---

## Requirements

- Python 3.9+
- `requests`, `beautifulsoup4`, `openpyxl` (installed via `requirements.txt`)
- Claude Code CLI

---

## License

MIT
