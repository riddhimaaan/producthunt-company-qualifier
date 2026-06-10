# ProductHunt Company Qualifier — Claude Code Skill

Score and label companies from any Product Hunt CSV against **your own** B2B ICP. Fetches live websites, runs keyword pattern scoring, and produces a ranked shortlist + split CSVs + an Excel workbook ready for Google Sheets.

**Output labels:** `priority_now` · `good_fit` · `maybe` · `skip`

---

## Prerequisites

Before installing, make sure you have these three things:

| Requirement | Check |
|-------------|-------|
| [Claude Code](https://claude.ai/code) installed | Run `claude --version` in your terminal |
| Python 3.9 or newer | Run `python --version` or `python3 --version` |
| Git | Run `git --version` |

---

## Installation (do this once)

### Mac / Linux — open Terminal and run these two commands:

**Step 1 — Download the skill:**
```bash
git clone https://github.com/riddhimaaan/producthunt-company-qualifier \
  ~/.claude/skills/producthunt-company-qualifier
```

**Step 2 — Install Python dependencies:**
```bash
pip install -r ~/.claude/skills/producthunt-company-qualifier/requirements.txt
```

---

### Windows — open PowerShell and run these two commands:

**Step 1 — Download the skill:**
```powershell
git clone https://github.com/riddhimaaan/producthunt-company-qualifier "$env:USERPROFILE\.claude\skills\producthunt-company-qualifier"
```

**Step 2 — Install Python dependencies:**
```powershell
pip install -r "$env:USERPROFILE\.claude\skills\producthunt-company-qualifier\requirements.txt"
```

> **Note:** If `pip` gives an error, try `pip3` instead.

---

## First-time setup (happens automatically)

The first time you run the skill, Claude will walk you through a **one-time setup** to personalize the scoring for your ICP. It will ask you:

1. Which companies have you worked with or sold to? (your Tier 1 / Tier 2 clients)
2. What are the strongest signals that a company is a great fit for you?
3. What are adjacent or secondary signals?
4. How would you describe the four outcome buckets for your leads?

Claude saves your answers and the skill is now **tuned to your ICP** — not anyone else's.

You only do this once. After that it goes straight to processing.

---

## How to use it

**Step 1 — Get your CSV ready**

Your CSV needs these four columns (Product Hunt exports work out of the box):

| Column | Description | Example |
|--------|-------------|---------|
| `url` | Company website URL | `https://clay.com` |
| `name` | Company name | `Clay` |
| `domain` | Domain only | `clay.com` |
| `tagline` | One-line description | `Enrich data and automate outreach` |

Any extra columns you have will be kept in the output unchanged.

**Step 2 — Open Claude Code**

Open your terminal and start Claude Code in any folder:
```bash
claude
```

**Step 3 — Run the skill**

Type this command:
```
/producthunt-company-qualifier
```

Claude will ask for the path to your CSV file. Paste the full path, for example:
- Mac/Linux: `/Users/yourname/Downloads/producthunt-export.csv`
- Windows: `C:\Users\yourname\Downloads\producthunt-export.csv`

**Step 4 — Wait for it to finish**

The script fetches each company's live website in parallel (up to 24 at a time). A typical run of 500 companies takes about 1–2 minutes.

**Step 5 — Find your output files**

All output files are saved in an `outputs/` folder inside whichever directory you ran Claude Code from:

```
outputs/
├── qualified-products-broad.csv             ← Every company, scored
├── qualified-products-broad-shortlist.csv   ← Top 121 by score
├── shortlist-review.txt                     ← Human-readable preview
├── qualified-products-curated.csv           ← Shortlist with your labels
├── qualified-products-priority_now.csv      ← Direct ICP fit
├── qualified-products-good_fit.csv          ← Strong adjacent fit
├── qualified-products-maybe.csv             ← Weaker overlap, worth a look
├── qualified-products-skip.csv              ← Off-ICP, disqualified
└── qualified-products-google-upload.xlsx    ← Excel workbook, all tabs
```

**The file to start with:** open `qualified-products-curated.csv` — it has the shortlist ranked and split by your four labels.

---

## Optional: Push results to Google Sheets

After the run, tell Claude: *"push the results to Google Sheets"* — it will use Composio to create a spreadsheet with separate tabs for each label.

---

## Customizing your ICP later

Your scoring profile is saved at:
- **Mac/Linux:** `~/.claude/skills/producthunt-company-qualifier/references/scoring-config.json`
- **Windows:** `%USERPROFILE%\.claude\skills\producthunt-company-qualifier\references\scoring-config.json`

You can edit this JSON file at any time to add new companies you've worked with, adjust scoring weights, or update your heuristic keywords. The human-readable version is in `references/target-profile.md` in the same folder.

---

## Troubleshooting

**"command not found: claude"**
Claude Code isn't installed or isn't in your PATH. [Download it here](https://claude.ai/code).

**"No module named 'requests'"**
The Python dependencies weren't installed. Re-run Step 2 from the installation section above.

**"Missing required columns: domain, name, tagline, url"**
Your CSV is missing one of the four required columns. Check the column names match exactly (lowercase, no spaces).

**The skill doesn't appear when I type `/producthunt-company-qualifier`**
The skill folder might be in the wrong location. Verify the folder exists at:
- Mac/Linux: `~/.claude/skills/producthunt-company-qualifier/SKILL.md`
- Windows: `%USERPROFILE%\.claude\skills\producthunt-company-qualifier\SKILL.md`

---

## Requirements

- Claude Code CLI
- Python 3.9+
- `requests`, `beautifulsoup4`, `openpyxl` (installed via Step 2 above)
- Git

---

## License

MIT
