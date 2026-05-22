# South Dakota Campaign Finance – Weekly Committee Scraper

Automatically queries [sdcfr.sdsos.gov](https://sdcfr.sdsos.gov/Search/Search.aspx) every Monday, extracts all 1,600+ registered committees, and exports a formatted Excel report with full contact information.

## What it collects

| Column | Description |
|---|---|
| Committee Name | Full registered name |
| Committee Type | Legislative / PAC / Candidate / Ballot Question / County Party / etc. |
| Candidate | Candidate name (if applicable) |
| Campaign Status | Active / Terminated / Suspended |
| Committee Address | Physical address |
| Committee Telephone | Main phone |
| Committee Website | Website URL |
| Committee Chair | Chair full name |
| Chair Phone | Chair daytime phone |
| Chair Email | Chair email (decoded from Cloudflare obfuscation) |
| Committee Treasurer | Treasurer full name |
| Treasurer Phone | Treasurer daytime phone |
| Treasurer Email | Treasurer email |
| Detail URL | Link to the committee's page on sdcfr.sdsos.gov |

## Schedule

Runs every **Monday at 12:00 UTC** until **September 1, 2026**, then auto-exits.

You can also trigger a run manually from the **Actions** tab → **SD Campaign Finance – Weekly Scraper** → **Run workflow**.

## Downloading reports

1. Go to the **Actions** tab in this repo  
2. Click on the latest successful workflow run  
3. Scroll to **Artifacts** and click **sd-committees-report-XXXXXX** to download a ZIP  
4. Unzip to get `SD_Committees_YYYY-MM-DD.xlsx` and a summary `.txt` file  

Reports are retained for **90 days**.

## How it works

```
sd_weekly_scraper.py
  Phase 1 — Playwright (headless Chromium):
    • Navigates to the Search page
    • Selects "Search By Committee Name" → "All"
    • Pages through all 164 pages via the "Next Page" button
    • Collects every committee's detail-page URL (cid + rid)

  Phase 2 — requests (20 parallel workers):
    • Fetches each committee's detail page
    • Decodes Cloudflare email obfuscation (XOR algorithm)
    • Extracts phone numbers and email addresses

  Phase 3 — openpyxl:
    • Writes formatted Excel file with alternating row colors
    • Adds a Summary sheet with type breakdown
    • Saves to ./sd_reports/SD_Committees_YYYY-MM-DD.xlsx
```

## Local development

```bash
pip install -r requirements.txt
playwright install chromium
python sd_weekly_scraper.py
```

Output will be written to `./sd_reports/`.
