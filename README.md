# CBG Production Dashboard

A one-page dashboard for Carib Brewery Grenada: production cases, cases per hour, PM compliance (target 70%), process loss and the four utilities (water, fuel, CO₂, electricity), by YTD, month or date.

It rebuilds itself every time a file is added to `data/raw`, and the **Generate report** button at the top produces a tailored report (Managing Director summary, operations detail, production and process, or utilities) for any year, month, date or date range, ready to print, save as PDF, download or email.

## How it works

1. You upload a raw Excel file into `data/raw`.
2. GitHub Actions runs `scripts/build-data.js`, which reads every file in `data/raw` and creates the dashboard data.
3. The updated site is published on GitHub Pages, usually within 1–2 minutes.

## Adding new data (every month / week)

1. Open the repository on github.com and go into **data → raw**.
2. Click **Add file → Upload files** and drag in the new files.
3. Click **Commit changes**.
4. Wait a minute or two, then refresh the dashboard.

Files it understands automatically:

| File | What it provides |
|---|---|
| `Gross_Efficiency_2026.xlsx` | Cases bottled and available time (cases per hour = cases ÷ available hours) |
| `<Month>_26_Utilities_Tracking.xlsx` | Electricity, water, fuel, CO₂ and production hl (Utility Analysis sheet), plus targets (Daily KPIs sheet) |
| `<Month>_26_Daily_Process_Report_2026.xlsx` | Process Loss (Volume) table: actual, month to date, year to date and target, overall and by brand; brewhouse extract recovery |
| `<Month>_2026.xls` (monthly process loss file with an MTD sheet) | Backup only: FV and BBT volumes, used for process loss if a month has no Daily Process Report |
| `CBG_OEE_..._2026.xlsx` | OEE month to date and year to date (Summary sheet) |
| `FTR_Calculations_2026.xlsx` | FTR month to date and year to date (Summary Report sheet) |
| Any file with **PM** in its name | PMs planned / completed per date (use `templates/PM_Compliance_Template.xlsx`) |

How the loss measures are worked out:

- **Process loss (volume)** comes from the Process Loss (Volume) table in the Daily Process Report: the overall MTD figure for a month, the overall YTD figure for the year, and on a single date the month to date as at that day's report. Target 5.50%.
- **Process loss actual** is (FV + GFE − BBT) ÷ (FV + GFE) for the filtrations in the period, from the same report (for a single date it matches the table's Actual column).
- **Brewing loss** is 100 minus brewhouse extract recovery, from the Daily Process Report (month to date, year to date, or the average of that day's brews). Target 2% (100 − 98%).
- **OEE and FTR** are monthly figures, so they show month to date and year to date only.

Tips:

- **Updating a file you already uploaded** (for example the Gross Efficiency file, which grows all year): upload it with the **same file name**. GitHub replaces the old copy.
- Don't upload duplicate copies such as `Gross_Efficiency_2026 (1).xlsx`.
- **Earlier years (2025)** and later years (2027, 2028) work the same way: upload that year's files to **data/raw** whenever you're ready, with the year in the file name (`Gross_Efficiency_2025.xlsx`, `August_25_Utilities_Tracking.xlsx`, `FTR Calculations 2025.xlsx`, a 2025 Ops KPIs MD report …). They sit alongside the 2026 files, the year appears in the **Year** list, and the 2026 figures are then compared with the same period of 2025. The year in a Gross Efficiency file's name decides which year its sheets belong to.
- To check what was read, open `<your site>/data/build-report.txt`, or the latest run on the **Actions** tab.
- The build never changes `index.html` or `README.md`, so you can upload new versions of them at any time.

## Compare periods

Comparisons open in a **Compare** panel in the middle of the screen, over the dashboard (nothing is added to the page). Close it with **Done**, the **×**, the Esc key or by clicking outside it.

Open it with:
- **Compare with…** next to the period filters: *Previous month*, *Same month last year*, *Another month…* (Date view: *Previous day*, *Same day last year*, *Another date…*; YTD view: *Last year*, *Another year…*). If several months, dates or years are selected on the dashboard, *The selected … side by side* puts them all in the panel.
- The **Compare** button at the top right.

In the panel:
- **Years / Months / Dates**, then the periods as pills. **+ Add month** (or date, year) adds more, up to six; **×** removes one.
- **Quick comparisons**: This month vs last month · Last month vs the month before · This month vs same month last year · This year vs last year.
- The line under the title says how many KPIs got better or worse from the first period to the last. Values are green (on target), amber (within 5%) or red; the Change column says *better* or *worse*. Click a KPI to chart it.
- **Email** opens a new email with the comparison written out; **Print** prints it on one landscape page; **Download PDF** saves it (title, coloured table and chart) as a PDF you can attach; **Copy table** copies it for Excel.
- When the latest period is still running, the **Same days of each month / Same part of each year** switch (on by default) compares like for like. Turn it off to compare whole months.

## Targets

Targets found in the Utilities Tracking and Daily Process Report files are used automatically. To override any of them, edit `data/targets.json` on GitHub (leave `null` to use the file's value). PM compliance is set to 70 and process loss to 5.50.

## First-time setup

1. Create a new repository on github.com (for example `cbg-production-dashboard`).
2. Upload everything in this folder, **including the hidden `.github` folder**.
   (On a Mac press Cmd + Shift + . in Finder to see it. GitHub Desktop is the easiest way to upload the whole folder.)
3. Go to **Settings → Pages**. Under **Build and deployment → Source**, choose **GitHub Actions** (recommended). "Deploy from a branch" (main, / root) also works.
4. Go to the **Actions** tab, open **Update dashboard** and click **Run workflow**.
5. When it finishes, the dashboard is at `https://<your-username>.github.io/<repository-name>/`.

## Privacy

GitHub Pages sites are public on free and Pro plans: anyone with the link can see the dashboard, and in a public repository anyone can download the raw files. Only GitHub Enterprise Cloud can restrict a Pages site to your organisation.

## Running it on your own computer (optional)

Requires Node.js 18 or later.

```
npm ci
node scripts/build-data.js
npx serve _site
```

## Troubleshooting

**"Data could not be loaded (HTTP 404)"**: `data/dashboard-data.json` isn't on the site.
- Check that `data/dashboard-data.json` exists in the repository. If not, upload it into the `data` folder.
- Check the **Actions** tab. If there is no "Update dashboard" workflow, the hidden `.github` folder wasn't uploaded: use **Add file → Create new file**, name it `.github/workflows/update-dashboard.yml` and paste in the contents of that file.
- If a run shows "Permission denied" when saving the data file, go to **Settings → Actions → General → Workflow permissions** and choose **Read and write permissions**.

## Adding new data (short version)

Upload the new or updated files into **data/raw** and commit. That's all: the **Update dashboard** workflow reads every file, rebuilds the data and republishes the site in 1–2 minutes.

- File names don't need to match exactly. The dashboard recognises each file by what's inside it (Gross Efficiency, Utilities Tracking, Daily Process Report, monthly process loss .xls, OEE, FTR, PM compliance).
- If two files cover the same month (for example a renamed copy), the one with the newest data is used.
- **Ops KPIs MD report** (`Ops_KPIs_2026_MD_Report_new_format.xlsx`, one sheet per month such as "New MD September 26") holds the official, corrected monthly figures. For every month it covers, its numbers replace the calculated ones: production cases, cases per hour, brews per day, brewing loss (100 − extract recovery), process loss, utilities, OEE, FTR and PM compliance, and its budgets become the targets. The year to date uses its YTD column, with any later months added on. Add each new month's sheet to the workbook and upload it again under the same name. Single days are still calculated from the daily files.
- **Production cases** = cases bottled + kegs from the Gross Efficiency workbook (1 keg = 2.75 cases). Cases per hour uses cases bottled only.
- **PM compliance** comes from the PM workbook (one sheet per month, e.g. "Sept PM"), using its "% COMPLETION" figure. Add each new month as a new sheet and re-upload the workbook.
- **Brews per day** comes from the Daily Brewing Plan block in the Daily Process Report (average brews per day, MTD / YTD / target).
