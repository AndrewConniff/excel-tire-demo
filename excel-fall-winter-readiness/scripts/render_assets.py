from __future__ import annotations

from pathlib import Path

from config import COMPANY_NAME


ROOT = Path(__file__).resolve().parents[1]


def write(path: Path, text: str) -> None:
    path.write_text(text.strip() + "\n", encoding="utf-8")


def render_readme() -> None:
    write(
        ROOT / "README.md",
        f"""
# Analyze with Microsoft 365 Copilot in Excel — Fall/Winter Readiness

**Classification: Public**

This package provides a GitHub-ready, demo-led training set for a 60-minute VILT session:
**"Analyze with Microsoft 365 Copilot in Excel"**.

- `COMPANY_NAME` source: `scripts/config.py`
- Default `COMPANY_NAME`: **{COMPANY_NAME}**
- Audience: new/light Copilot users from store operations and corporate office
- Locale: en-CA, CAD, metric units, YYYY-MM-DD dates

This demo uses a fictional Microsoft Fake Company created for demonstration purposes. Any resemblance to real organizations, people, products, services, or data is coincidental. Do not use customer confidential information in this package.

## Purpose

Model a **fall-to-winter readiness review** as of Monday 2026-11-09 across a national retailer with department/home, hardware, and automotive service businesses.

## COMPANY_NAME switch

All generated assets are built from `scripts/config.py`:

```python
COMPANY_NAME = "Canadian Tire"
```

To change to a fictional public name (for example `Contoso Home & Auto`), update `COMPANY_NAME` and regenerate.

## Regenerate package

From `excel-fall-winter-readiness/`:

```powershell
python scripts/generate_workbook.py
python scripts/render_assets.py
```

This regenerates:
- `sample-data/Fall_Winter_Readiness_2026.xlsx`
- `sample-data/csv/*.csv`
- `answer-key.md`
- `sample-data/answer-key-metrics.json`
- `README.md`, `demo-script.md`, `prompt-card.md`, `presenter-setup.md`, `index.html`

## Folder tree

```text
excel-fall-winter-readiness/
  README.md
  demo-script.md
  answer-key.md
  prompt-card.md
  presenter-setup.md
  index.html
  scripts/
    config.py
    generate_workbook.py
    render_assets.py
  sample-data/
    Fall_Winter_Readiness_2026.xlsx
    answer-key-metrics.json
    csv/
      Stores.csv
      Sales.csv
      AutoService.csv
      Inventory.csv
      Targets.csv
      Feedback.csv
```
""",
    )


def render_demo_script() -> None:
    write(
        ROOT / "demo-script.md",
        f"""
# Demo script — Analyze with Microsoft 365 Copilot in Excel (60 minutes)

**Classification: Public**  
**Company context:** {COMPANY_NAME} (fictional demo data)  
**Date context:** Monday 2026-11-09  
**Storyline:** Fall-to-winter readiness review before winter tire rush and holiday season

## 0–5 min: Welcome + storyline

Talk track:
- "Today we are reviewing 13 weeks ending Saturdays, from 2026-08-15 to 2026-11-07."
- "We’ll use Copilot in Excel to find trends, risks, and executive-ready updates."
- "Everything in this file is invented for training."

## 5–15 min: Copilot in Excel essentials

- Open workbook from **OneDrive/SharePoint** with **AutoSave on**.
- Ensure data is in a proper **Excel table**.
- Prompt anatomy for data tasks:
  1. Goal
  2. Table and columns
  3. Output type (chart, highlight, column, summary)
- Always verify Copilot output against visible numbers.

## 15–50 min: Five demos (~7 min each)

### 1) Summarize trends (skill: summarize trends)

**Click path:** Open `Sales` sheet → click any cell in `tblSales` → Copilot pane.  
**Prompt:**
```text
Using tblSales, summarize the top 5 trends over the last 13 weeks.
Focus on category and department movement by week and by region.
```
**Follow-up prompt:**
```text
Create a line chart of weekly Net Sales (CAD) by Department using Week Ending on the x-axis.
```
**Expected result:** Copilot surfaces winter tire acceleration and department-level weekly movement; chart shows Automotive lift.  
**What to say:** "We’re seeing the winter ramp emerge clearly, especially in Automotive."  
**Fallback if different:** Filter to `Category = Winter Tires`, then re-ask trend summary by region.

### 2) Highlight key data points (skill: highlight key data points)

**Click path:** `Inventory` sheet → select `tblInventory`.  
**Prompt:**
```text
In tblInventory, highlight rows where Category is Winter Tires and Weeks of Supply is below 1.5.
Then sort these rows by Next Delivery Date ascending.
```
**Follow-up prompt:**
```text
Which stores in that highlighted set are not restocked before 2026-11-20?
```
**Expected result:** Copilot flags winter tire low-WOS stores and identifies delayed replenishment risk.  
**What to say:** "This is the shortlist for immediate inventory action."  
**Fallback:** Manually filter `Category = Winter Tires` and `Weeks of Supply < 1.5`, then ask Copilot to summarize visible rows.

### 3) Retrieve insights (skill: retrieve insights)

**Click path:** `AutoService` sheet → select `tblService`.  
**Prompt:**
```text
From tblService, show Bay Utilization % by Region by Week Ending.
Create either a PivotTable or chart that makes region changes easy to compare.
```
**Follow-up prompt:**
```text
Which region is over capacity, and when did that start?
```
**Expected result:** Prairies emerges as the capacity hotspot from late October, with wait-time pressure.  
**What to say:** "Capacity pressure starts before peak season fully lands, so this is a staffing trigger."  
**Fallback:** Build Pivot manually (Rows: Week Ending, Columns: Region, Values: Avg Bay Utilization %) then ask Copilot to interpret.

### 4) Explain meaning (skill: explain meaning)

**Click path:** `Sales` + `Targets` sheets.  
**Prompt:**
```text
Compare 13-week actual Net Sales (CAD) from tblSales to 13-Week Net Sales Target (CAD) in tblTargets by Region and Department.
Create a summary with Actual, Target, and Variance %.
```
**Follow-up prompt:**
```text
Explain the variance formula in plain language and identify which areas are behind target and by how much.
```
**Expected result:** Seasonal shortfalls in BC/Ontario, Automotive ahead nationally, Fixing near target.  
**What to say:** "Copilot can compute and explain performance gaps in plain language for mixed audiences."  
**Fallback:** Ask Copilot to produce a new summary table first, then request explanation only.

### 5) Corporate-office wrap (skill: explain and summarize)

**Click path:** `Feedback` sheet, then workbook-level context in Copilot pane.  
**Prompt:**
```text
From tblFeedback, list the top 3 Auto Service feedback themes and include short example quotes.
```
**Follow-up prompt:**
```text
Write 5 bullets summarizing this workbook for a regional operations call.
```
**Expected result:** Themes around wait times, booking confusion, and technician praise; concise call-ready wrap.  
**What to say:** "This is the bridge from analysis to communication in Teams or Outlook."  
**Fallback:** Filter Channel = Auto Service and re-ask for themes from visible data only.

## 50–60 min: Recap, prompt card, Q&A

- Recap four skills: summarize trends, highlight key data points, retrieve insights, explain meaning.
- Share `prompt-card.md`.
- Optional learner practice card if hands-on approval is granted.

## Backup demos

1. **Flyer margin check:**  
   Prompt: "Find unusual Gross Margin % behavior in flyer weeks and identify outliers."
2. **Copilot Chat handoff:**  
   Prompt: "Using this workbook, draft an email to Quebec store managers about winter tire stock risk and next actions."
""",
    )


def render_prompt_card() -> None:
    write(
        ROOT / "prompt-card.md",
        f"""
# Prompt card — Copilot in Excel (en-CA)

**Classification: Public**  
**Company context:** {COMPANY_NAME}

## Quick tips

1. Keep data in Excel tables.
2. Open workbook from OneDrive/SharePoint with AutoSave on.
3. Ask for explicit outputs (chart, table, highlights, bullets).
4. Verify numbers before sharing.

## 8 copy/paste prompts

1. ```text
Summarize the top trends in tblSales over the last 13 weeks by category, department, and region.
```
2. ```text
Create a line chart of weekly Net Sales (CAD) by Department from tblSales.
```
3. ```text
In tblInventory, highlight Winter Tires rows where Weeks of Supply is below 1.5 and sort by Next Delivery Date.
```
4. ```text
From the highlighted inventory rows, which stores are not restocked before 2026-11-20?
```
5. ```text
Using tblService, show Bay Utilization % by Region by Week Ending and explain where capacity pressure starts.
```
6. ```text
Compare actual Net Sales to tblTargets by Region and Department with a Variance % column.
```
7. ```text
Explain the variance formula in plain language for non-technical operations leaders.
```
8. ```text
From tblFeedback, summarize the top 3 Auto Service themes and write 5 bullets for a regional ops call.
```

## Optional learner practice card (if hands-on is approved)

Practice task:
1. Find one inventory risk and one service-capacity risk for a region.
2. Build one chart and one short summary paragraph.
3. Share one action recommendation with owner + date.
""",
    )


def render_presenter_setup() -> None:
    write(
        ROOT / "presenter-setup.md",
        f"""
# Presenter setup

**Classification: Public**  
**Company context:** {COMPANY_NAME}

1. Upload `sample-data/Fall_Winter_Readiness_2026.xlsx` to OneDrive.
2. Create one clean workbook copy per session (for example):
   - `Fall_Winter_Readiness_2026_Tue.xlsx`
   - `Fall_Winter_Readiness_2026_Wed.xlsx`
   - `Fall_Winter_Readiness_2026_Thu.xlsx`
3. Keep AutoSave on during demo. Copilot may insert columns, formulas, highlights, and charts.
4. Run all planned prompts once the day before delivery.
5. Screenshot expected charts/tables and save as fallback visuals.
6. Keep `answer-key.md` and `sample-data/answer-key-metrics.json` open for fast validation.
7. If Copilot returns a different framing, use the fallback prompts in `demo-script.md`.
""",
    )


def render_html() -> None:
    # Excel green + complementary purple for copy buttons.
    write(
        ROOT / "index.html",
        f"""
<!doctype html>
<html lang="en-CA">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Analyze with Copilot in Excel - Demo Guide</title>
  <style>
    :root {{
      --excel-green: #217346;
      --copy-complement: #7a3ea1;
      --ink: #1f2937;
      --panel: #f3f7f4;
    }}
    body {{ margin: 0; font-family: "Segoe UI", Arial, sans-serif; color: var(--ink); }}
    header {{ background: var(--excel-green); color: #fff; padding: 24px; }}
    header h1 {{ margin: 0 0 6px; }}
    main {{ max-width: 1100px; margin: 0 auto; padding: 22px; }}
    .meta {{ background: var(--panel); border-left: 6px solid var(--excel-green); padding: 12px 14px; margin-bottom: 16px; }}
    .links a {{ display: inline-block; margin: 6px 10px 6px 0; color: #0b57d0; text-decoration: none; font-weight: 600; }}
    .box {{ border: 1px solid #dbe5df; border-radius: 8px; margin: 12px 0; overflow: hidden; }}
    .boxhead {{ background: #edf5f0; padding: 10px 12px; display: flex; justify-content: space-between; align-items: center; }}
    pre {{ margin: 0; padding: 12px; white-space: pre-wrap; font-family: Consolas, "Courier New", monospace; font-size: 0.9rem; }}
    button {{ border: none; background: var(--copy-complement); color: #fff; padding: 7px 10px; border-radius: 6px; cursor: pointer; }}
    h2 {{ margin-top: 28px; }}
    .small {{ color: #4b5563; font-size: 0.92rem; }}
  </style>
</head>
<body>
  <header>
    <h1>Analyze with Microsoft 365 Copilot in Excel</h1>
    <div>60-minute VILT demo guide | Company context: {COMPANY_NAME} | en-CA, CAD, metric</div>
  </header>
  <main>
    <div class="meta">
      <strong>Instructor instructions:</strong> Open workbook from OneDrive with AutoSave on, keep Copilot pane visible, and verify outputs against table numbers before narrating conclusions.
      <br><br>
      Storyline: It's Monday 2026-11-09. Regional operations is reviewing 13 weeks ending 2026-11-07 before winter tire and holiday peak.
    </div>

    <div class="links">
      <a href="./README.md">README.md</a>
      <a href="./demo-script.md">demo-script.md</a>
      <a href="./answer-key.md">answer-key.md</a>
      <a href="./prompt-card.md">prompt-card.md</a>
      <a href="./presenter-setup.md">presenter-setup.md</a>
      <a href="./sample-data/Fall_Winter_Readiness_2026.xlsx">Workbook (.xlsx)</a>
    </div>

    <h2>Step-by-step run order</h2>
    <ol>
      <li>Open workbook from OneDrive and confirm AutoSave is on.</li>
      <li>Start on <strong>Sales</strong> table for trend summary and chart.</li>
      <li>Move to <strong>Inventory</strong> for stock-risk highlights and delivery cutoff check.</li>
      <li>Move to <strong>AutoService</strong> for capacity and wait-time insight by region/week.</li>
      <li>Use <strong>Sales + Targets</strong> for actual-vs-target variance explanation.</li>
      <li>Close with <strong>Feedback + workbook wrap</strong> for call-ready summary bullets.</li>
    </ol>

    <h2>Copy-ready prompts</h2>
    <div class="box"><div class="boxhead"><strong>Demo 1: Summarize trends</strong><button onclick="copyPrompt('p1')">Copy</button></div><pre id="p1">Using tblSales, summarize the top 5 trends over the last 13 weeks.
Focus on category and department movement by week and by region.</pre></div>
    <div class="box"><div class="boxhead"><strong>Demo 1 follow-up</strong><button onclick="copyPrompt('p1b')">Copy</button></div><pre id="p1b">Create a line chart of weekly Net Sales (CAD) by Department using Week Ending on the x-axis.</pre></div>
    <div class="box"><div class="boxhead"><strong>Demo 2: Highlight key data points</strong><button onclick="copyPrompt('p2')">Copy</button></div><pre id="p2">In tblInventory, highlight rows where Category is Winter Tires and Weeks of Supply is below 1.5.
Then sort these rows by Next Delivery Date ascending.</pre></div>
    <div class="box"><div class="boxhead"><strong>Demo 2 follow-up</strong><button onclick="copyPrompt('p2b')">Copy</button></div><pre id="p2b">Which stores in that highlighted set are not restocked before 2026-11-20?</pre></div>
    <div class="box"><div class="boxhead"><strong>Demo 3: Retrieve insights</strong><button onclick="copyPrompt('p3')">Copy</button></div><pre id="p3">From tblService, show Bay Utilization % by Region by Week Ending.
Create either a PivotTable or chart that makes region changes easy to compare.</pre></div>
    <div class="box"><div class="boxhead"><strong>Demo 3 follow-up</strong><button onclick="copyPrompt('p3b')">Copy</button></div><pre id="p3b">Which region is over capacity, and when did that start?</pre></div>
    <div class="box"><div class="boxhead"><strong>Demo 4: Explain meaning</strong><button onclick="copyPrompt('p4')">Copy</button></div><pre id="p4">Compare 13-week actual Net Sales (CAD) from tblSales to 13-Week Net Sales Target (CAD) in tblTargets by Region and Department.
Create a summary with Actual, Target, and Variance %.</pre></div>
    <div class="box"><div class="boxhead"><strong>Demo 4 follow-up</strong><button onclick="copyPrompt('p4b')">Copy</button></div><pre id="p4b">Explain the variance formula in plain language and identify which areas are behind target and by how much.</pre></div>
    <div class="box"><div class="boxhead"><strong>Demo 5: Corporate wrap</strong><button onclick="copyPrompt('p5')">Copy</button></div><pre id="p5">From tblFeedback, list the top 3 Auto Service feedback themes and include short example quotes.</pre></div>
    <div class="box"><div class="boxhead"><strong>Demo 5 follow-up</strong><button onclick="copyPrompt('p5b')">Copy</button></div><pre id="p5b">Write 5 bullets summarizing this workbook for a regional operations call.</pre></div>

    <h2>Backup prompts</h2>
    <div class="box"><div class="boxhead"><strong>Backup A</strong><button onclick="copyPrompt('b1')">Copy</button></div><pre id="b1">Find unusual Gross Margin % in flyer weeks and identify category/week/store outliers.</pre></div>
    <div class="box"><div class="boxhead"><strong>Backup B</strong><button onclick="copyPrompt('b2')">Copy</button></div><pre id="b2">Using this workbook, draft an email to Quebec store managers about winter tire stock risk and next actions.</pre></div>

    <p class="small">Tip: If Copilot output differs, use the fallback approach in <code>demo-script.md</code> and validate with <code>answer-key.md</code>.</p>
  </main>
  <script>
    async function copyPrompt(id) {{
      await navigator.clipboard.writeText(document.getElementById(id).innerText);
      alert("Copied.");
    }}
  </script>
</body>
</html>
""",
    )


def main() -> None:
    render_readme()
    render_demo_script()
    render_prompt_card()
    render_presenter_setup()
    render_html()
    print("Rendered markdown/html assets.")


if __name__ == "__main__":
    main()
