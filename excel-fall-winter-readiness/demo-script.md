# Demo script — Analyze with Microsoft 365 Copilot in Excel (60 minutes)

**Classification: Public**  
**Company context:** Canadian Tire (fictional demo data)  
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
**File/context in use:** `sample-data/Fall_Winter_Readiness_2026.xlsx` (`Sales` table view).  
**Alternative:** In Copilot Chat, attach the workbook and run the same prompt.
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
**File/context in use:** same workbook, `Inventory` sheet only.  
**Alternative:** In Copilot Chat, attach workbook or `sample-data/csv/Inventory.csv`.
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
**File/context in use:** same workbook, `AutoService` sheet only.  
**Alternative:** In Copilot Chat, attach workbook or `sample-data/csv/AutoService.csv`.
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
**File/context in use:** same workbook with both `tblSales` and `tblTargets`.  
**Alternative:** In Copilot Chat, attach workbook and ask for an aggregated summary table.
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
**File/context in use:** same workbook, `Feedback` table plus workbook context.  
**Alternative:** In Copilot Chat, attach workbook or `sample-data/csv/Feedback.csv`.
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
