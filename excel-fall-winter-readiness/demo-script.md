# Demo script — Copilot in Excel Follow-Along (60 minutes)

**Classification: Public**

## 0–10 min: Setup

1. Open `sample-data/Canada-Wide-Winter-Tire-Sales-Weather-2023-2026.xlsx` in Excel from OneDrive/SharePoint.
2. Confirm AutoSave is on.
3. Open Copilot pane in Excel.
4. Keep Copilot Chat available in a browser tab for the research section.

## 10–40 min: Excel follow-along prompts

### E1 — Trend summary

**Where this happens:** Excel, workbook `Canada-Wide-Winter-Tire-Sales-Weather-2023-2026.xlsx`, sheet `Data`, table `tblWinterHistory`.

```text
Summarize the top 5 trends in tblWinterHistory across seasons, regions, snowfall, and winter tire sales.
```

### E2 — Chart creation

**Where this happens:** Same workbook/context as E1.

```text
Create a chart comparing weekly Winter Tire Sales (CAD) and Snowfall (cm) by region.
Use a layout that makes trend differences easy to explain in 60 seconds.
```

### E3 — Region risk insight

**Where this happens:** Same workbook/context as E1.

```text
Which regions appear highest risk for winter tire demand pressure this season, and why?
Return 3 ranked regions with brief evidence.
```

### E4 — Executive summary bullets

**Where this happens:** Same workbook/context as E1.

```text
Write 5 executive bullets summarizing the most important findings from this workbook.
```

### E5 — Action table

**Where this happens:** Same workbook/context as E1.

```text
Create an action table with columns:
Action | Owner Role | Target Date | Reason.
Use the workbook findings to populate at least 5 rows.
```

## 40–55 min: Chat research → build workbook in Excel

### R1 — Research prompt in Copilot Chat

**Where this happens:** Copilot Chat (browser/app), no workbook required yet.

```text
Research public information themes on winter tire demand patterns in Canada.
Give me:
1) key demand drivers,
2) weather-related factors,
3) regional differences,
4) likely signals to track this winter.
Include concise source references where possible.
```

### R2 — Ask Chat for starter dataset

**Where this happens:** Copilot Chat.

```text
Create a starter dataset I can paste into Excel as CSV with these columns:
Week Ending, Region, Avg Temp (C), Snowfall (cm), Winter Tire Units Sold, Winter Tire Sales (CAD), Notes.
Generate 20 realistic fictional rows in YYYY-MM-DD and CAD conventions.
```

### R3 — Build workbook from Chat output

**Where this happens:** Excel, new blank workbook.

```text
I pasted a CSV table on sheet ResearchData.
Convert this into a clean Excel table and suggest a useful table name.
```

### R4 — Chart + summary from research table

**Where this happens:** Excel, new workbook with pasted table.

```text
Create one chart and one 4-bullet summary from this table that I can share with leadership.
```

### R5 — Communication handoff

**Where this happens:** Excel Copilot or Copilot Chat with table context.

```text
Draft:
1) a short Outlook update email, and
2) a short Teams channel post
based on this research table and chart.
```

## 55–60 min: Close

1. Ask learners which prompt they would use first in their role.
2. Remind learners to verify numbers before sharing outputs.

## MS-4021 style demo sequence (tire-industry adaptation)

Use this sequence when you want to mirror the Microsoft Learning "Copilot in Excel" format.

Source demo page:
https://microsoftlearning.github.io/MS-4021-Copilot-Immersion-Experience/Instructions/Exec/7-copilot-in-excel.html

Source sample file download link:
https://github.com/MicrosoftLearning/MS-4021-Copilot-Immersion-Experience/raw/master/ResourceFiles/EV_Charger_Sales_Analysis_v1.xlsx

1. Open `Canada-Wide-Winter-Tire-Sales-Weather-2023-2026.xlsx` and navigate to `Data`.
2. Prompt:

```text
Calculate monthly Winter Tire Sales (CAD) by Region and Season from this sheet.
Add a helper column with a 4-week rolling sales average and explain the formula used.
```

3. Prompt:

```text
What is the total Winter Tire Sales (CAD) for each region so far this season?
Create a bar chart and add it to a new sheet named "Regional Revenue".
```

4. Prompt:

```text
Highlight rows where Winter Tire Sales (CAD) is less than 120000.
```

5. Import `sample-data/csv/Feedback.csv` into a new sheet named `Reviews`.
6. Prompt:

```text
Summarize the top 3 customer concerns we should focus on from the Reviews sheet.
```

7. Prompt:

```text
Highlight reviews that mention wait times, out-of-stock issues, or installation delays.
```
