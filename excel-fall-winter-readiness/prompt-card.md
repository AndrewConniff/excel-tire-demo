# Prompt card — Copilot in Excel (Follow-Along)

**Classification: Public**

## Quick tips

1. Keep workbook in OneDrive/SharePoint with AutoSave on.
2. Ensure data is in an Excel table with headers.
3. Use clear prompt anatomy: goal + table/sheet + output.
4. Verify key numbers before sharing.

## Copy prompts

### 1) Excel trend summary
**Where:** Excel → `Canada-Wide-Winter-Tire-Sales-Weather-2023-2026.xlsx` → table `tblWinterHistory`.

```text
Summarize the top 5 trends in tblWinterHistory across seasons, regions, snowfall, and winter tire sales.
```

### 2) Excel chart prompt
**Where:** Excel → same workbook/table.

```text
Create a chart comparing weekly Winter Tire Sales (CAD) and Snowfall (cm) by region.
```

### 3) Excel risk ranking
**Where:** Excel → same workbook/table.

```text
Which regions appear highest risk for winter tire demand pressure this season, and why?
Return 3 ranked regions with brief evidence.
```

### 4) Excel executive bullets
**Where:** Excel → same workbook/table.

```text
Write 5 executive bullets summarizing the most important findings from this workbook.
```

### 5) Copilot Chat research
**Where:** Copilot Chat.

```text
Research public information themes on winter tire demand patterns in Canada.
Give key demand drivers, weather factors, regional differences, and signals to track this winter.
```

### 6) Copilot Chat starter CSV
**Where:** Copilot Chat.

```text
Create a starter dataset I can paste into Excel as CSV with columns:
Week Ending, Region, Avg Temp (C), Snowfall (cm), Winter Tire Units Sold, Winter Tire Sales (CAD), Notes.
Generate 20 realistic fictional rows in YYYY-MM-DD and CAD conventions.
```

### 7) Excel clean-table prompt
**Where:** Excel → new workbook with pasted CSV.

```text
I pasted a CSV table on sheet ResearchData.
Convert this into a clean Excel table and suggest a useful table name.
```

### 8) Communication handoff
**Where:** Excel Copilot or Copilot Chat with research table context.

```text
Draft:
1) a short Outlook update email, and
2) a short Teams channel post
based on this research table and chart.
```
