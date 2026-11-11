# DEMO-INSTRUCTIONS - Copilot in Excel (Auto Parts and Tire Scenario)

**Classification: Public**

This guide mirrors the Microsoft Learning MS-4021 Copilot in Excel flow and adapts it to the Canada-wide tire and auto-service readiness scenario.

## Source reference

- Microsoft Learning page: https://microsoftlearning.github.io/MS-4021-Copilot-Immersion-Experience/Instructions/Exec/7-copilot-in-excel.html
- Source sample file download: https://github.com/MicrosoftLearning/MS-4021-Copilot-Immersion-Experience/raw/master/ResourceFiles/EV_Charger_Sales_Analysis_v1.xlsx

## Demo workbook

- Primary workbook for this scenario: `sample-data/Canada-Wide-Winter-Tire-Sales-Weather-2023-2026.xlsx`
- Optional review input file: `sample-data/csv/Feedback.csv`

## Demo steps (MS-4021 style, adapted)

1. Open `Canada-Wide-Winter-Tire-Sales-Weather-2023-2026.xlsx` in Excel and navigate to the `Data` sheet.
2. Run this prompt in Excel Copilot:

```text
Calculate monthly Winter Tire Sales (CAD) by Region and Season from this sheet.
Add a helper column with a 4-week rolling sales average and explain the formula used.
```

3. Run this prompt:

```text
What is the total Winter Tire Sales (CAD) for each region so far this season?
Create a bar chart and add it to a new sheet named "Regional Revenue".
```

4. Run this prompt:

```text
Highlight rows where Winter Tire Sales (CAD) is less than 120000.
```

5. Import `sample-data/csv/Feedback.csv` into a new sheet named `Reviews`.
6. Run this prompt in the `Reviews` sheet:

```text
Summarize the top 3 customer concerns we should focus on from the Reviews sheet.
```

7. Run this prompt:

```text
Highlight reviews that mention wait times, out-of-stock issues, or installation delays.
```

## Wrap-up

Use the resulting revenue analysis and customer-feedback insights to brief regional leaders on inventory, staffing, and service-readiness actions for the upcoming winter cycle.
