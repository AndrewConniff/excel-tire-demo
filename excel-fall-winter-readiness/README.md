# Copilot in Excel Follow-Along Lab

**Classification: Public**

This package is now a **follow-along** lab for commercial Microsoft 365 Copilot users.  
Participants can run prompts in their own Excel/Chat context without preloaded tenant conversations.

## Primary workbook for follow-along

- `sample-data/Canada-Wide-Winter-Tire-Sales-Weather-2023-2026.xlsx`

Use this workbook for:
1. trend summaries
2. chart generation
3. weather-vs-sales comparisons
4. email/Teams-ready insight bullets

## Secondary workbook (legacy scenario)

- `sample-data/Fall_Winter_Readiness_2026.xlsx`

Use only if you want the original table-driven scenario with `tblSales`, `tblInventory`, and `tblService`.

## Chat-to-Excel research section (new)

The lab includes prompts for:
1. researching tire-sales signals in Canada in Copilot Chat
2. generating a starter tabular dataset from Chat output
3. building a new workbook from that research in Excel
4. turning results into charts and communication-ready summaries

## What to open

- `index.html` for the full runbook (all instructions in one place)
- `DEMO-INSTRUCTIONS.md` for the markdown presenter guide
- `DEMO-INSTRUCTIONS.docx` for the Word presenter guide
- optional reference docs:
  - `demo-script.md`
  - `prompt-card.md`
  - `presenter-setup.md`

## Added: MS-4021 style demo-steps tab

`index.html` now includes an **MS-4021 style demo steps** tab adapted to the auto parts and tire scenario.

It includes:
1. a source reference to the Microsoft Learning Copilot in Excel flow
2. a direct download link to the source sample file:
   - `https://github.com/MicrosoftLearning/MS-4021-Copilot-Immersion-Experience/raw/master/ResourceFiles/EV_Charger_Sales_Analysis_v1.xlsx`
3. a tire-industry equivalent sequence (M1-M7) for revenue and customer review analysis

## Regeneration (existing generator assets)

If you need to regenerate the original synthetic scenario files:

```powershell
python scripts/generate_workbook.py
python scripts/render_assets.py
```

This demo uses fictional/sample data and does not contain customer confidential information.
