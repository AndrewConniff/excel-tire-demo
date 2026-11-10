# Analyze with Microsoft 365 Copilot in Excel — Fall/Winter Readiness

**Classification: Public**

This package provides a GitHub-ready, demo-led training set for a 60-minute VILT session:
**"Analyze with Microsoft 365 Copilot in Excel"**.

- `COMPANY_NAME` source: `scripts/config.py`
- Default `COMPANY_NAME`: **Canadian Tire**
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

## Where to run prompts (Excel vs Chat)

| Demo step | Run in Excel (preferred) | Open this first | Copilot Chat alternative |
|---|---|---|---|
| 1. Summarize trends | Yes | `Fall_Winter_Readiness_2026.xlsx` → `Sales` sheet (`tblSales`) | Attach workbook in Chat and ask same prompt |
| 2. Highlight stock risk | Yes | `Inventory` sheet (`tblInventory`) | Attach workbook or `Inventory.csv` in Chat |
| 3. Retrieve service insights | Yes | `AutoService` sheet (`tblService`) | Attach workbook or `AutoService.csv` in Chat |
| 4. Explain target variance | Yes | `Sales` + `Targets` sheets | Attach workbook, ask for summarized variance table |
| 5. Corporate wrap | Yes | `Feedback` sheet, then workbook context | Attach workbook or `Feedback.csv` in Chat |

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
