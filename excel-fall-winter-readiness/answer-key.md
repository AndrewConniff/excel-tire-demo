# Answer key — Fall-to-winter readiness patterns

**Classification: Public**

Company context from `COMPANY_NAME`: **Canadian Tire**

All values below are calculated by `scripts/generate_workbook.py` from generated data.

## Pattern A — Winter Tires trend
- Last 4 weeks top category by Net Sales: **Winter Tires**
- Winter Tires ramp ratio (late 4-week avg vs early 5-week avg) by region:
  - Atlantic: **6.6x**
  - BC: **5.83x**
  - Ontario: **6.33x**
  - Prairies: **6.5x**
  - Quebec: **7.31x**

## Pattern B — Inventory stock risk / overstock
- Winter Tire risk rows (WOS < 1.0 and delivery after 2026-11-20): **4**
  - Store 205 (Quebec), Winter Tires - Core, WOS 0.72, next delivery 2026-11-24
  - Store 212 (Quebec), Winter Tires - Core, WOS 0.65, next delivery 2026-11-27
  - Store 318 (Ontario), Winter Tires - Core, WOS 0.8, next delivery 2026-11-22
  - Store 417 (Prairies), Winter Tires - Core, WOS 0.45, next delivery 2026-11-22
- Ontario All-Season overstock rows (WOS > 10): **2**
  - Store 305 All-Season Tires - Core, WOS 13.25
  - Store 327 All-Season Tires - Core, WOS 12.91

## Pattern C — Auto Service capacity
- Prairies first week average utilization above 95%: **2026-10-24**
- Tire Changeover share of all jobs (Oct-Nov): **71.19%**
- Prairies rows above 100% utilization: **6 shown (top 10 list in metrics file)**

## Pattern D — Target gaps
- Seasonal variance vs target (Actual/Target - 1):
  - BC: **-13.0%**
  - Ontario: **-13.0%**
- Automotive national variance: **6.38%**
- Fixing national variance: **-0.1%**

## Pattern E — Flyer effect and outlier
- Flyer unit lift vs non-flyer: **26.61%**
- Flyer Gross Margin % delta: **-5.0 points**
- Power Tools negative margin rows on 2026-10-03: **2**
  - Store 314: GM% -5.14%, Net Sales $5981.11, Gross Margin $-307.15
  - Store 323: GM% -3.76%, Net Sales $6174.89, Gross Margin $-232.12

## Pattern F — Feedback themes
- Comments mentioning waits: **60**
- Comments mentioning booking confusion: **51**
- Comments praising technicians/friendly service: **61**
- Prairies Auto Service comments rated 1-2: **43**

## Verification source
- Machine-readable metrics: `sample-data/answer-key-metrics.json`
