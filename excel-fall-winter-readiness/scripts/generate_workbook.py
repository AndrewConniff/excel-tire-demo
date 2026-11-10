from __future__ import annotations

import json
import random
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from config import COMPANY_NAME, RANDOM_SEED


ROOT = Path(__file__).resolve().parents[1]
SAMPLE_DATA = ROOT / "sample-data"
CSV_DIR = SAMPLE_DATA / "csv"
WORKBOOK_PATH = SAMPLE_DATA / "Fall_Winter_Readiness_2026.xlsx"
ANSWER_KEY_PATH = ROOT / "answer-key.md"
METRICS_JSON_PATH = SAMPLE_DATA / "answer-key-metrics.json"


def week_endings() -> List[date]:
    start = date(2026, 8, 15)
    return [start + timedelta(days=7 * i) for i in range(13)]


@dataclass
class StoreRecord:
    store_id: int
    store_name: str
    city: str
    province: str
    region: str
    fmt: str
    bays: int
    manager: str


def build_stores(rng: random.Random) -> pd.DataFrame:
    rows: List[StoreRecord] = []
    manager_first = [
        "Avery",
        "Jordan",
        "Riley",
        "Casey",
        "Morgan",
        "Taylor",
        "Cameron",
        "Harper",
        "Dakota",
        "Quinn",
        "Parker",
        "Rowan",
    ]
    manager_last = [
        "Lavoie",
        "Chow",
        "Mercer",
        "Bains",
        "Nguyen",
        "Lemieux",
        "Singh",
        "Bouchard",
        "Fraser",
        "Morrison",
        "Khan",
        "Delorme",
    ]

    store_blueprint = [
        (101, "Halifax North", "Halifax", "NS", "Atlantic"),
        (104, "Moncton West", "Moncton", "NB", "Atlantic"),
        (109, "St. John's East", "St. John's", "NL", "Atlantic"),
        (205, "Montreal Saint-Laurent", "Montreal", "QC", "Quebec"),
        (208, "Quebec City Beauport", "Quebec City", "QC", "Quebec"),
        (212, "Laval Chomedey", "Laval", "QC", "Quebec"),
        (216, "Sherbrooke Centre", "Sherbrooke", "QC", "Quebec"),
        (219, "Gatineau Hull", "Gatineau", "QC", "Quebec"),
        (301, "Toronto East York", "Toronto", "ON", "Ontario"),
        (305, "Mississauga Erin Mills", "Mississauga", "ON", "Ontario"),
        (309, "Ottawa South Keys", "Ottawa", "ON", "Ontario"),
        (314, "Hamilton Mountain", "Hamilton", "ON", "Ontario"),
        (318, "London North", "London", "ON", "Ontario"),
        (323, "Kitchener South", "Kitchener", "ON", "Ontario"),
        (327, "Sudbury South", "Sudbury", "ON", "Ontario"),
        (332, "Kingston East", "Kingston", "ON", "Ontario"),
        (401, "Winnipeg St. James", "Winnipeg", "MB", "Prairies"),
        (406, "Regina South", "Regina", "SK", "Prairies"),
        (412, "Red Deer South", "Red Deer", "AB", "Prairies"),
        (417, "Calgary Seton", "Calgary", "AB", "Prairies"),
        (422, "Saskatoon Stonebridge", "Saskatoon", "SK", "Prairies"),
        (501, "Vancouver Cambie", "Vancouver", "BC", "BC"),
        (506, "Kelowna North", "Kelowna", "BC", "BC"),
        (511, "Prince George West", "Prince George", "BC", "BC"),
    ]

    for store_id, title, city, prov, region in store_blueprint:
        fmt = rng.choices(["Large", "Medium", "Small"], weights=[0.35, 0.45, 0.2], k=1)[0]
        bays = {"Large": rng.randint(7, 10), "Medium": rng.randint(5, 7), "Small": rng.randint(3, 5)}[fmt]
        mgr = f"{rng.choice(manager_first)} {rng.choice(manager_last)}"
        rows.append(
            StoreRecord(
                store_id=store_id,
                store_name=f"Store {store_id} – {title}",
                city=city,
                province=prov,
                region=region,
                fmt=fmt,
                bays=bays,
                manager=mgr,
            )
        )

    return pd.DataFrame(
        [
            {
                "Store ID": r.store_id,
                "Store Name": r.store_name,
                "City": r.city,
                "Province": r.province,
                "Region": r.region,
                "Format": r.fmt,
                "Auto Service Bays": r.bays,
                "Store Manager": r.manager,
            }
            for r in rows
        ]
    )


def sales_categories() -> Dict[str, List[str]]:
    return {
        "Automotive": ["Winter Tires", "All-Season Tires", "Auto Parts & Batteries", "Car Care"],
        "Fixing": ["Power Tools", "Hardware & Paint"],
        "Living": ["Kitchen & Home"],
        "Playing": ["Outdoor & Sports"],
        "Seasonal": ["Holiday Décor", "Snow Removal"],
    }


def winter_ramp_multiplier(region: str, week_idx: int) -> float:
    base = [0.56, 0.58, 0.60, 0.64, 0.72, 0.84, 1.02, 1.28, 1.55, 1.88, 2.22, 2.58, 2.92][week_idx]
    region_early = {
        "Quebec": [1.06, 1.12, 1.20, 1.30, 1.42, 1.56, 1.72, 1.86, 2.00, 2.10, 2.18, 2.24, 2.30],
        "Atlantic": [0.98, 1.00, 1.05, 1.10, 1.18, 1.26, 1.36, 1.46, 1.56, 1.64, 1.70, 1.76, 1.80],
        "Ontario": [0.96, 0.98, 1.00, 1.04, 1.08, 1.14, 1.22, 1.30, 1.38, 1.46, 1.52, 1.58, 1.62],
        "Prairies": [0.95, 0.96, 0.97, 0.98, 1.00, 1.04, 1.12, 1.24, 1.36, 1.45, 1.52, 1.58, 1.62],
        "BC": [0.92, 0.94, 0.96, 0.98, 1.01, 1.06, 1.12, 1.18, 1.24, 1.30, 1.36, 1.42, 1.46],
    }
    return base * region_early[region][week_idx]


def build_sales(stores: pd.DataFrame, rng: random.Random) -> pd.DataFrame:
    weeks = week_endings()
    category_map = sales_categories()
    categories = [(dept, cat) for dept, cat_list in category_map.items() for cat in cat_list]
    assert len(categories) == 10

    base_units = {
        "Winter Tires": 18,
        "All-Season Tires": 22,
        "Auto Parts & Batteries": 42,
        "Car Care": 55,
        "Power Tools": 28,
        "Hardware & Paint": 64,
        "Kitchen & Home": 36,
        "Outdoor & Sports": 26,
        "Holiday Décor": 34,
        "Snow Removal": 16,
    }
    unit_price = {
        "Winter Tires": 182.0,
        "All-Season Tires": 146.0,
        "Auto Parts & Batteries": 64.0,
        "Car Care": 22.0,
        "Power Tools": 198.0,
        "Hardware & Paint": 38.0,
        "Kitchen & Home": 54.0,
        "Outdoor & Sports": 71.0,
        "Holiday Décor": 34.0,
        "Snow Removal": 88.0,
    }
    base_margin = {
        "Winter Tires": 0.24,
        "All-Season Tires": 0.26,
        "Auto Parts & Batteries": 0.33,
        "Car Care": 0.36,
        "Power Tools": 0.29,
        "Hardware & Paint": 0.34,
        "Kitchen & Home": 0.35,
        "Outdoor & Sports": 0.31,
        "Holiday Décor": 0.39,
        "Snow Removal": 0.28,
    }
    region_size = {"Atlantic": 0.86, "Quebec": 1.03, "Ontario": 1.11, "Prairies": 0.95, "BC": 0.90}
    flyer_weeks = {date(2026, 9, 5), date(2026, 9, 26), date(2026, 10, 17), date(2026, 11, 7)}
    outlier_stores = {314, 323}

    rows = []
    for _, store in stores.iterrows():
        for i, wk in enumerate(weeks):
            fiscal_week = f"FW2026-{33 + i:02d}"
            flyer = "Y" if wk in flyer_weeks else "N"
            for dept, cat in categories:
                units = base_units[cat] * region_size[store["Region"]]
                if cat == "Winter Tires":
                    units *= winter_ramp_multiplier(store["Region"], i)
                elif cat == "All-Season Tires":
                    units *= max(0.42, 1.12 - i * 0.055)
                elif cat == "Snow Removal":
                    snow_ramp = [0.38, 0.42, 0.46, 0.54, 0.60, 0.66, 0.71, 0.75, 0.80, 0.86, 0.92, 0.98, 1.04][i]
                    if store["Region"] in {"BC", "Ontario"}:
                        snow_ramp *= 0.72
                    units *= snow_ramp
                elif cat == "Holiday Décor":
                    holiday = [0.48, 0.52, 0.58, 0.66, 0.74, 0.88, 1.00, 1.15, 1.28, 1.42, 1.55, 1.68, 1.80][i]
                    units *= holiday
                else:
                    units *= (0.94 + rng.random() * 0.16)

                if flyer == "Y":
                    units *= 1.19
                units *= 0.94 + rng.random() * 0.12
                units_int = max(1, int(round(units)))

                net_sales = units_int * unit_price[cat] * (0.97 + rng.random() * 0.08)
                margin_pct = base_margin[cat] + (rng.random() - 0.5) * 0.03
                if flyer == "Y":
                    margin_pct -= 0.05
                if wk == date(2026, 10, 3) and cat == "Power Tools" and int(store["Store ID"]) in outlier_stores:
                    margin_pct = -0.035 - rng.random() * 0.02
                gross_margin = net_sales * margin_pct
                online_share = {
                    "Winter Tires": 0.14,
                    "All-Season Tires": 0.16,
                    "Auto Parts & Batteries": 0.22,
                    "Car Care": 0.24,
                    "Power Tools": 0.19,
                    "Hardware & Paint": 0.15,
                    "Kitchen & Home": 0.31,
                    "Outdoor & Sports": 0.28,
                    "Holiday Décor": 0.26,
                    "Snow Removal": 0.18,
                }[cat] + (rng.random() - 0.5) * 0.05

                rows.append(
                    {
                        "Week Ending": wk,
                        "Fiscal Week": fiscal_week,
                        "Store ID": int(store["Store ID"]),
                        "Region": store["Region"],
                        "Province": store["Province"],
                        "Department": dept,
                        "Category": cat,
                        "Units": units_int,
                        "Net Sales (CAD)": round(net_sales, 2),
                        "Gross Margin (CAD)": round(gross_margin, 2),
                        "Gross Margin %": round(margin_pct, 4),
                        "Flyer Week (Y/N)": flyer,
                        "Online Share %": round(max(0.02, min(0.5, online_share)), 4),
                    }
                )

    df = pd.DataFrame(rows)
    assert len(df) == 3120
    return df


def build_service(stores: pd.DataFrame, rng: random.Random) -> pd.DataFrame:
    weeks = week_endings()
    service_types = [
        "Tire Changeover",
        "Tire Purchase & Install",
        "Oil Change",
        "Battery & Electrical",
        "Brakes",
    ]
    service_price = {
        "Tire Changeover": 119.0,
        "Tire Purchase & Install": 298.0,
        "Oil Change": 79.0,
        "Battery & Electrical": 142.0,
        "Brakes": 319.0,
    }
    rows = []
    for _, store in stores.iterrows():
        store_id = int(store["Store ID"])
        region = store["Region"]
        bays = int(store["Auto Service Bays"])
        for wk in weeks:
            avail = bays * 48
            if wk >= date(2026, 10, 3):
                shares = [0.71, 0.13, 0.07, 0.05, 0.04]
            else:
                shares = [0.40, 0.20, 0.19, 0.12, 0.09]
            total_jobs = int(round((bays * 22) * (0.9 + rng.random() * 0.28)))

            if region == "Prairies" and wk >= date(2026, 10, 24):
                util = 0.96 + rng.random() * 0.07
                wait_days = 7.1 + rng.random() * 1.8
                if store_id == 412 and wk >= date(2026, 10, 31):
                    util = 1.01 + rng.random() * 0.06
            else:
                util = 0.74 + rng.random() * 0.19
                wait_days = 1.8 + rng.random() * 1.7

            used = avail * util
            type_jobs = [max(1, int(round(total_jobs * s * (0.92 + rng.random() * 0.18)))) for s in shares]

            for service_type, jobs in zip(service_types, type_jobs):
                revenue = jobs * service_price[service_type] * (0.95 + rng.random() * 0.12)
                rating = 3.8 + rng.random() * 1.1
                if region == "Prairies" and wk >= date(2026, 10, 24):
                    rating -= 0.35 + rng.random() * 0.25
                rows.append(
                    {
                        "Week Ending": wk,
                        "Store ID": store_id,
                        "Region": region,
                        "Service Type": service_type,
                        "Jobs Completed": jobs,
                        "Bay Hours Available": round(avail, 1),
                        "Bay Hours Used": round(used, 1),
                        "Bay Utilization %": round(used / avail, 4),
                        "Avg Wait (days)": round(wait_days, 2),
                        "Revenue (CAD)": round(revenue, 2),
                        "Avg Customer Rating (1-5)": round(max(1.0, min(5.0, rating)), 2),
                    }
                )

    df = pd.DataFrame(rows)
    assert len(df) == 1560
    return df


def build_inventory(stores: pd.DataFrame, rng: random.Random) -> pd.DataFrame:
    products = [
        ("Winter Tires - Core", "Winter Tires"),
        ("Winter Tires - Premium", "Winter Tires"),
        ("All-Season Tires - Core", "All-Season Tires"),
        ("All-Season Tires - Touring", "All-Season Tires"),
        ("Battery 94R", "Auto Parts & Batteries"),
        ("Battery 35", "Auto Parts & Batteries"),
        ("Windshield Fluid 4L", "Car Care"),
        ("Wiper Blades Pair", "Car Care"),
        ("Snow Shovel Steel", "Snow Removal"),
        ("Ice Melt 10kg", "Snow Removal"),
        ("Holiday Light Set", "Holiday Décor"),
        ("Storage Tote 90L", "Kitchen & Home"),
    ]

    at_risk = {205, 212, 318, 417}
    overstock_all_season = {305, 327}
    rows = []
    for _, store in stores.iterrows():
        store_id = int(store["Store ID"])
        region = store["Region"]
        for product, category in products:
            avg_sold = max(3, int(round(8 + rng.random() * 42)))
            on_hand = int(round(avg_sold * (1.5 + rng.random() * 5.5)))
            on_order = int(round(avg_sold * (0.4 + rng.random() * 2.6)))
            next_delivery = date(2026, 11, 10) + timedelta(days=rng.randint(0, 12))

            if product == "Winter Tires - Core" and store_id in at_risk:
                avg_sold = int(round(42 + rng.random() * 18))
                on_hand = int(round(avg_sold * (0.45 + rng.random() * 0.35)))
                on_order = int(round(avg_sold * (0.2 + rng.random() * 0.5)))
                next_delivery = date(2026, 11, 21) + timedelta(days=rng.randint(0, 7))

            if product == "All-Season Tires - Core" and store_id in overstock_all_season:
                avg_sold = int(round(7 + rng.random() * 4))
                on_hand = int(round(avg_sold * (10.8 + rng.random() * 2.6)))
                on_order = int(round(avg_sold * (2.0 + rng.random() * 1.8)))
                next_delivery = date(2026, 11, 12) + timedelta(days=rng.randint(0, 4))

            weeks_supply = round(on_hand / avg_sold, 2)
            rows.append(
                {
                    "Store ID": store_id,
                    "Region": region,
                    "Product": product,
                    "Category": category,
                    "On Hand Units": on_hand,
                    "On Order Units": on_order,
                    "Avg Weekly Units Sold (last 4 wks)": avg_sold,
                    "Weeks of Supply": weeks_supply,
                    "Next Delivery Date": next_delivery,
                    "Supplier": f"{rng.choice(['NorthPeak', 'TrueLine', 'PrairieRoute', 'AtlasRoad', 'MapleSupply'])} Wholesale",
                }
            )

    df = pd.DataFrame(rows)
    assert len(df) == 288
    return df


def build_feedback(stores: pd.DataFrame, rng: random.Random) -> pd.DataFrame:
    store_region = {int(r["Store ID"]): r["Region"] for _, r in stores.iterrows()}
    region_stores: Dict[str, List[int]] = {}
    for sid, region in store_region.items():
        region_stores.setdefault(region, []).append(sid)

    themes = (
        ["auto_wait"] * 72
        + ["booking_confusing"] * 34
        + ["tech_praise"] * 34
        + ["winter_oos"] * 40
        + ["checkout_lines"] * 34
        + ["holiday_praise"] * 26
    )
    assert len(themes) == 240
    rng.shuffle(themes)

    rows = []
    start = date(2026, 8, 15)
    for i, theme in enumerate(themes):
        d = start + timedelta(days=rng.randint(0, 86))
        if theme == "auto_wait":
            region = rng.choices(["Prairies", "Ontario", "Quebec", "Atlantic", "BC"], weights=[0.52, 0.20, 0.14, 0.08, 0.06], k=1)[0]
            channel = "Auto Service"
            rating = rng.choice([1, 1, 2, 2, 2, 3])
            comment = rng.choice(
                [
                    "Had to wait over a week for a tire appointment after the first snowfall. Staff were polite but the delay was tough.",
                    "Tire change booking took longer than expected and the next slot was several days out.",
                    "Auto Service team was friendly, but wait times for tire work are still too long right now.",
                ]
            )
        elif theme == "booking_confusing":
            region = rng.choices(["Prairies", "Ontario", "Quebec", "Atlantic", "BC"], weights=[0.44, 0.24, 0.14, 0.10, 0.08], k=1)[0]
            channel = "Auto Service"
            rating = rng.choice([2, 2, 3, 3, 3, 4])
            comment = rng.choice(
                [
                    "Online booking showed one time, then the confirmation email showed another.",
                    "The booking flow is confusing when switching between tire purchase and install-only options.",
                    "I had trouble understanding which appointments included tire installation versus inspection.",
                ]
            )
        elif theme == "tech_praise":
            region = rng.choice(list(region_stores.keys()))
            channel = "Auto Service"
            rating = rng.choice([4, 4, 5, 5, 5])
            comment = rng.choice(
                [
                    "Technicians explained everything clearly and got me back on the road quickly.",
                    "Friendly Auto Service team and clear advice on winter tires.",
                    "Great service from the tech team, very helpful and professional.",
                ]
            )
        elif theme == "winter_oos":
            region = rng.choices(["Quebec", "Ontario", "Prairies", "Atlantic", "BC"], weights=[0.30, 0.30, 0.18, 0.12, 0.10], k=1)[0]
            channel = rng.choice(["In-store", "Online"])
            rating = rng.choice([1, 2, 2, 3])
            comment = rng.choice(
                [
                    "Winter tires I needed were out of stock and next delivery was later than expected.",
                    "Could not find the right winter tire size in store and online stock looked limited too.",
                    "Inventory for winter tires seems tight right now; had to check another location.",
                ]
            )
        elif theme == "checkout_lines":
            region = rng.choice(list(region_stores.keys()))
            channel = "In-store"
            rating = rng.choice([2, 3, 3, 4])
            comment = rng.choice(
                [
                    "Checkout lines were long during flyer weekend, but staff kept things moving.",
                    "Busy flyer day with long checkout waits near noon.",
                    "Lineups were noticeable this weekend; more tills open would help.",
                ]
            )
        else:
            region = rng.choice(list(region_stores.keys()))
            channel = "In-store"
            rating = rng.choice([4, 5, 5, 5])
            comment = rng.choice(
                [
                    "Holiday displays looked excellent and made finding gifts easier.",
                    "Really liked the holiday setup and signage this week.",
                    "Seasonal display quality was strong and the store felt well organized.",
                ]
            )

        store_id = rng.choice(region_stores[region])
        rows.append(
            {
                "Date": d,
                "Store ID": store_id,
                "Region": region,
                "Channel": channel,
                "Rating (1-5)": rating,
                "Comment": comment,
            }
        )
    return pd.DataFrame(rows)


def build_targets(sales: pd.DataFrame) -> pd.DataFrame:
    agg = sales.groupby(["Region", "Department"], as_index=False)["Net Sales (CAD)"].sum()
    rows = []
    for _, r in agg.iterrows():
        region = r["Region"]
        dept = r["Department"]
        actual = float(r["Net Sales (CAD)"])
        if dept == "Seasonal" and region in {"BC", "Ontario"}:
            target = actual / 0.87
        elif dept == "Automotive":
            target = actual * 0.94
        elif dept == "Fixing":
            target = actual * 1.001
        elif dept == "Living":
            target = actual * 0.99
        else:
            target = actual * 1.01
        rows.append(
            {
                "Region": region,
                "Department": dept,
                "13-Week Net Sales Target (CAD)": round(target, 2),
            }
        )
    df = pd.DataFrame(rows).sort_values(["Region", "Department"]).reset_index(drop=True)
    assert len(df) == 25
    return df


def write_about_sheet(ws, company_name: str) -> None:
    ws.title = "About"
    ws["A1"] = "Fall-to-winter readiness review workbook"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = f"Company context: {company_name} (demo training data, fully fictional)."
    ws["A3"] = "Period covered: week-ending Saturdays 2026-08-15 through 2026-11-07."
    ws["A4"] = "All values use en-CA conventions, CAD currency, metric units, and YYYY-MM-DD dates."
    ws["A6"] = "Tables"
    ws["A6"].font = Font(bold=True)
    lines = [
        "tblStores: 24 stores with region, format, auto service bay counts, and manager names.",
        "tblSales: 13-week weekly sales by store, department, and category with margin and flyer flags.",
        "tblService: auto service jobs, bay utilization, wait time, and revenue by week and service type.",
        "tblInventory: 2026-11-07 inventory snapshot including weeks of supply and next delivery date.",
        "tblTargets: 13-week regional department targets for net sales.",
        "tblFeedback: short customer comments with channel and rating.",
    ]
    for i, line in enumerate(lines, start=7):
        ws[f"A{i}"] = f"- {line}"
    ws.column_dimensions["A"].width = 130


def dataframe_to_sheet_with_table(wb: Workbook, sheet_name: str, table_name: str, df: pd.DataFrame) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.freeze_panes = "A2"
    ws.append(df.columns.tolist())
    for row in df.itertuples(index=False):
        ws.append(list(row))

    end_col = get_column_letter(df.shape[1])
    end_row = df.shape[0] + 1
    table_ref = f"A1:{end_col}{end_row}"
    tab = Table(displayName=table_name, ref=table_ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    # Widths and formats
    for i, col in enumerate(df.columns, start=1):
        values = [str(v) for v in df[col].head(80)]
        max_len = max([len(col)] + [len(v) for v in values])
        ws.column_dimensions[get_column_letter(i)].width = min(max(11, max_len + 2), 42)

        if "Date" in col or "Week Ending" in col:
            for r in range(2, end_row + 1):
                ws.cell(row=r, column=i).number_format = "yyyy-mm-dd"
        elif "(CAD)" in col:
            for r in range(2, end_row + 1):
                ws.cell(row=r, column=i).number_format = '"$"#,##0.00'
        elif "%" in col:
            for r in range(2, end_row + 1):
                ws.cell(row=r, column=i).number_format = "0.00%"
        elif "Rating" in col:
            for r in range(2, end_row + 1):
                ws.cell(row=r, column=i).number_format = "0.00"
        elif any(k in col for k in ["Units", "Jobs", "Hours", "Week", "ID"]):
            for r in range(2, end_row + 1):
                ws.cell(row=r, column=i).number_format = "#,##0"


def write_workbook(dfs: Dict[str, Tuple[str, pd.DataFrame]]) -> None:
    wb = Workbook()
    ws = wb.active
    write_about_sheet(ws, COMPANY_NAME)

    for sheet_name, (table_name, df) in dfs.items():
        dataframe_to_sheet_with_table(wb, sheet_name, table_name, df)
    wb.save(WORKBOOK_PATH)


def export_csvs(dfs: Dict[str, Tuple[str, pd.DataFrame]]) -> None:
    CSV_DIR.mkdir(parents=True, exist_ok=True)
    for sheet_name, (_, df) in dfs.items():
        out = CSV_DIR / f"{sheet_name}.csv"
        df.to_csv(out, index=False, date_format="%Y-%m-%d")


def compute_metrics(
    stores: pd.DataFrame,
    sales: pd.DataFrame,
    service: pd.DataFrame,
    inventory: pd.DataFrame,
    targets: pd.DataFrame,
    feedback: pd.DataFrame,
) -> Dict[str, object]:
    last4 = sorted(sales["Week Ending"].unique())[-4:]
    last4_sales = (
        sales[sales["Week Ending"].isin(last4)]
        .groupby("Category", as_index=False)["Net Sales (CAD)"]
        .sum()
        .sort_values("Net Sales (CAD)", ascending=False)
    )
    top_last4 = last4_sales.iloc[0]["Category"]

    winter_region = (
        sales[sales["Category"] == "Winter Tires"]
        .groupby(["Region", "Week Ending"], as_index=False)["Net Sales (CAD)"]
        .sum()
    )
    early_weeks = sorted(sales["Week Ending"].unique())[:5]
    late_weeks = sorted(sales["Week Ending"].unique())[-4:]
    region_ramp = {}
    for region in winter_region["Region"].unique():
        early = winter_region[(winter_region["Region"] == region) & (winter_region["Week Ending"].isin(early_weeks))][
            "Net Sales (CAD)"
        ].mean()
        late = winter_region[(winter_region["Region"] == region) & (winter_region["Week Ending"].isin(late_weeks))][
            "Net Sales (CAD)"
        ].mean()
        region_ramp[region] = round(late / early, 2)

    stock_risk = inventory[
        (inventory["Category"] == "Winter Tires")
        & (inventory["Weeks of Supply"] < 1.0)
        & (inventory["Next Delivery Date"] > date(2026, 11, 20))
    ][["Store ID", "Region", "Product", "Weeks of Supply", "Next Delivery Date"]].sort_values("Store ID")

    all_season_over = inventory[
        (inventory["Category"] == "All-Season Tires")
        & (inventory["Region"] == "Ontario")
        & (inventory["Weeks of Supply"] > 10.0)
    ][["Store ID", "Product", "Weeks of Supply"]].sort_values("Store ID")

    svc_week = service.groupby(["Region", "Week Ending"], as_index=False)["Bay Utilization %"].mean()
    prairies = svc_week[svc_week["Region"] == "Prairies"].sort_values("Week Ending")
    prairies_over95 = prairies[prairies["Bay Utilization %"] > 0.95]
    over_capacity_store = (
        service[(service["Region"] == "Prairies") & (service["Bay Utilization %"] > 1.0)][
            ["Store ID", "Week Ending", "Bay Utilization %"]
        ]
        .drop_duplicates()
        .sort_values(["Week Ending", "Store ID"])
    )

    oct_nov = service[service["Week Ending"] >= date(2026, 10, 3)]
    tire_jobs = oct_nov[oct_nov["Service Type"] == "Tire Changeover"]["Jobs Completed"].sum()
    all_jobs = oct_nov["Jobs Completed"].sum()
    tire_share = tire_jobs / all_jobs

    actual = sales.groupby(["Region", "Department"], as_index=False)["Net Sales (CAD)"].sum()
    merged = actual.merge(targets, on=["Region", "Department"])
    merged["Variance %"] = merged["Net Sales (CAD)"] / merged["13-Week Net Sales Target (CAD)"] - 1
    seasonal_gap = merged[(merged["Department"] == "Seasonal") & (merged["Region"].isin(["BC", "Ontario"]))][
        ["Region", "Variance %"]
    ]
    automotive_national_var = (
        merged[merged["Department"] == "Automotive"]["Net Sales (CAD)"].sum()
        / merged[merged["Department"] == "Automotive"]["13-Week Net Sales Target (CAD)"].sum()
        - 1
    )
    fixing_national_var = (
        merged[merged["Department"] == "Fixing"]["Net Sales (CAD)"].sum()
        / merged[merged["Department"] == "Fixing"]["13-Week Net Sales Target (CAD)"].sum()
        - 1
    )

    flyer = sales[sales["Flyer Week (Y/N)"] == "Y"]
    non_flyer = sales[sales["Flyer Week (Y/N)"] == "N"]
    units_lift = flyer["Units"].mean() / non_flyer["Units"].mean() - 1
    gm_delta_pts = (flyer["Gross Margin %"].mean() - non_flyer["Gross Margin %"].mean()) * 100
    margin_outliers = sales[
        (sales["Category"] == "Power Tools")
        & (sales["Week Ending"] == date(2026, 10, 3))
        & (sales["Gross Margin %"] < 0)
    ][["Store ID", "Gross Margin %", "Net Sales (CAD)", "Gross Margin (CAD)"]].sort_values("Store ID")

    fb = feedback.copy()
    fb["lc"] = fb["Comment"].str.lower()
    wait_mentions = fb[fb["lc"].str.contains("wait", regex=False)]
    booking_mentions = fb[fb["lc"].str.contains("booking", regex=False)]
    tech_praise = fb[fb["lc"].str.contains("friendly") | fb["lc"].str.contains("tech")]
    prairies_negative_auto = fb[
        (fb["Region"] == "Prairies") & (fb["Channel"] == "Auto Service") & (fb["Rating (1-5)"] <= 2)
    ]

    metrics = {
        "company_name": COMPANY_NAME,
        "pattern_a": {
            "top_category_last_4_weeks": top_last4,
            "winter_tires_ramp_ratio_by_region": region_ramp,
        },
        "pattern_b": {
            "at_risk_winter_tire_rows": stock_risk.to_dict(orient="records"),
            "ontario_all_season_overstock_rows": all_season_over.to_dict(orient="records"),
        },
        "pattern_c": {
            "prairies_first_week_over_95": str(prairies_over95["Week Ending"].min()),
            "prairies_weekly_utilization": [
                {"Week Ending": str(r["Week Ending"]), "Bay Utilization %": round(float(r["Bay Utilization %"]), 4)}
                for _, r in prairies.iterrows()
            ],
            "over_capacity_rows": over_capacity_store.head(10).to_dict(orient="records"),
            "tire_changeover_share_oct_nov": round(float(tire_share), 4),
        },
        "pattern_d": {
            "seasonal_variance_pct": [
                {"Region": r["Region"], "Variance %": round(float(r["Variance %"]), 4)}
                for _, r in seasonal_gap.iterrows()
            ],
            "automotive_national_variance_pct": round(float(automotive_national_var), 4),
            "fixing_national_variance_pct": round(float(fixing_national_var), 4),
        },
        "pattern_e": {
            "flyer_units_lift_pct": round(float(units_lift), 4),
            "flyer_margin_delta_points": round(float(gm_delta_pts), 2),
            "power_tools_negative_margin_rows": margin_outliers.to_dict(orient="records"),
        },
        "pattern_f": {
            "wait_mentions": int(len(wait_mentions)),
            "booking_mentions": int(len(booking_mentions)),
            "tech_praise_mentions": int(len(tech_praise)),
            "prairies_negative_auto_comments": int(len(prairies_negative_auto)),
        },
    }
    return metrics


def write_answer_key(metrics: Dict[str, object]) -> None:
    b = metrics["pattern_b"]
    c = metrics["pattern_c"]
    d = metrics["pattern_d"]
    e = metrics["pattern_e"]
    f = metrics["pattern_f"]
    a = metrics["pattern_a"]

    lines = [
        "# Answer key — Fall-to-winter readiness patterns",
        "",
        "**Classification: Public**",
        "",
        f"Company context from `COMPANY_NAME`: **{COMPANY_NAME}**",
        "",
        "All values below are calculated by `scripts/generate_workbook.py` from generated data.",
        "",
        "## Pattern A — Winter Tires trend",
        f"- Last 4 weeks top category by Net Sales: **{a['top_category_last_4_weeks']}**",
        "- Winter Tires ramp ratio (late 4-week avg vs early 5-week avg) by region:",
    ]
    for region, ratio in a["winter_tires_ramp_ratio_by_region"].items():
        lines.append(f"  - {region}: **{ratio}x**")

    lines += [
        "",
        "## Pattern B — Inventory stock risk / overstock",
        f"- Winter Tire risk rows (WOS < 1.0 and delivery after 2026-11-20): **{len(b['at_risk_winter_tire_rows'])}**",
    ]
    for row in b["at_risk_winter_tire_rows"]:
        lines.append(
            f"  - Store {row['Store ID']} ({row['Region']}), {row['Product']}, WOS {row['Weeks of Supply']}, next delivery {row['Next Delivery Date']}"
        )
    lines.append(f"- Ontario All-Season overstock rows (WOS > 10): **{len(b['ontario_all_season_overstock_rows'])}**")
    for row in b["ontario_all_season_overstock_rows"]:
        lines.append(f"  - Store {row['Store ID']} {row['Product']}, WOS {row['Weeks of Supply']}")

    lines += [
        "",
        "## Pattern C — Auto Service capacity",
        f"- Prairies first week average utilization above 95%: **{c['prairies_first_week_over_95']}**",
        f"- Tire Changeover share of all jobs (Oct-Nov): **{round(c['tire_changeover_share_oct_nov'] * 100, 2)}%**",
        f"- Prairies rows above 100% utilization: **{len(c['over_capacity_rows'])} shown (top 10 list in metrics file)**",
    ]

    lines += [
        "",
        "## Pattern D — Target gaps",
        "- Seasonal variance vs target (Actual/Target - 1):",
    ]
    for row in d["seasonal_variance_pct"]:
        lines.append(f"  - {row['Region']}: **{round(row['Variance %'] * 100, 2)}%**")
    lines.append(f"- Automotive national variance: **{round(d['automotive_national_variance_pct'] * 100, 2)}%**")
    lines.append(f"- Fixing national variance: **{round(d['fixing_national_variance_pct'] * 100, 2)}%**")

    lines += [
        "",
        "## Pattern E — Flyer effect and outlier",
        f"- Flyer unit lift vs non-flyer: **{round(e['flyer_units_lift_pct'] * 100, 2)}%**",
        f"- Flyer Gross Margin % delta: **{e['flyer_margin_delta_points']} points**",
        f"- Power Tools negative margin rows on 2026-10-03: **{len(e['power_tools_negative_margin_rows'])}**",
    ]
    for row in e["power_tools_negative_margin_rows"]:
        lines.append(
            f"  - Store {row['Store ID']}: GM% {round(row['Gross Margin %'] * 100, 2)}%, Net Sales ${row['Net Sales (CAD)']:.2f}, Gross Margin ${row['Gross Margin (CAD)']:.2f}"
        )

    lines += [
        "",
        "## Pattern F — Feedback themes",
        f"- Comments mentioning waits: **{f['wait_mentions']}**",
        f"- Comments mentioning booking confusion: **{f['booking_mentions']}**",
        f"- Comments praising technicians/friendly service: **{f['tech_praise_mentions']}**",
        f"- Prairies Auto Service comments rated 1-2: **{f['prairies_negative_auto_comments']}**",
        "",
        "## Verification source",
        "- Machine-readable metrics: `sample-data/answer-key-metrics.json`",
    ]
    ANSWER_KEY_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    rng = random.Random(RANDOM_SEED)
    SAMPLE_DATA.mkdir(parents=True, exist_ok=True)
    CSV_DIR.mkdir(parents=True, exist_ok=True)

    stores = build_stores(rng)
    sales = build_sales(stores, rng)
    service = build_service(stores, rng)
    inventory = build_inventory(stores, rng)
    targets = build_targets(sales)
    feedback = build_feedback(stores, rng)

    dfs = {
        "Stores": ("tblStores", stores),
        "Sales": ("tblSales", sales),
        "AutoService": ("tblService", service),
        "Inventory": ("tblInventory", inventory),
        "Targets": ("tblTargets", targets),
        "Feedback": ("tblFeedback", feedback),
    }
    write_workbook(dfs)
    export_csvs(dfs)

    metrics = compute_metrics(stores, sales, service, inventory, targets, feedback)
    METRICS_JSON_PATH.write_text(json.dumps(metrics, indent=2, default=str), encoding="utf-8")
    write_answer_key(metrics)

    print(f"Generated workbook: {WORKBOOK_PATH}")
    print(f"Generated answer key: {ANSWER_KEY_PATH}")
    print(f"Generated metrics: {METRICS_JSON_PATH}")


if __name__ == "__main__":
    main()
