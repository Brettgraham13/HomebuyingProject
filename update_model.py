from __future__ import annotations

import math
from pathlib import Path
from typing import Dict, Any

import pandas as pd
import yaml
import folium
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = ROOT / "data" / "listings_input.csv"
CONFIG_PATH = ROOT / "config.yaml"
OUTPUT_DIR = ROOT / "output"
XLSX_PATH = OUTPUT_DIR / "philly_home_model.xlsx"
MAP_PATH = OUTPUT_DIR / "philly_deal_map.html"


def money(x: float) -> float:
    if pd.isna(x):
        return 0.0
    return float(x)


def load_config() -> Dict[str, Any]:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def score_row(row: pd.Series, cfg: Dict[str, Any]) -> Dict[str, Any]:
    m = cfg["model"]
    t = cfg["thresholds"]
    w = cfg["weights"]

    price = money(row.get("price"))
    hoa = money(row.get("hoa"))
    rent = money(row.get("est_rent"))
    insurance = money(row.get("insurance_monthly")) or m["insurance_monthly_default"]
    beds = money(row.get("beds"))

    interest = price * m["mortgage_rate"] / 12
    maintenance = price * m["maintenance_annual_pct"] / 12
    true_cost = interest + hoa + insurance + maintenance + m["property_tax_monthly"]
    delta_vs_rent = rent - true_cost

    management = rent * m["management_fee_pct"]
    vacancy = rent * m["vacancy_pct"]
    rental_phase_cost = true_cost + management + vacancy
    net_rental_cashflow = rent - rental_phase_cost

    # Component scores, clipped to 0-100
    rent_spread_score = max(0, min(100, 50 + (delta_vs_rent / 10)))
    rental_upside_score = 80 if beds >= 2 else 55
    hoa_score = max(0, min(100, 100 - (hoa / max(t["ideal_max_hoa"], 1) * 50)))
    location_score = 85  # starter value; adjust manually after viewing commute/block quality
    liquidity_score = 75 if price <= t["ideal_max_price"] else 60

    risk_penalty = 0
    flags = []
    if str(row.get("va_status", "")).lower() not in {"approved", "yes"}:
        flags.append("Verify VA condo approval")
        risk_penalty += 10
    if str(row.get("rental_status", "")).lower() not in {"allowed", "yes"}:
        flags.append("Verify rental rules")
        risk_penalty += 10
    if str(row.get("hoa_verified", "")).lower() not in {"yes", "verified"}:
        flags.append("HOA unverified")
        risk_penalty += 5
    if hoa > t["caution_hoa"]:
        flags.append("High HOA")
        risk_penalty += 15
    if price > t["hard_max_price"]:
        flags.append("Above VA preapproval")
        risk_penalty += 30

    risk_score = max(0, 100 - risk_penalty)

    total_score = (
        rent_spread_score * w["rent_spread"]
        + rental_upside_score * w["rental_upside"]
        + hoa_score * w["hoa_reasonableness"]
        + location_score * w["location"]
        + liquidity_score * w["liquidity"]
        + risk_score * w["risk_flags"]
    ) / sum(w.values())

    if total_score >= 75 and delta_vs_rent >= 0:
        rec = "Strong fit"
    elif total_score >= 65:
        rec = "Worth diligence"
    elif total_score >= 55:
        rec = "Neutral / price-sensitive"
    else:
        rec = "Avoid unless terms improve"

    return {
        "interest_monthly": interest,
        "maintenance_monthly": maintenance,
        "true_monthly_cost": true_cost,
        "delta_vs_rent": delta_vs_rent,
        "management_fee": management,
        "vacancy_reserve": vacancy,
        "net_rental_cashflow": net_rental_cashflow,
        "score": total_score,
        "recommendation": rec,
        "risk_flags": "; ".join(flags) if flags else "None",
    }


def build_outputs() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    cfg = load_config()
    df = pd.read_csv(DATA_PATH)

    derived = df.apply(lambda r: pd.Series(score_row(r, cfg)), axis=1)
    out = pd.concat([df, derived], axis=1)
    out = out.sort_values(["score", "delta_vs_rent"], ascending=[False, False]).reset_index(drop=True)

    build_workbook(out, cfg)
    build_map(out)
    print(f"Wrote {XLSX_PATH}")
    print(f"Wrote {MAP_PATH}")


def build_workbook(df: pd.DataFrame, cfg: Dict[str, Any]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ranked Listings"
    ws.sheet_view.showGridLines = False

    dark = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    input_fill = PatternFill("solid", fgColor="FFF2CC")
    flag_fill = PatternFill("solid", fgColor="FCE4D6")
    green_fill = PatternFill("solid", fgColor="E2F0D9")
    thin_gray = Side(style="thin", color="D9E1F2")

    columns = [
        "recommendation", "score", "address", "neighborhood", "url", "price", "beds", "baths", "sqft", "hoa",
        "est_rent", "interest_monthly", "insurance_monthly", "maintenance_monthly", "true_monthly_cost",
        "delta_vs_rent", "management_fee", "vacancy_reserve", "net_rental_cashflow", "risk_flags",
        "va_status", "rental_status", "hoa_verified", "lat", "lon", "notes"
    ]
    headers = [c.replace("_", " ").title() for c in columns]
    ws.append(headers)

    for cell in ws[1]:
        cell.fill = dark
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for _, row in df.iterrows():
        ws.append([row.get(c, "") for c in columns])

    # formatting
    currency_cols = {"price", "hoa", "est_rent", "interest_monthly", "insurance_monthly", "maintenance_monthly", "true_monthly_cost", "delta_vs_rent", "management_fee", "vacancy_reserve", "net_rental_cashflow"}
    percent_cols = set()
    for idx, c in enumerate(columns, start=1):
        col_letter = get_column_letter(idx)
        ws.column_dimensions[col_letter].width = 18
        for cell in ws[col_letter][1:]:
            cell.border = Border(bottom=thin_gray)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c in currency_cols:
                cell.number_format = '$#,##0;[Red]($#,##0);-'
            elif c == "score":
                cell.number_format = '0.0'
            if c in {"price", "hoa", "est_rent", "insurance_monthly", "va_status", "rental_status", "hoa_verified", "notes"}:
                cell.font = Font(color="0000FF")
            elif c in {"interest_monthly", "maintenance_monthly", "true_monthly_cost", "delta_vs_rent", "management_fee", "vacancy_reserve", "net_rental_cashflow", "score", "recommendation", "risk_flags"}:
                cell.font = Font(color="000000")
            if c == "url" and cell.value:
                cell.hyperlink = cell.value
                cell.style = "Hyperlink"

    # highlight recommendation and risk fields
    rec_col = columns.index("recommendation") + 1
    risk_col = columns.index("risk_flags") + 1
    hoa_verified_col = columns.index("hoa_verified") + 1
    for r in range(2, ws.max_row + 1):
        rec = str(ws.cell(r, rec_col).value)
        if rec == "Strong fit":
            ws.cell(r, rec_col).fill = green_fill
        elif "Avoid" in rec:
            ws.cell(r, rec_col).fill = flag_fill
        else:
            ws.cell(r, rec_col).fill = input_fill
        ws.cell(r, risk_col).fill = flag_fill
        if str(ws.cell(r, hoa_verified_col).value).lower() not in {"yes", "verified"}:
            ws.cell(r, hoa_verified_col).fill = input_fill

    # comments / sources
    ws["F1"].comment = Comment("Price is imported from listing_input.csv. Verify against listing source before making an offer.", "OpenAI")
    ws["J1"].comment = Comment("HOA is the highest-leverage input. Verify from listing docs/condo resale package.", "OpenAI")
    ws["K1"].comment = Comment("Estimated rent is an assumption from local comparable rents; replace with realtor/property manager comps.", "OpenAI")

    # assumptions sheet
    a = wb.create_sheet("Assumptions")
    a.sheet_view.showGridLines = False
    a.append(["Assumption", "Value", "Notes"])
    for cell in a[1]:
        cell.fill = dark
        cell.font = Font(color="FFFFFF", bold=True)
    for key, value in cfg["model"].items():
        a.append([key, value, "Editable model assumption"])
    for r in range(2, a.max_row + 1):
        a.cell(r, 2).font = Font(color="0000FF")
    a.column_dimensions["A"].width = 28
    a.column_dimensions["B"].width = 18
    a.column_dimensions["C"].width = 45

    # deal breakers
    d = wb.create_sheet("Diligence Checklist")
    d.sheet_view.showGridLines = False
    checklist = [
        ["Question", "Why it matters", "Status"],
        ["Is the condo project VA-approved?", "Deal-breaker for VA financing unless approval is obtained.", "To verify"],
        ["Are rentals allowed after your occupancy period?", "Your year-3 rental plan fails if rentals are banned or capped.", "To verify"],
        ["Is there a rental cap or waiting list?", "A cap may prevent renting even if rentals are generally allowed.", "To verify"],
        ["Any pending special assessments?", "Can destroy economics even when HOA seems manageable.", "To verify"],
        ["Are reserves adequately funded?", "Weak reserves can lead to future assessments.", "To verify"],
        ["Any litigation or insurance issues?", "Can block lending or signal building risk.", "To verify"],
        ["What exactly does HOA include?", "Utilities included can partly offset a high HOA.", "To verify"],
    ]
    for row in checklist:
        d.append(row)
    for cell in d[1]:
        cell.fill = dark
        cell.font = Font(color="FFFFFF", bold=True)
    d.column_dimensions["A"].width = 38
    d.column_dimensions["B"].width = 65
    d.column_dimensions["C"].width = 18
    for r in range(2, d.max_row + 1):
        d.cell(r, 3).fill = input_fill
        d.cell(r, 3).font = Font(color="0000FF")

    # map data
    m = wb.create_sheet("Map Data")
    m.sheet_view.showGridLines = False
    m.append(["Address", "Latitude", "Longitude", "Recommendation", "Score", "Map Link"])
    for _, row in df.iterrows():
        m.append([row["address"], row["lat"], row["lon"], row["recommendation"], row["score"], "Open Map"])
    for cell in m[1]:
        cell.fill = dark
        cell.font = Font(color="FFFFFF", bold=True)
    for r in range(2, m.max_row + 1):
        lat = m.cell(r, 2).value
        lon = m.cell(r, 3).value
        c = m.cell(r, 6)
        c.hyperlink = f"https://www.google.com/maps/search/?api=1&query={lat},{lon}"
        c.style = "Hyperlink"
    m.column_dimensions["A"].width = 42
    m.column_dimensions["F"].width = 18

    wb.save(XLSX_PATH)


def build_map(df: pd.DataFrame) -> None:
    valid = df.dropna(subset=["lat", "lon"])
    if valid.empty:
        return
    center = [valid["lat"].mean(), valid["lon"].mean()]
    fmap = folium.Map(location=center, zoom_start=14, tiles="OpenStreetMap")

    def color_for(rec: str) -> str:
        if rec == "Strong fit":
            return "green"
        if "Avoid" in rec:
            return "red"
        return "orange"

    for _, row in valid.iterrows():
        popup_html = f"""
        <b>{row['address']}</b><br>
        Recommendation: {row['recommendation']}<br>
        Score: {row['score']:.1f}<br>
        Price: ${row['price']:,.0f}<br>
        HOA: ${row['hoa']:,.0f}<br>
        Est. rent: ${row['est_rent']:,.0f}<br>
        True cost: ${row['true_monthly_cost']:,.0f}<br>
        Delta vs rent: ${row['delta_vs_rent']:,.0f}<br>
        Flags: {row['risk_flags']}<br>
        <a href="{row['url']}" target="_blank">Listing</a>
        """
        folium.CircleMarker(
            location=[row["lat"], row["lon"]],
            radius=8,
            color=color_for(row["recommendation"]),
            fill=True,
            fill_opacity=0.8,
            popup=folium.Popup(popup_html, max_width=350),
            tooltip=f"{row['recommendation']}: {row['address']}",
        ).add_to(fmap)

    fmap.save(MAP_PATH)


if __name__ == "__main__":
    build_outputs()
