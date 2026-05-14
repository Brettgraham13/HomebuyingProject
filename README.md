# Philly Home Listing Agent

A lightweight listing-ingestion and scoring agent for your Philadelphia VA-loan buy-vs-rent model.

This version is intentionally **semi-automated**: you paste listing rows from Redfin/Zillow/Realtor/MLS emails into `data/listings_input.csv`, then run one command to generate an updated Excel model and interactive map. This avoids brittle or non-compliant scraping while still giving you a repeatable decision engine.

## What it does

- Imports available listings from CSV
- Normalizes listing fields
- Estimates non-recoverable monthly cost:
  - mortgage interest only
  - HOA
  - insurance
  - maintenance reserve
  - vacancy/management in rental phase
- Excludes property tax by default
- Scores each listing against your criteria
- Flags VA/rental/HOA diligence risks
- Exports:
  - `output/philly_home_model.xlsx`
  - `output/philly_deal_map.html`
  - `Viable Listings` worksheet (non-"Avoid" opportunities)

## Quick start

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python update_model.py
```

Open:

- `output/philly_home_model.xlsx` (or repo root if `output/` is not present)
- `output/philly_deal_map.html` (or repo root if `output/` is not present)

## Update workflow

1. Add listings to `data/listings_input.csv` (or `listings_input.csv` in repo root).
2. Run `python update_model.py`.
3. Review the ranked sheet and map.
4. Verify any yellow/red due-diligence fields with your realtor.

## Important notes

This tool does **not** guarantee listing accuracy, VA condo eligibility, rental permissibility, or HOA reserve health. Treat it as a screening engine. Your realtor/lender should verify:

- VA condo approval
- condo rental restrictions
- HOA amount
- HOA reserve health
- pending assessments
- building litigation
- insurance requirements

## Optional future integrations

The cleanest fully independent version would connect to one of:

- MLS/RESO feed through your realtor or brokerage
- Bridge Interactive / MLS partner access
- a paid listings API where licensed for your intended use
- manual CSV exports from your agent's search portal
