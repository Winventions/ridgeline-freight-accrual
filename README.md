# Ridgeline Foods Freight Accrual Engine

This project builds an activity-based April 2026 freight accrual for Ridgeline Foods using shipment activity, contracted carrier rate cards, historical invoice detail, and Denise's trailing-average baseline.

## Setup

Install Python 3.10+ with the required libraries:

```bash
pip install -r requirements.txt
```

Run the engine from the project root:

```bash
python run_accrual.py --data-dir data --output-dir output
```

The source files expected in `data/` are:

- `shipments_apr2026.csv`
- `rate_card_peak_logistics.csv`
- `rate_card_heartland_freight.csv`
- `rate_card_coastal_express.csv`
- `freight_invoices_oct2025_mar2026_v2.csv`
- `denise_accruals_v2.csv`

## Outputs

The run creates:

- `output/ridgeline_freight_accrual_audit_pack.xlsx`
- `output/ridgeline_freight_journal_entry.csv`
- `output/exception_log.csv`

## Web Submission App

This repo also includes a Vercel-ready web app:

- `index.html`
- `styles.css`
- `assets/accrual_breakdown.png`
- `api/regenerate.py`
- `vercel.json`

The page links directly to the generated April 2026 audit pack, journal entry CSV, exception log, README, requirements, and source script. It also includes an upload runner that accepts a fresh set of the six required CSVs and regenerates a new close package.

The upload runner posts the files to `api/regenerate.py`, which:

1. Writes the uploaded CSVs to a temporary serverless directory.
2. Executes the same Python accrual engine used locally.
3. Creates a fresh audit workbook, journal CSV, exception log, and `run_summary.json`.
4. Returns the generated files as `ridgeline_freight_accrual_outputs.zip`.

To preview locally, open `index.html` in a browser.

To deploy on Vercel:

1. Push this project to GitHub.
2. In Vercel, choose **Add New Project** and import the GitHub repo.
3. Leave the framework preset as **Other** or static.
4. Use the repository root as the project root.
5. Deploy.

No build command is required. Vercel will install dependencies from `requirements.txt` for the Python serverless function.

The workbook includes these tabs:

1. Executive Summary
2. Journal Entry
3. Carrier Accrual Detail
4. Shipment-Level Estimate
5. Exception Log
6. Baseline Comparison
7. Assumptions & Controls
8. Source File Manifest

## Accounting Approach

The accrual is based on April shipment activity rather than a trailing average. Each shipment is evaluated using the relevant carrier's contracted rate logic:

- Peak Logistics: per-mile rate by weight tier, 14% fuel surcharge, minimum charge, and accessorials.
- Heartland Freight: ZIP-prefix zone flat rate, quarterly cumulative volume discount reset for Q2 on April 1, and accessorials.
- Coastal Express: destination ZIP region per-pound rate, minimum charge, 9.5% fuel surcharge, residential delivery surcharge by weight tier, and other accessorials.

April actual invoices are not included in the provided data, so April is clearly labeled as an estimate. Denise's workbook is used as a baseline by calculating a January-March 2026 trailing three-month average by carrier and comparing that to the activity-based April estimate.

## Evidence Classes

Each shipment has an evidence class:

- `Direct Rate Card`: calculated directly from a valid rate-card mapping.
- `Historical Estimate`: calculated from historical invoice averages because direct rate-card mapping was incomplete.
- `Assumption Required`: calculated with a broader historical fallback, such as carrier-level averages.
- `Flagged Exception`: source data or mapping was insufficient for a reliable estimate.

The engine separates assumption-based dollars from direct rate-card dollars in the shipment detail and carrier summary.

## Key Controls

Controls built into the run:

- Required field validation for shipment ID, carrier, date, destination, and weight.
- Carrier and service-level normalization for messy 3PL names.
- Carrier-specific rate-card parsing from human-formatted CSVs.
- Heartland quarterly volume discount sequencing with April as the first month of Q2.
- Missing weight estimation from historical invoice averages.
- Historical fallback for rate-card mapping gaps, while still logging the mapping issue.
- Stale rate-card warnings for rate cards older than 12 months.
- Outlier shipment cost checks against carrier historical invoice patterns.
- Tie-out across shipment detail, carrier summary, and journal entry.

## Review Notes

Peak Logistics remains the highest-risk carrier because the Peak rate card does not include mileage for several April destinations that appear in historical invoices. Those shipments are estimated from historical invoice averages and flagged in the exception log for review with the carrier or account manager.

All three rate cards are effective January 2025 and are flagged as stale for April 2026 close unless management confirms that no newer cards exist.

## Rerunning Next Month

To rerun for a new close period:

1. Replace the shipment file in `data/` with the new month's activity.
2. Add or replace carrier rate cards if newer versions are available.
3. Add the latest invoice history once received.
4. Update Denise or baseline history if a new baseline workbook is provided.
5. Run `python run_accrual.py --data-dir data --output-dir output`.
6. Review `exception_log.csv` before posting the journal entry.

If a new carrier is added, add its aliases and rate logic to `run_accrual.py`, then confirm it appears in the carrier detail and tie-out controls.
