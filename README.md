# Backup Storage Cost Calculator

Generates an Excel workbook for comparing S3-based backup retention costs against AWS Backup over a 12-month period.

The workbook is driven by `backup_config.json` and includes editable Excel formulas for backup counts, storage tiers, retention periods, pricing, early deletion penalties, and AWS Backup comparison costs.

## Files

- `backup_cost_calculator.py`: Main workbook generator.
- `backup_config.json`: Default pricing, retention, backup tier, and AWS Backup comparison settings.
- `verify_cost_calculations.py`: Console helper for manually checking the default Month 1 calculations.
- `requirements.txt`: Python package requirements.

## Installation

```bash
python -m pip install -r requirements.txt
```

## Usage

```bash
python backup_cost_calculator.py
```

This creates `Backup_Cost_Calculator_YYYYMMDD.xlsx` with three sheets:

- `Configuration`: Backup tier selector, backup counts, storage tier assignments, retention settings, and tier reference values.
- `Pricing`: S3 storage pricing, minimum storage durations, retrieval costs, and AWS Backup comparison pricing.
- `Cost Calculation`: Monthly storage distribution, storage cost, early deletion penalty, total monthly cost, AWS Backup cost, and annual totals.

## Configuration

Edit `backup_config.json` to change the defaults used when a workbook is generated:

- `pricing`: Storage price per GB-month for each S3 tier.
- `min_duration`: Minimum storage duration in days for each tier.
- `retrieval_costs`: Retrieval cost per GB for each tier.
- `backup_config`: Default storage tier and retention settings for hourly, daily, weekly, and off-account backups.
- `storage_size_gb`: Size of each backup increment.
- `tier_retention`: Default retention periods used for tier-level settings.
- `aws_backup_pricing`: AWS Backup storage and retrieval comparison rates.

In the generated workbook, blue cells are intended to be edited directly. Use the `Configuration` sheet to switch between Tier 1 through Tier 4 backup count profiles, adjust backup storage tiers, and tune retention. Use the `Pricing` sheet to update rates.

## Storage Tiers

Storage tier codes used in the `Configuration` sheet:

- `1`: S3 Standard
- `2`: S3 Intelligent-Tiering
- `3`: S3 Standard-IA
- `4`: S3 Glacier Instant Retrieval
- `5`: S3 Glacier Deep Archive

## Verification

Run the manual verification helper:

```bash
python verify_cost_calculations.py
```

It prints the expected Month 1 values for the default configuration so they can be compared with the generated workbook.

## Notes

- Pricing values are configurable estimates. Check current AWS pricing for the target region before using the workbook for decisions.
- Early deletion penalties are calculated for S3 Standard-IA, S3 Glacier Instant Retrieval, and S3 Glacier Deep Archive when retention is below the configured minimum duration.
- The generated workbook uses formulas, so values update when edited in Excel or another spreadsheet application with formula support.
