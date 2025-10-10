# Backup Storage Cost Calculator

A Python script that calculates Amazon S3 storage costs for backup retention policies across multiple storage tiers over a 1-year period.

## Features

- **Configurable Backup Types**: Hourly, daily, weekly, monthly, and off-account backups
- **Storage Tier Transitions**: Automatic transitions between S3 Standard → Glacier IR → Glacier Deep Archive
- **Minimum Duration Compliance**: Accounts for minimum storage duration requirements
- **Retrieval Cost Tracking**: Includes retrieval costs for each storage tier
- **Monthly Cost Projections**: 12-month cost breakdown and annual totals

## Installation

```bash
pip install pandas openpyxl
```

## Usage

```bash
python backup_cost_calculator.py
```

## Configuration Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| Hourly backups to keep | 24 | Number of hourly backups to keep |
| Daily backups to keep | 7 | Number of daily backups to keep |
| Weekly backups to keep | 4 | Number of weekly backups to keep |
| Monthly backups to keep | 12 | Number of monthly backups to keep |
| Off-account backups per year | 4 | Number of off-site backups retained yearly |
| Incremental storage size (GB) | 100 | Size of each backup increment |
| S3 Standard retention (days) | 30 | Days to retain in S3 Standard (0=skip) |
| S3 Intelligent retention (days) | 60 | Days to retain in S3 Intelligent (0=skip) |
| S3-IA retention (days) | 90 | Days to retain in S3-IA (0=skip, min 30) |
| Glacier IR retention (days) | 180 | Days to retain in Glacier IR (0=skip, min 90) |
| Glacier Deep Archive retention (days) | 365 | Days to retain in Glacier Deep Archive (0=skip, min 180) |

## Storage Tier Pricing (USD per GB/month)

| Tier | Price | Min Duration | Retrieval Cost |
|------|-------|--------------|----------------|
| S3 Standard | $0.023 | 0 days | $0.00 |
| S3 Intelligent | $0.0125 | 0 days | $0.00 |
| S3-IA | $0.0125 | 30 days | $0.01 |
| Glacier IR | $0.004 | 90 days | $0.03 |
| Glacier Deep Archive | $0.00099 | 180 days | $0.02 |

## Output Files

- `Backup_Cost_Calculator_YYYYMMDD.xlsx` with three sheets:
  - **Configuration**: Editable parameters
  - **Pricing**: Storage tier costs and requirements
  - **Cost Calculation**: Monthly breakdown and totals

## Backup Distribution Logic

1. **Hourly + Daily backups**: Always stored in S3 Standard (short-term retention)
2. **Weekly + Monthly + Off-account backups**: Flow through storage tiers based on retention settings:
   - If S3 Intelligent retention > 0: Stored in S3 Intelligent
   - If S3 Intelligent = 0 and S3-IA retention > 0: Stored in S3-IA
   - If both above = 0 and Glacier IR retention > 0: Stored in Glacier IR
   - If all above = 0 and Glacier DA retention > 0: Stored in Glacier Deep Archive
3. **Tier skipping**: Set retention to 0 to skip that tier entirely
4. **Monthly growth**: Monthly backups grow from 1 to 12 over the year
5. **Off-account timing**: Off-account backups only appear after month 12

## Customization

Modify the configuration parameters in the script or edit the blue cells in the Excel output to adjust:
- Backup frequency and retention
- Storage transition timings
- Incremental backup sizes
- Pricing (update with current AWS rates)

## Example Output

With default settings:
- **Annual storage cost**: ~$20,272
- **Average monthly storage**: ~82,283 GB
- **Peak storage**: Grows monthly as backups accumulate

## Notes

- Pricing reflects current AWS S3 rates (US East N. Virginia)
- Minimum storage duration charges are included in calculations
- Retrieval costs are tracked but not automatically calculated
- Update pricing regularly to reflect current AWS rates