# Backup Storage Cost Calculator

A Python script that generates an Excel workbook to calculate Amazon S3 storage costs for backup retention policies across multiple storage tiers over a 1-year period.

## Features

- **Configurable Backup Types**: Hourly, daily, weekly, and off-account backups with selectable storage tiers
- **Individual Retention Periods**: Separate retention settings for hourly and daily backups
- **Storage Tier Selection**: Each backup type can be assigned to any S3 storage tier (1-5)
- **Minimum Duration Compliance**: Accounts for minimum storage duration requirements and penalties
- **Cost Calculations**: Monthly storage costs and early deletion penalties
- **Data Flow Visualization**: Visual representation of backup storage assignments

## Installation

```bash
pip install pandas openpyxl
```

## Usage

```bash
python backup_cost_calculator.py
```

## Configuration Parameters

### Backup Configuration
| Parameter | Default | Description |
|-----------|---------|-------------|
| Hourly backups count | 1 | Number of hourly backups to keep |
| Hourly backup tier | 1 (S3 Standard) | Storage tier for hourly backups |
| Hourly retention | 7 days | Individual retention period for hourly backups |
| Daily backups count | 1 | Number of daily backups to keep |
| Daily backup tier | 1 (S3 Standard) | Storage tier for daily backups |
| Daily retention | 30 days | Individual retention period for daily backups |
| Weekly backups count | 1 | Number of weekly backups to keep |
| Weekly backup tier | 2 (S3 Intelligent) | Storage tier for weekly backups |
| Off-account backups count | 1 | Number of off-account backups to keep |
| Off-account backup tier | 4 (Glacier IR) | Storage tier for off-account backups |
| Incremental storage size (GB) | 100 | Size of each backup increment |

### Storage Tier Retention
| Parameter | Default | Description |
|-----------|---------|-------------|
| S3 Standard retention (days) | 15 | Days to retain in S3 Standard |
| S3 Intelligent retention (days) | 15 | Days to retain in S3 Intelligent |
| S3-IA retention (days) | 15 | Days to retain in S3-IA (min 30 required) |
| Glacier IR retention (days) | 15 | Days to retain in Glacier IR (min 90 required) |
| Glacier Deep Archive retention (days) | 15 | Days to retain in Glacier Deep Archive (min 180 required) |

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
  - **Configuration**: Editable backup parameters, tier assignments, and data flow visualization
  - **Pricing**: Storage tier costs, minimum durations, and retrieval costs
  - **Cost Calculation**: Monthly storage breakdown, costs, and penalties

## Backup Distribution Logic

1. **Hourly backups**: Use individual retention period (E4) and configurable tier (C4)
2. **Daily backups**: Use individual retention period (E5) and configurable tier (C5)
3. **Weekly backups**: Use tier-based retention (B14-B18) and configurable tier (C6)
4. **Off-account backups**: Use tier-based retention (B14-B18) and configurable tier (C7)
5. **Tier Codes**: 1=S3 Standard, 2=S3 Intelligent, 3=S3-IA, 4=Glacier IR, 5=Glacier DA
6. **Storage Calculation**: Each backup type contributes to its assigned tier's monthly storage
7. **Cost Calculation**: Includes both storage costs and early deletion penalties for minimum duration violations

## Excel Sheet Structure

### Configuration Sheet
- **Backup Configuration**: Blue cells for backup counts, tier assignments, and individual retention periods
- **Storage Tier Retention**: Blue cells for tier-based retention periods
- **Data Flow Visualization**: Shows current backup-to-tier assignments and retention periods
- **Tier Codes Reference**: Quick reference for tier numbers

### Pricing Sheet
- **Blue cells**: Editable pricing per GB/month for each storage tier
- **Min Duration**: AWS minimum storage duration requirements
- **Retrieval Costs**: Per-GB retrieval costs for each tier

### Cost Calculation Sheet
- **Monthly Breakdown**: Storage distribution across all tiers by month
- **Storage Costs**: Monthly costs based on tier pricing and retention periods
- **Early Deletion Penalties**: Penalties for retention periods below minimum requirements
- **Annual Totals**: Sum of all monthly costs and storage amounts

## Customization

Edit the blue cells in the Excel output to adjust:
- Backup counts and tier assignments
- Individual retention periods for hourly/daily backups
- Tier-based retention periods for weekly/off-account backups
- Storage pricing (update with current AWS rates)

## Key Features

- **Flexible Tier Assignment**: Any backup type can use any storage tier
- **Individual vs Tier-Based Retention**: Hourly/daily use individual periods, weekly/off-account use tier-based periods
- **Penalty Calculation**: Automatic early deletion penalty calculation for sub-minimum retention periods
- **Visual Data Flow**: See exactly which backups go to which tiers with what retention
- **Real-time Updates**: Excel formulas update automatically when parameters change

## Notes

- Pricing reflects current AWS S3 rates (US East N. Virginia)
- Minimum storage duration charges are included in calculations
- Retrieval costs are tracked but not automatically calculated
- Update pricing regularly to reflect current AWS rates