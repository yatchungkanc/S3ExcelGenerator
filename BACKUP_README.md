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

**Note**: The script requires a `backup_config.json` configuration file to run.

## Usage

1. Create a `backup_config.json` file with your configuration (see Configuration section)
2. Run the script:

```bash
python backup_cost_calculator.py
```

## Configuration

The script requires a `backup_config.json` file with the following structure:

```json
{
  "pricing": {
    "S3 Standard": 0.023,
    "S3 Intelligent": 0.0125,
    "S3-IA": 0.0125,
    "Glacier IR": 0.004,
    "Glacier Deep Archive": 0.00099
  },
  "min_duration": {
    "S3 Standard": 0,
    "S3 Intelligent": 0,
    "S3-IA": 30,
    "Glacier IR": 90,
    "Glacier Deep Archive": 180
  },
  "retrieval_costs": {
    "S3 Standard": 0.00,
    "S3 Intelligent": 0.00,
    "S3-IA": 0.01,
    "Glacier IR": 0.03,
    "Glacier Deep Archive": 0.02
  },
  "backup_config": {
    "hourly": {"count": 1, "tier": 1, "retention": 7},
    "daily": {"count": 1, "tier": 1, "retention": 30},
    "weekly": {"count": 1, "tier": 2},
    "off_account": {"count": 1, "tier": 4}
  },
  "storage_size_gb": 100,
  "tier_retention": {
    "S3 Standard": 15,
    "S3 Intelligent": 15,
    "S3-IA": 15,
    "Glacier IR": 15,
    "Glacier Deep Archive": 15
  }
}
```

### Configuration Parameters

- **pricing**: Cost per GB/month for each storage tier (USD)
- **min_duration**: AWS minimum storage duration requirements (days)
- **retrieval_costs**: Cost per GB for data retrieval (USD)
- **backup_config**: Backup type settings (count, tier, retention)
- **storage_size_gb**: Size of each backup increment
- **tier_retention**: Days to retain data in each storage tier



## Output Files

- `Backup_Cost_Calculator_YYYYMMDD.xlsx` with three sheets:
  - **Configuration**: Editable backup parameters, tier assignments, and data flow visualization
  - **Pricing**: Storage tier costs, minimum durations, and retrieval costs
  - **Cost Calculation**: Monthly storage breakdown, costs, and penalties

## Error Handling

The script will raise a `FileNotFoundError` if the `backup_config.json` file is not found. Ensure the configuration file exists in the same directory as the script.

## Backup Distribution Logic

1. **Hourly backups**: Use individual retention period and configurable tier (1-5)
2. **Daily backups**: Use individual retention period and configurable tier (1-5)
3. **Weekly backups**: Use tier-based retention from configuration and configurable tier (1-5)
4. **Off-account backups**: Use tier-based retention from configuration and configurable tier (1-5)
5. **Tier Codes**: 1=S3 Standard, 2=S3 Intelligent, 3=S3-IA, 4=Glacier IR, 5=Glacier DA
6. **Storage Calculation**: Each backup type contributes to its assigned tier's monthly storage based on retention periods
7. **Cost Calculation**: Includes storage costs, early deletion penalties, and minimum duration compliance checks

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

### JSON Configuration
Edit the `backup_config.json` file to adjust:
- Backup counts and tier assignments
- Individual retention periods for hourly/daily backups
- Tier-based retention periods for all storage tiers
- Storage pricing (update with current AWS rates)
- Minimum duration requirements
- Retrieval costs

### Excel Output
Edit the blue cells in the generated Excel file to adjust:
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