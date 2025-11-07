# Backup Storage Cost Calculator

A Python script that generates an Excel workbook to calculate Amazon S3 storage costs for backup retention policies across multiple storage tiers over a 1-year period.

## Features

- **Step-by-Step Configuration**: Organized 3-step setup process with clear visual flow
- **Dropdown Tier Selection**: Auto-populates backup counts based on predefined tier configurations
- **Individual Retention Control**: Separate retention settings for all backup types (hourly, daily, weekly, off-account)
- **Storage Tier Flexibility**: Each backup type can be assigned to any S3 storage tier (1-5)
- **Early Deletion Penalties**: Automatic calculation of penalties for retention periods below AWS minimums
- **AWS Backup Comparison**: Side-by-side cost comparison with AWS Backup service
- **Real-time Updates**: Excel formulas update automatically when parameters change

## Installation

```bash
pip install openpyxl
```

**Note**: The script requires a `backup_config.json` configuration file to run.

## Usage

1. Ensure `backup_config.json` exists with your configuration (see Configuration section)
2. Run the script:

```bash
python backup_cost_calculator.py
```

3. Open the generated Excel file and use the dropdown in Step 1 to select different tier configurations

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
  - **Configuration**: Step-by-step setup with dropdown tier selection and backup tier reference table
  - **Pricing**: Storage tier costs, minimum durations, retrieval costs, and AWS Backup pricing
  - **Cost Calculation**: Monthly storage breakdown, costs, penalties, AWS Backup comparison, and cost differences

## Error Handling

The script will raise a `FileNotFoundError` if the `backup_config.json` file is not found. Ensure the configuration file exists in the same directory as the script.

## Configuration Steps

### **Step 1: Backup Tier Selection**
- **Dropdown menu**: Select from Tier 1, Tier 2, Tier 3, or Tier 4 configurations
- **Auto-population**: Backup counts automatically update based on selected tier
- **Reference table**: Shows predefined backup counts for each tier

### **Step 2: Backup Configuration** 
- **Individual control**: Set backup counts, storage tiers, and retention periods
- **Blue highlighting**: All editable parameters clearly marked
- **Real-time updates**: Tier names update automatically based on storage tier selection

### **Step 3: Storage Tier Retention**
- **Tier-based retention**: Fallback values when individual retention not specified
- **Minimum compliance**: Shows AWS minimum duration requirements and penalty status
- **Usage note**: Explains when these values are used in calculations

## Excel Sheet Structure

### Configuration Sheet
- **Step 1**: Dropdown tier selection with auto-population
- **Step 2**: Individual backup configuration (counts, tiers, retention)
- **Step 3**: Tier-based retention settings with compliance status
- **Reference Table**: Predefined tier configurations for easy copy-paste

### Pricing Sheet
- **Storage Tiers**: S3 Standard, S3 Intelligent, S3-IA, Glacier IR, Glacier Deep Archive
- **AWS Backup**: Separate pricing row for AWS Backup service comparison
- **Editable Pricing**: Blue-highlighted cells for easy price updates
- **Minimum Durations**: AWS requirements for each storage tier

### Cost Calculation Sheet
- **Monthly Analysis**: 12-month breakdown of storage and costs
- **Storage Distribution**: GB amounts across all storage tiers
- **Cost Components**: Storage costs, early deletion penalties, total monthly costs
- **AWS Comparison**: AWS Backup costs and cost difference calculations
- **Annual Summary**: Total costs and storage amounts for the year

## Customization

### JSON Configuration
Edit the `backup_config.json` file to adjust:
- Default backup configurations and tier assignments
- Individual retention periods for all backup types
- Tier-based retention periods for fallback scenarios
- Storage pricing (update with current AWS rates)
- AWS Backup pricing for comparison
- Minimum duration requirements and retrieval costs

### Excel Interface
**Step 1**: Use dropdown to quickly switch between predefined tier configurations
**Step 2**: Edit blue cells to customize:
- Backup counts for each type
- Storage tier assignments (1-5)
- Individual retention periods
- Storage size per backup
**Step 3**: Adjust tier-based retention periods and view compliance status
**Pricing Sheet**: Update blue cells with current AWS pricing

## Key Features

- **User-Friendly Interface**: Step-by-step configuration with dropdown selections
- **Flexible Configuration**: Any backup type can use any storage tier with individual retention
- **Comprehensive Cost Analysis**: Storage costs, penalties, and AWS Backup comparison
- **Early Deletion Penalties**: Automatic calculation based on individual retention vs AWS minimums
- **Real-time Updates**: All calculations update automatically when parameters change
- **Professional Formatting**: Clean layout with blue highlighting for editable parameters

## Notes

- **Pricing**: Reflects current AWS S3 rates (US East N. Virginia region)
- **Early Deletion Penalties**: Calculated for S3-IA (30 days), Glacier IR (90 days), and Glacier Deep Archive (180 days)
- **AWS Backup Comparison**: Uses configurable AWS Backup pricing for cost comparison
- **Individual Retention**: All backup types (hourly, daily, weekly, off-account) use their individual retention periods for calculations
- **Tier-based Retention**: Used only as fallback when individual retention is not specified
- **Regular Updates**: Update pricing in both JSON config and Excel pricing sheet to reflect current AWS rates

## Verification

Run `python verify_cost_calculations.py` to manually verify calculation accuracy and see expected Excel values.