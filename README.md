# S3 Storage Cost Calculator

A Python script that generates an Excel worksheet to calculate Amazon S3 storage costs over a 1-year period.

## Features

- **Input Parameters**: Base size, monthly incremental growth, storage period, transition timings
- **Storage Tiers**: Standard, Standard-IA, Glacier Flexible Retrieval
- **Automatic Transitions**: Data moves between tiers based on age constraints
- **Cost Calculations**: Monthly and annual cost projections
- **Current Pricing**: Includes latest AWS S3 pricing (as of 2024)

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python s3_cost_calculator.py
```

This generates an Excel file named `S3_Cost_Calculator_YYYYMMDD.xlsx` with:

### Input Section
- Base Size (GB): Initial storage amount
- Monthly Incremental Size (GB): Growth per month
- Storage Period (months): Calculation duration
- Transition to IA after (days): When data moves to Standard-IA
- Transition to Glacier after (days): When data moves to Glacier

### Pricing Section
- Current AWS S3 pricing per GB/month for each tier
- Easily updatable for price changes

### Calculation Table
- Monthly breakdown of storage distribution across tiers
- Cost calculations for each tier
- Total monthly and annual costs

## Customization

Modify the blue cells in the Excel file to adjust:
- Storage parameters
- Transition policies
- Pricing (update with current AWS rates)

## Notes

- Pricing reflects current AWS S3 rates (update as needed)
- Transitions are based on data age in days
- All calculations assume US East (N. Virginia) region pricing