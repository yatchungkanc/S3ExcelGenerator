#!/usr/bin/env python3
import json

def verify_cost_calculations():
    with open('backup_config.json', 'r') as f:
        config = json.load(f)
    
    # Configuration values
    backup_config = config['backup_config']
    storage_size_gb = config['storage_size_gb']
    tier_retention = config['tier_retention']
    pricing = config['pricing']
    aws_backup_pricing = config['aws_backup_pricing']
    
    print("=== Cost Calculation Verification ===")
    print(f"Storage size: {storage_size_gb} GB per backup")
    print(f"AWS Backup price: ${aws_backup_pricing['storage_cost_per_gb_month']}/GB/month")
    print()
    
    # Current configuration (Tier 1 defaults)
    hourly_count = 24  # From Tier 1 reference
    daily_count = 7    # From Tier 1 reference  
    weekly_count = 6   # From Tier 1 reference
    offaccount_count = 2  # From Tier 1 reference
    
    hourly_retention = backup_config['hourly']['retention']  # 7 days
    daily_retention = backup_config['daily']['retention']    # 30 days
    weekly_retention = tier_retention['S3 Intelligent']      # 15 days (E11)
    offaccount_retention = tier_retention['Glacier IR']      # 15 days (E12)
    
    s3_standard_price = pricing['S3 Standard']  # $0.023
    
    print("Configuration:")
    print(f"Hourly: {hourly_count} backups, {hourly_retention} days retention, Tier 1 (S3 Standard)")
    print(f"Daily: {daily_count} backups, {daily_retention} days retention, Tier 1 (S3 Standard)")
    print(f"Weekly: {weekly_count} backups, {weekly_retention} days retention, Tier 2 (S3 Intelligent)")
    print(f"Off-account: {offaccount_count} backups, {offaccount_retention} days retention, Tier 4 (Glacier IR)")
    print()
    
    # Month 1 calculations
    month = 1
    print(f"Month {month} Manual Calculations:")
    
    # Total backups
    total_backups = (hourly_count + daily_count + weekly_count + offaccount_count) * storage_size_gb
    print(f"Total Backups: ({hourly_count}+{daily_count}+{weekly_count}+{offaccount_count}) × {storage_size_gb} = {total_backups} GB")
    
    # S3 Standard storage (Hourly + Daily using tier 1)
    hourly_s3_std = hourly_count * storage_size_gb * min(hourly_retention, month * 30) / 30
    daily_s3_std = daily_count * storage_size_gb * min(daily_retention, month * 30) / 30
    total_s3_std = hourly_s3_std + daily_s3_std
    
    print(f"S3 Standard Storage:")
    print(f"  Hourly: {hourly_count} × {storage_size_gb} × min({hourly_retention}, {month*30})/30 = {hourly_s3_std:.2f} GB")
    print(f"  Daily: {daily_count} × {storage_size_gb} × min({daily_retention}, {month*30})/30 = {daily_s3_std:.2f} GB")
    print(f"  Total S3 Standard: {total_s3_std:.2f} GB")
    
    # S3 Intelligent storage (Weekly using tier 2)
    weekly_s3_int = weekly_count * storage_size_gb * min(weekly_retention, month * 30) / 30
    print(f"S3 Intelligent Storage:")
    print(f"  Weekly: {weekly_count} × {storage_size_gb} × min({weekly_retention}, {month*30})/30 = {weekly_s3_int:.2f} GB")
    
    # Glacier IR storage (Off-account using tier 4)
    offaccount_glacier = offaccount_count * storage_size_gb * min(offaccount_retention, month * 30) / 30
    print(f"Glacier IR Storage:")
    print(f"  Off-account: {offaccount_count} × {storage_size_gb} × min({offaccount_retention}, {month*30})/30 = {offaccount_glacier:.2f} GB")
    
    # Total storage
    total_storage = total_s3_std + weekly_s3_int + offaccount_glacier
    print(f"Total Storage: {total_storage:.2f} GB")
    
    # Storage costs
    s3_std_cost = total_s3_std * s3_standard_price
    s3_int_cost = weekly_s3_int * pricing['S3 Intelligent']
    glacier_ir_cost = offaccount_glacier * pricing['Glacier IR']
    total_storage_cost = s3_std_cost + s3_int_cost + glacier_ir_cost
    
    print(f"\nStorage Costs:")
    print(f"  S3 Standard: {total_s3_std:.2f} × ${s3_standard_price} = ${s3_std_cost:.4f}")
    print(f"  S3 Intelligent: {weekly_s3_int:.2f} × ${pricing['S3 Intelligent']} = ${s3_int_cost:.4f}")
    print(f"  Glacier IR: {offaccount_glacier:.2f} × ${pricing['Glacier IR']} = ${glacier_ir_cost:.4f}")
    print(f"  Total Storage Cost: ${total_storage_cost:.4f}")
    
    # Early deletion penalty calculation
    glacier_ir_min_days = config['min_duration']['Glacier IR']  # 90 days
    penalty_days = glacier_ir_min_days - offaccount_retention  # 90 - 15 = 75 days
    early_deletion_penalty = offaccount_glacier * pricing['Glacier IR'] * (penalty_days / 30)
    
    print(f"\nEarly Deletion Penalty:")
    print(f"  Off-account Glacier IR: {offaccount_glacier:.2f} GB × ${pricing['Glacier IR']} × ({glacier_ir_min_days}-{offaccount_retention})/30 = ${early_deletion_penalty:.4f}")
    
    total_monthly_cost = total_storage_cost + early_deletion_penalty
    print(f"  Total Monthly Cost: ${total_storage_cost:.4f} + ${early_deletion_penalty:.4f} = ${total_monthly_cost:.4f}")
    
    # AWS Backup cost
    aws_backup_cost = total_storage * aws_backup_pricing['storage_cost_per_gb_month']
    print(f"\nAWS Backup Cost: {total_storage:.2f} × ${aws_backup_pricing['storage_cost_per_gb_month']} = ${aws_backup_cost:.4f}")
    
    # Cost difference
    cost_difference = total_monthly_cost - aws_backup_cost
    print(f"Cost Difference: ${total_monthly_cost:.4f} - ${aws_backup_cost:.4f} = ${cost_difference:.4f}")
    
    print(f"\n{'='*50}")
    print("Expected Excel Formulas:")
    print("B2 (Total Backups): =(Configuration!B9+Configuration!B10+Configuration!B11+Configuration!B12)*Configuration!B14")
    print("C2 (S3 Standard): Should show", f"{total_s3_std:.2f} GB")
    print("D2 (S3 Intelligent): Should show", f"{weekly_s3_int:.2f} GB") 
    print("F2 (Glacier IR): Should show", f"{offaccount_glacier:.2f} GB")
    print("H2 (Total Storage): Should show", f"{total_storage:.2f} GB")
    print("I2 (Storage Cost): Should show", f"${total_storage_cost:.4f}")
    print("J2 (Early Deletion Penalty): Should show", f"${early_deletion_penalty:.4f}")
    print("K2 (Total Monthly Cost): Should show", f"${total_monthly_cost:.4f}")
    print("L2 (AWS Backup Cost): Should show", f"${aws_backup_cost:.4f}")
    print("M2 (Cost Difference): Should show", f"${cost_difference:.4f}")

if __name__ == "__main__":
    verify_cost_calculations()