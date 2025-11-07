#!/usr/bin/env python3
import json

def test_calculations():
    with open('backup_config.json', 'r') as f:
        config = json.load(f)
    
    print("=== Testing Calculation Formulas ===")
    
    # Test Month 1 with Tier 1 configuration
    print("\nTier 1 Configuration:")
    print("Hourly: 24 backups, Daily: 7 backups, Weekly: 6 backups, Off-account: 2 backups")
    print("Storage size: 100 GB")
    
    month = 1
    storage_size = 100
    
    # Tier 1 counts
    hourly_count = 24
    daily_count = 7
    weekly_count = 6
    offaccount_count = 2
    
    # Retention periods
    hourly_retention = 7
    daily_retention = 30
    tier_retention = 15  # S3 Standard tier retention
    
    # Calculate total backups
    total_backups = (hourly_count + daily_count + weekly_count + offaccount_count) * storage_size
    print(f"Total backups: ({hourly_count}+{daily_count}+{weekly_count}+{offaccount_count}) × {storage_size} = {total_backups} GB")
    
    # Calculate S3 Standard storage (all using tier 1 = S3 Standard)
    hourly_storage = hourly_count * storage_size * min(hourly_retention, month * 30) / 30
    daily_storage = daily_count * storage_size * min(daily_retention, month * 30) / 30
    weekly_storage = weekly_count * storage_size * min(tier_retention, month * 30) / 30
    offaccount_storage = offaccount_count * storage_size * min(tier_retention, month * 30) / 30
    
    total_s3_standard = hourly_storage + daily_storage + weekly_storage + offaccount_storage
    
    print(f"S3 Standard storage:")
    print(f"  Hourly: {hourly_count} × {storage_size} × min({hourly_retention}, {month*30})/30 = {hourly_storage:.2f} GB")
    print(f"  Daily: {daily_count} × {storage_size} × min({daily_retention}, {month*30})/30 = {daily_storage:.2f} GB")
    print(f"  Weekly: {weekly_count} × {storage_size} × min({tier_retention}, {month*30})/30 = {weekly_storage:.2f} GB")
    print(f"  Off-account: {offaccount_count} × {storage_size} × min({tier_retention}, {month*30})/30 = {offaccount_storage:.2f} GB")
    print(f"  Total S3 Standard: {total_s3_standard:.2f} GB")
    
    # AWS Backup cost
    aws_backup_price = 0.05
    aws_backup_cost = total_s3_standard * aws_backup_price
    print(f"AWS Backup cost: {total_s3_standard:.2f} × ${aws_backup_price} = ${aws_backup_cost:.4f}")
    
    print("\n=== Formula Verification ===")
    print("Configuration sheet formulas should reference:")
    print("- B9 (Hourly count): =IF(B3=\"Tier 1\",G25,IF(B3=\"Tier 2\",H25,...))")
    print("- B10 (Daily count): =IF(B3=\"Tier 1\",G26,IF(B3=\"Tier 2\",H26,...))")
    print("- Calculation sheet should use Configuration!B9, B10, B11, B12 for counts")

if __name__ == "__main__":
    test_calculations()