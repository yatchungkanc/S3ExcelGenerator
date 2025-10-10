#!/usr/bin/env python3
"""
Setup script for S3 Cost Calculator
"""
import subprocess
import sys

def install_requirements():
    """Install required packages"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        print("✓ openpyxl installed successfully")
        return True
    except subprocess.CalledProcessError:
        print("✗ Failed to install openpyxl")
        print("You can still use the CSV version: python3 s3_cost_calculator_csv.py")
        return False

def main():
    print("S3 Cost Calculator Setup")
    print("=" * 30)
    
    if install_requirements():
        print("\nSetup complete! You can now run:")
        print("  python3 s3_cost_calculator.py      # Generate Excel file")
        print("  python3 s3_cost_calculator_csv.py  # Generate CSV file")
    else:
        print("\nFallback option available:")
        print("  python3 s3_cost_calculator_csv.py  # Generate CSV file")

if __name__ == "__main__":
    main()