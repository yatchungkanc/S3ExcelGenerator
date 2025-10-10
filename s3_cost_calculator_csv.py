#!/usr/bin/env python3
import csv
import datetime

def create_s3_cost_calculator_csv():
    filename = f"S3_Cost_Calculator_{datetime.datetime.now().strftime('%Y%m%d')}.csv"
    
    # Default parameters
    base_size = 100  # GB
    monthly_increment = 50  # GB
    storage_period = 12  # months
    ia_transition_days = 30
    glacier_transition_days = 90
    
    # S3 Pricing (USD per GB/month)
    standard_price = 0.023
    ia_price = 0.0125
    glacier_price = 0.004
    
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        
        # Header
        writer.writerow(["S3 Storage Cost Calculator (1 Year)"])
        writer.writerow([])
        
        # Input parameters
        writer.writerow(["Input Parameters"])
        writer.writerow(["Base Size (GB)", base_size])
        writer.writerow(["Monthly Incremental Size (GB)", monthly_increment])
        writer.writerow(["Storage Period (months)", storage_period])
        writer.writerow(["Transition to IA after (days)", ia_transition_days])
        writer.writerow(["Transition to Glacier after (days)", glacier_transition_days])
        writer.writerow([])
        
        # Pricing
        writer.writerow(["S3 Pricing (USD per GB/month)"])
        writer.writerow(["Standard", standard_price])
        writer.writerow(["Standard-IA", ia_price])
        writer.writerow(["Glacier Flexible", glacier_price])
        writer.writerow([])
        
        # Calculation headers
        writer.writerow(["Month", "Total Size (GB)", "Standard (GB)", "Standard-IA (GB)", 
                        "Glacier (GB)", "Standard Cost", "IA Cost", "Glacier Cost", "Monthly Total"])
        
        total_annual_cost = 0
        
        # Monthly calculations
        for month in range(1, 13):
            total_size = base_size + (monthly_increment * (month - 1))
            
            # Determine storage distribution based on age
            days_elapsed = month * 30
            
            # Simple transition logic
            if days_elapsed <= ia_transition_days:
                standard_gb = total_size
                ia_gb = 0
                glacier_gb = 0
            elif days_elapsed <= glacier_transition_days:
                standard_gb = 0
                ia_gb = total_size
                glacier_gb = 0
            else:
                standard_gb = 0
                ia_gb = 0
                glacier_gb = total_size
            
            # Cost calculations
            standard_cost = standard_gb * standard_price
            ia_cost = ia_gb * ia_price
            glacier_cost = glacier_gb * glacier_price
            monthly_total = standard_cost + ia_cost + glacier_cost
            
            total_annual_cost += monthly_total
            
            writer.writerow([
                month, 
                f"{total_size:.2f}",
                f"{standard_gb:.2f}",
                f"{ia_gb:.2f}",
                f"{glacier_gb:.2f}",
                f"${standard_cost:.2f}",
                f"${ia_cost:.2f}",
                f"${glacier_cost:.2f}",
                f"${monthly_total:.2f}"
            ])
        
        writer.writerow([])
        writer.writerow(["Summary"])
        writer.writerow(["Total Annual Cost", f"${total_annual_cost:.2f}"])
        writer.writerow(["Average Monthly Cost", f"${total_annual_cost/12:.2f}"])
        writer.writerow(["Final Storage Size", f"{base_size + (monthly_increment * 11):.2f} GB"])
        
        writer.writerow([])
        writer.writerow(["Instructions:"])
        writer.writerow(["1. Modify parameters in the script to adjust calculations"])
        writer.writerow(["2. Pricing reflects current AWS S3 rates"])
        writer.writerow(["3. Transitions occur based on data age in days"])
        writer.writerow(["4. Import this CSV into Excel for advanced formatting"])
    
    return filename

if __name__ == "__main__":
    filename = create_s3_cost_calculator_csv()
    print(f"CSV file created: {filename}")
    print("Import this file into Excel or Google Sheets for better formatting and calculations.")