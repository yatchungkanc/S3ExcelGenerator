#!/usr/bin/env python3
from datetime import datetime
import pandas as pd
import json
import os

def load_config(config_file='backup_config.json'):
    """Load configuration from JSON file"""
    if not os.path.exists(config_file):
        raise FileNotFoundError(f"Config file {config_file} not found")
    
    with open(config_file, 'r') as f:
        return json.load(f)

def create_backup_cost_calculator(config_file='backup_config.json'):
    config = load_config(config_file)
    
    pricing = config['pricing']
    min_duration = config['min_duration']
    retrieval_costs = config['retrieval_costs']
    backup_config = config['backup_config']
    storage_size_gb = config['storage_size_gb']
    tier_retention = config['tier_retention']
    aws_backup_pricing = config.get('aws_backup_pricing', {'storage_cost_per_gb_month': 0.05, 'retrieval_cost_per_gb': 0.02})
    
    # Create Excel writer
    filename = f"Backup_Cost_Calculator_{datetime.now().strftime('%Y%m%d')}.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        
        # Create structured configuration layout
        workbook = writer.book
        config_sheet = workbook.create_sheet('Configuration')
        
        # Section 1: Backup Configuration
        config_sheet['A1'] = 'BACKUP CONFIGURATION'
        config_sheet['A3'] = 'Backup Type'
        config_sheet['B3'] = 'Count'
        config_sheet['C3'] = 'Storage Tier'
        config_sheet['D3'] = 'Tier Name'
        config_sheet['E3'] = 'Retention (days)'
        
        config_sheet['A4'] = 'Hourly backups'
        config_sheet['B4'] = backup_config['hourly']['count']
        config_sheet['C4'] = backup_config['hourly']['tier']
        config_sheet['D4'] = '=IF(C4=1,"S3 Standard",IF(C4=2,"S3 Intelligent",IF(C4=3,"S3-IA",IF(C4=4,"Glacier IR",IF(C4=5,"Glacier DA","INVALID")))))'
        config_sheet['E4'] = backup_config['hourly']['retention']
        
        config_sheet['A5'] = 'Daily backups'
        config_sheet['B5'] = backup_config['daily']['count']
        config_sheet['C5'] = backup_config['daily']['tier']
        config_sheet['D5'] = '=IF(C5=1,"S3 Standard",IF(C5=2,"S3 Intelligent",IF(C5=3,"S3-IA",IF(C5=4,"Glacier IR",IF(C5=5,"Glacier DA","INVALID")))))'
        config_sheet['E5'] = backup_config['daily']['retention']
        
        config_sheet['A6'] = 'Weekly backups'
        config_sheet['B6'] = backup_config['weekly']['count']
        config_sheet['C6'] = backup_config['weekly']['tier']
        config_sheet['D6'] = '=IF(C6=1,"S3 Standard",IF(C6=2,"S3 Intelligent",IF(C6=3,"S3-IA",IF(C6=4,"Glacier IR",IF(C6=5,"Glacier DA","INVALID")))))'
        
        config_sheet['A7'] = 'Off-account backups'
        config_sheet['B7'] = backup_config['off_account']['count']
        config_sheet['C7'] = backup_config['off_account']['tier']
        config_sheet['D7'] = '=IF(C7=1,"S3 Standard",IF(C7=2,"S3 Intelligent",IF(C7=3,"S3-IA",IF(C7=4,"Glacier IR",IF(C7=5,"Glacier DA","INVALID")))))'
        
        config_sheet['A9'] = 'Incremental storage size (GB)'
        config_sheet['B9'] = storage_size_gb
        
        # Section 2: Storage Tier Retention
        config_sheet['A11'] = 'STORAGE TIER RETENTION'
        config_sheet['A13'] = 'Storage Tier'
        config_sheet['B13'] = 'Retention (days)'
        config_sheet['C13'] = 'Min Required'
        config_sheet['D13'] = 'Status'
        
        config_sheet['A14'] = 'S3 Standard'
        config_sheet['B14'] = tier_retention['S3 Standard']
        config_sheet['C14'] = min_duration['S3 Standard']
        config_sheet['D14'] = '=IF(B14>=C14,"OK","PENALTY")'
        
        config_sheet['A15'] = 'S3 Intelligent'
        config_sheet['B15'] = tier_retention['S3 Intelligent']
        config_sheet['C15'] = min_duration['S3 Intelligent']
        config_sheet['D15'] = '=IF(B15>=C15,"OK","PENALTY")'
        
        config_sheet['A16'] = 'S3-IA'
        config_sheet['B16'] = tier_retention['S3-IA']
        config_sheet['C16'] = min_duration['S3-IA']
        config_sheet['D16'] = '=IF(B16>=C16,"OK","PENALTY")'
        
        config_sheet['A17'] = 'Glacier IR'
        config_sheet['B17'] = tier_retention['Glacier IR']
        config_sheet['C17'] = min_duration['Glacier IR']
        config_sheet['D17'] = '=IF(B17>=C17,"OK","PENALTY")'
        
        config_sheet['A18'] = 'Glacier Deep Archive'
        config_sheet['B18'] = tier_retention['Glacier Deep Archive']
        config_sheet['C18'] = min_duration['Glacier Deep Archive']
        config_sheet['D18'] = '=IF(B18>=C18,"OK","PENALTY")'
        
        # Section 3: Data Flow Visualization - FIXED with all 4 backup types
        config_sheet['F1'] = 'DATA FLOW VISUALIZATION'
        config_sheet['F3'] = 'Backup Type'
        config_sheet['G3'] = 'Storage Tier'
        config_sheet['H3'] = 'Retention'
        
        config_sheet['F4'] = 'Hourly'
        config_sheet['G4'] = '=D4'
        config_sheet['H4'] = '=IF(E4>0,CONCATENATE(E4," days"),"SKIP")'
        
        config_sheet['F5'] = 'Daily'
        config_sheet['G5'] = '=D5'
        config_sheet['H5'] = '=IF(E5>0,CONCATENATE(E5," days"),"SKIP")'
        
        config_sheet['F6'] = 'Weekly'
        config_sheet['G6'] = '=D6'
        config_sheet['H6'] = '=IF(C6=1,IF(B14>0,CONCATENATE(B14," days"),"SKIP"),IF(C6=2,IF(B15>0,CONCATENATE(B15," days"),"SKIP"),IF(C6=3,IF(B16>0,CONCATENATE(B16," days"),"SKIP"),IF(C6=4,IF(B17>0,CONCATENATE(B17," days"),"SKIP"),IF(C6=5,IF(B18>0,CONCATENATE(B18," days"),"SKIP"),"INVALID")))))'
        
        config_sheet['F7'] = 'Off-account'
        config_sheet['G7'] = '=D7'
        config_sheet['H7'] = '=IF(C7=1,IF(B14>0,CONCATENATE(B14," days"),"SKIP"),IF(C7=2,IF(B15>0,CONCATENATE(B15," days"),"SKIP"),IF(C7=3,IF(B16>0,CONCATENATE(B16," days"),"SKIP"),IF(C7=4,IF(B17>0,CONCATENATE(B17," days"),"SKIP"),IF(C7=5,IF(B18>0,CONCATENATE(B18," days"),"SKIP"),"INVALID")))))'
        
        # Section 4: Tier Codes Reference
        config_sheet['F9'] = 'TIER CODES'
        config_sheet['F10'] = '1 = S3 Standard'
        config_sheet['F11'] = '2 = S3 Intelligent'
        config_sheet['F12'] = '3 = S3-IA'
        config_sheet['F13'] = '4 = Glacier IR'
        config_sheet['F14'] = '5 = Glacier DA'
        
        # Section 5: Backup Tier Reference Table
        config_sheet['F16'] = 'BACKUP TIER REFERENCE'
        config_sheet['F18'] = 'Backup Type'
        config_sheet['G18'] = 'Tier 1'
        config_sheet['H18'] = 'Tier 2'
        config_sheet['I18'] = 'Tier 3'
        config_sheet['J18'] = 'Tier 4'
        
        config_sheet['F19'] = 'Hourly backups'
        config_sheet['G19'] = 24
        config_sheet['H19'] = 24
        config_sheet['I19'] = 6
        config_sheet['J19'] = 0
        
        config_sheet['F20'] = 'Daily backups'
        config_sheet['G20'] = 7
        config_sheet['H20'] = 7
        config_sheet['I20'] = 7
        config_sheet['J20'] = 7
        
        config_sheet['F21'] = 'Weekly backups'
        config_sheet['G21'] = 6
        config_sheet['H21'] = 6
        config_sheet['I21'] = 4
        config_sheet['J21'] = 2
        
        config_sheet['F22'] = 'Off-account backups'
        config_sheet['G22'] = 2
        config_sheet['H22'] = 2
        config_sheet['I22'] = 2
        config_sheet['J22'] = 2
        
        # Format editable cells
        from openpyxl.styles import PatternFill, Font
        blue_fill = PatternFill(start_color='E7F3FF', end_color='E7F3FF', fill_type='solid')
        bold_font = Font(bold=True)
        
        config_sheet['A1'].font = bold_font
        config_sheet['A11'].font = bold_font
        config_sheet['F1'].font = bold_font
        config_sheet['F9'].font = bold_font
        config_sheet['F16'].font = bold_font
        
        # Make backup counts and tiers editable
        for row in [4, 5, 6, 7, 9]:
            config_sheet[f'B{row}'].fill = blue_fill
        for row in [4, 5, 6, 7]:
            config_sheet[f'C{row}'].fill = blue_fill
        # Make hourly and daily retention editable
        for row in [4, 5]:
            config_sheet[f'E{row}'].fill = blue_fill
        
        # Make retention days editable
        for row in range(14, 19):
            config_sheet[f'B{row}'].fill = blue_fill
        
        # Pricing sheet
        pricing_data = {
            'Storage Tier': list(pricing.keys()) + ['AWS Backup'],
            'Price per GB/month (USD)': list(pricing.values()) + [aws_backup_pricing['storage_cost_per_gb_month']],
            'Min Duration (days)': [min_duration[tier] for tier in pricing.keys()] + [0],
            'Retrieval Cost per GB (USD)': [retrieval_costs[tier] for tier in pricing.keys()] + [aws_backup_pricing['retrieval_cost_per_gb']]
        }
        
        pricing_df = pd.DataFrame(pricing_data)
        pricing_df.to_excel(writer, sheet_name='Pricing', index=False)
        
        # Create calculation sheet with formulas
        calc_data = {
            'Month': [f'Month {i}' for i in range(1, 13)],
            'Total Backups (GB)': [''] * 12,
            'S3 Standard (GB)': [''] * 12,
            'S3 Intelligent (GB)': [''] * 12,
            'S3-IA (GB)': [''] * 12,
            'Glacier IR (GB)': [''] * 12,
            'Glacier DA (GB)': [''] * 12,
            'Total Storage (GB)': [''] * 12,
            'Storage Cost (USD)': [''] * 12,
            'Early Deletion Penalty (USD)': [''] * 12,
            'Total Monthly Cost (USD)': [''] * 12,
            'AWS Backup Cost (USD)': [''] * 12,
            'Cost Difference (USD)': [''] * 12
        }
        
        calc_df = pd.DataFrame(calc_data)
        calc_df.to_excel(writer, sheet_name='Cost Calculation', index=False)
        
        # Add formulas to the calculation sheet
        workbook = writer.book
        calc_sheet = workbook['Cost Calculation']
        
        for row in range(2, 14):  # Rows 2-13 (months 1-12)
            month = row - 1
            
            # Total backups in GB
            calc_sheet[f'B{row}'] = f'=(Configuration!B4+Configuration!B5+Configuration!B6+Configuration!B7)*Configuration!B9'
            
            # S3 Standard storage - hourly(tier=1) + daily(tier=1) + weekly(tier=1) + off-account(tier=1)
            calc_sheet[f'C{row}'] = f'=(IF(Configuration!C4=1,Configuration!B4*Configuration!B9*MIN(Configuration!E4,{month}*30)/30,0)+IF(Configuration!C5=1,Configuration!B5*Configuration!B9*MIN(Configuration!E5,{month}*30)/30,0)+IF(Configuration!C6=1,Configuration!B6*Configuration!B9*MIN(Configuration!B14,{month}*30)/30,0)+IF(Configuration!C7=1,Configuration!B7*Configuration!B9*MIN(Configuration!B14,{month}*30)/30,0))'
            
            # S3 Intelligent storage - hourly(tier=2) + daily(tier=2) + weekly(tier=2) + off-account(tier=2)
            calc_sheet[f'D{row}'] = f'=IF(Configuration!B15>0,(IF(Configuration!C4=2,Configuration!B4,0)+IF(Configuration!C5=2,Configuration!B5,0)+IF(Configuration!C6=2,Configuration!B6,0)+IF(Configuration!C7=2,Configuration!B7,0))*Configuration!B9*MIN(Configuration!B15,{month}*30)/30,0)'
            
            # S3-IA storage - hourly(tier=3) + daily(tier=3) + weekly(tier=3) + off-account(tier=3)
            calc_sheet[f'E{row}'] = f'=IF(Configuration!B16>0,(IF(Configuration!C4=3,Configuration!B4,0)+IF(Configuration!C5=3,Configuration!B5,0)+IF(Configuration!C6=3,Configuration!B6,0)+IF(Configuration!C7=3,Configuration!B7,0))*Configuration!B9*MIN(Configuration!B16,{month}*30)/30,0)'
            
            # Glacier IR storage - hourly(tier=4) + daily(tier=4) + weekly(tier=4) + off-account(tier=4)
            calc_sheet[f'F{row}'] = f'=IF(Configuration!B17>0,(IF(Configuration!C4=4,Configuration!B4,0)+IF(Configuration!C5=4,Configuration!B5,0)+IF(Configuration!C6=4,Configuration!B6,0)+IF(Configuration!C7=4,Configuration!B7,0))*Configuration!B9*MIN(Configuration!B17,{month}*30)/30,0)'
            
            # Glacier Deep Archive storage - hourly(tier=5) + daily(tier=5) + weekly(tier=5) + off-account(tier=5)
            calc_sheet[f'G{row}'] = f'=IF(Configuration!B18>0,(IF(Configuration!C4=5,Configuration!B4,0)+IF(Configuration!C5=5,Configuration!B5,0)+IF(Configuration!C6=5,Configuration!B6,0)+IF(Configuration!C7=5,Configuration!B7,0))*Configuration!B9*MIN(Configuration!B18,{month}*30)/30,0)'
            
            # Total storage
            calc_sheet[f'H{row}'] = f'=C{row}+D{row}+E{row}+F{row}+G{row}'
            
            # Storage cost (normal pricing)
            calc_sheet[f'I{row}'] = f'=(IF(Configuration!C4=1,Configuration!B4*Configuration!B9*MIN(Configuration!E4,{month}*30)/30*Pricing!B2*(Configuration!E4/30),0)+IF(Configuration!C5=1,Configuration!B5*Configuration!B9*MIN(Configuration!E5,{month}*30)/30*Pricing!B2*(Configuration!E5/30),0)+IF(Configuration!C6=1,Configuration!B6*Configuration!B9*MIN(Configuration!B14,{month}*30)/30*Pricing!B2*(Configuration!B14/30),0)+IF(Configuration!C7=1,Configuration!B7*Configuration!B9*MIN(Configuration!B14,{month}*30)/30*Pricing!B2*(Configuration!B14/30),0))+D{row}*Pricing!B3*(Configuration!B15/30)+E{row}*Pricing!B4*(Configuration!B16/30)+F{row}*Pricing!B5*(Configuration!B17/30)+G{row}*Pricing!B6*(Configuration!B18/30)'
            
            # Early deletion penalty for retention < minimum storage duration
            calc_sheet[f'J{row}'] = f'=IF(Configuration!B16<Pricing!C4,E{row}*Pricing!B4*((Pricing!C4-Configuration!B16)/30),0)+IF(Configuration!B17<Pricing!C5,F{row}*Pricing!B5*((Pricing!C5-Configuration!B17)/30),0)+IF(Configuration!B18<Pricing!C6,G{row}*Pricing!B6*((Pricing!C6-Configuration!B18)/30),0)'
            
            # Total monthly cost (storage + penalties)
            calc_sheet[f'K{row}'] = f'=I{row}+J{row}'
            
            # AWS Backup cost (total storage * AWS Backup price)
            calc_sheet[f'L{row}'] = f'=H{row}*Pricing!B7'
            
            # Cost difference (S3 Direct - AWS Backup)
            calc_sheet[f'M{row}'] = f'=K{row}-L{row}'
        
        # Add annual summary
        calc_sheet['A14'] = 'Annual Total'
        calc_sheet['B14'] = '=SUM(B2:B13)'
        calc_sheet['C14'] = '=SUM(C2:C13)'
        calc_sheet['D14'] = '=SUM(D2:D13)'
        calc_sheet['E14'] = '=SUM(E2:E13)'
        calc_sheet['F14'] = '=SUM(F2:F13)'
        calc_sheet['G14'] = '=SUM(G2:G13)'
        calc_sheet['H14'] = '=SUM(H2:H13)'
        calc_sheet['I14'] = '=SUM(I2:I13)'
        calc_sheet['J14'] = '=SUM(J2:J13)'
        calc_sheet['K14'] = '=SUM(K2:K13)'
        calc_sheet['L14'] = '=SUM(L2:L13)'
        calc_sheet['M14'] = '=SUM(M2:M13)'
        
        # Format pricing sheet values as editable
        pricing_sheet = workbook['Pricing']
        for row in range(2, 8):  # Updated to include AWS Backup row
            pricing_sheet[f'B{row}'].fill = blue_fill
        
        # Auto-adjust column widths
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 20)
                sheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"Backup cost calculator created: {filename}")
    return filename

if __name__ == "__main__":
    create_backup_cost_calculator()