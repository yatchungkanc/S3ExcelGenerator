#!/usr/bin/env python3
from datetime import datetime
import pandas as pd

def create_backup_cost_calculator():
    # Storage tier pricing (USD per GB/month)
    pricing = {
        'S3 Standard': 0.023,
        'S3 Intelligent': 0.0125,
        'S3-IA': 0.0125,
        'Glacier IR': 0.004,
        'Glacier Deep Archive': 0.00099
    }
    
    # Minimum storage duration (days)
    min_duration = {
        'S3 Standard': 0,
        'S3 Intelligent': 0,
        'S3-IA': 30,
        'Glacier IR': 90,
        'Glacier Deep Archive': 180
    }
    
    # Retrieval costs (USD per GB)
    retrieval_costs = {
        'S3 Standard': 0,
        'S3 Intelligent': 0,
        'S3-IA': 0.01,
        'Glacier IR': 0.03,
        'Glacier Deep Archive': 0.02
    }
    
    # Create Excel writer
    filename = f"Backup_Cost_Calculator_{datetime.now().strftime('%Y%m%d')}.xlsx"
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        
        # Configuration sheet
        config_data = {
            'Parameter': [
                'Hourly backups to keep',
                'Daily backups to keep', 
                'Weekly backups to keep',
                'Monthly backups to keep',
                'Off-account backups per year',
                'Incremental storage size (GB)',
                'S3 Standard retention (days)',
                'S3 Intelligent retention (days)',
                'S3-IA retention (days)',
                'Glacier IR retention (days)',
                'Glacier Deep Archive retention (days)'
            ],
            'Value': [24, 7, 4, 12, 4, 100, 30, 60, 90, 180, 365],
            'Description': [
                'Number of hourly backups to keep',
                'Number of daily backups to keep',
                'Number of weekly backups to keep', 
                'Number of monthly backups to keep',
                'Number of off-account backups yearly',
                'Size of each incremental backup',
                'Days to retain in S3 Standard (0=skip)',
                'Days to retain in S3 Intelligent (0=skip)',
                'Days to retain in S3-IA (0=skip, min 30)',
                'Days to retain in Glacier IR (0=skip, min 90)',
                'Days to retain in Glacier Deep Archive (0=skip, min 180)'
            ]
        }
        
        config_df = pd.DataFrame(config_data)
        config_df.to_excel(writer, sheet_name='Configuration', index=False)
        
        # Add explanatory notes to Configuration sheet
        workbook = writer.book
        config_sheet = workbook['Configuration']
        
        # Add backup distribution logic explanation
        config_sheet['A14'] = 'BACKUP DISTRIBUTION LOGIC:'
        config_sheet['A15'] = '• Hourly + Daily backups: Always stored in S3 Standard'
        config_sheet['A16'] = '• Weekly + Monthly + Off-account backups: Flow through tiers based on retention'
        config_sheet['A17'] = '• If a tier has 0 retention days, backups skip to next available tier'
        config_sheet['A18'] = '• Example: If S3 Intelligent = 0, backups go directly to S3-IA'
        config_sheet['A19'] = '• Monthly backups grow from 1 to 12 over the year'
        config_sheet['A20'] = '• Off-account backups only appear after month 12'
        
        # Add visual data transition flow
        config_sheet['E14'] = 'DATA TRANSITION FLOW:'
        config_sheet['E15'] = 'Hourly + Daily →'
        config_sheet['F15'] = 'S3 Standard'
        config_sheet['G15'] = '=IF(B8>0,CONCATENATE("(",B8," days)"),"(SKIP)")'
        
        config_sheet['E17'] = 'Weekly + Monthly + Off-account →'
        config_sheet['F17'] = '=IF(B9>0,"S3 Intelligent",IF(B10>0,"S3-IA",IF(B11>0,"Glacier IR",IF(B12>0,"Glacier DA","NONE"))))'
        config_sheet['G17'] = '=IF(B9>0,CONCATENATE("(",B9," days)"),IF(B10>0,CONCATENATE("(",B10," days)"),IF(B11>0,CONCATENATE("(",B11," days)"),IF(B12>0,CONCATENATE("(",B12," days)"),"(SKIP)"))))'
        
        config_sheet['F18'] = '↓'
        config_sheet['F19'] = '=IF(AND(B9>0,B10>0),"S3-IA",IF(AND(B9=0,B10>0,B11>0),"Glacier IR","END"))'
        config_sheet['G19'] = '=IF(AND(B9>0,B10>0),CONCATENATE("(",B10," days)"),IF(AND(B9=0,B10>0,B11>0),CONCATENATE("(",B11," days)"),""))'
        
        config_sheet['F20'] = '↓'
        config_sheet['F21'] = '=IF(AND(B9>0,B10>0,B11>0),"Glacier IR","END")'
        config_sheet['G21'] = '=IF(AND(B9>0,B10>0,B11>0),CONCATENATE("(",B11," days)"),"")'
        
        config_sheet['F22'] = '↓'
        config_sheet['F23'] = '=IF(AND(B9>0,B10>0,B11>0,B12>0),"Glacier DA","END")'
        config_sheet['G23'] = '=IF(AND(B9>0,B10>0,B11>0,B12>0),CONCATENATE("(",B12," days)"),"")'
        
        # Pricing sheet
        pricing_data = {
            'Storage Tier': list(pricing.keys()),
            'Price per GB/month (USD)': list(pricing.values()),
            'Min Duration (days)': [min_duration[tier] for tier in pricing.keys()],
            'Retrieval Cost per GB (USD)': [retrieval_costs[tier] for tier in pricing.keys()]
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
            'Monthly Cost (USD)': [''] * 12
        }
        
        calc_df = pd.DataFrame(calc_data)
        calc_df.to_excel(writer, sheet_name='Cost Calculation', index=False)
        
        # Add formulas to the calculation sheet
        workbook = writer.book
        calc_sheet = workbook['Cost Calculation']
        
        for row in range(2, 14):  # Rows 2-13 (months 1-12)
            month = row - 1
            
            # Total backups in GB
            calc_sheet[f'B{row}'] = f'=(Configuration!B2+Configuration!B3+Configuration!B4+MIN(Configuration!B5,{month})+IF({month}>=12,Configuration!B6,0))*Configuration!B7'
            
            # S3 Standard storage - hourly + daily backups accumulated over retention period
            calc_sheet[f'C{row}'] = f'=IF(Configuration!B8>0,(Configuration!B2+Configuration!B3)*Configuration!B7*MIN(Configuration!B8,{month}*30)/30,0)'
            
            # S3 Intelligent storage - weekly/monthly/off-account accumulated over retention period
            calc_sheet[f'D{row}'] = f'=IF(Configuration!B9>0,(Configuration!B4+MIN(Configuration!B5,{month})+IF({month}>=12,Configuration!B6,0))*Configuration!B7*MIN(Configuration!B9,{month}*30)/30,0)'
            
            # S3-IA storage - weekly/monthly/off-account if retention > 0 and not in Intelligent
            calc_sheet[f'E{row}'] = f'=IF(AND(Configuration!B10>0,Configuration!B9=0),(Configuration!B4+MIN(Configuration!B5,{month})+IF({month}>=12,Configuration!B6,0))*Configuration!B7*MIN(Configuration!B10,{month}*30)/30,0)'
            
            # Glacier IR storage - weekly/monthly/off-account if retention > 0 and not in previous tiers
            calc_sheet[f'F{row}'] = f'=IF(AND(Configuration!B11>0,Configuration!B9=0,Configuration!B10=0),(Configuration!B4+MIN(Configuration!B5,{month})+IF({month}>=12,Configuration!B6,0))*Configuration!B7*MIN(Configuration!B11,{month}*30)/30,0)'
            
            # Glacier Deep Archive storage - weekly/monthly/off-account if retention > 0 and not in previous tiers
            calc_sheet[f'G{row}'] = f'=IF(AND(Configuration!B12>0,Configuration!B9=0,Configuration!B10=0,Configuration!B11=0),(Configuration!B4+MIN(Configuration!B5,{month})+IF({month}>=12,Configuration!B6,0))*Configuration!B7*MIN(Configuration!B12,{month}*30)/30,0)'
            
            # Total storage
            calc_sheet[f'H{row}'] = f'=C{row}+D{row}+E{row}+F{row}+G{row}'
            
            # Monthly cost (prorated by retention days)
            calc_sheet[f'I{row}'] = f'=C{row}*Pricing!B2*(Configuration!B8/30)+D{row}*Pricing!B3*(Configuration!B9/30)+E{row}*Pricing!B4*(Configuration!B10/30)+F{row}*Pricing!B5*(Configuration!B11/30)+G{row}*Pricing!B6*(Configuration!B12/30)'
        
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
        
        # Format configuration sheet values as editable
        from openpyxl.styles import PatternFill, Font
        blue_fill = PatternFill(start_color='E7F3FF', end_color='E7F3FF', fill_type='solid')
        bold_font = Font(bold=True)
        
        for row in range(2, 12):
            config_sheet[f'B{row}'].fill = blue_fill
        
        # Format explanation headers
        config_sheet['A14'].font = bold_font
        config_sheet['E14'].font = bold_font
        
        # Format pricing sheet values as editable
        pricing_sheet = workbook['Pricing']
        for row in range(2, 7):
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