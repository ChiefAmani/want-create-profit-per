import xlsxwriter

def create_calculator():
    workbook = xlsxwriter.Workbook('Profit Per Job Calculator.xlsx')
    
    # --- FORMATTING ---
    header_format = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
    sub_header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
    currency_format = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
    percent_format = workbook.add_format({'num_format': '0.0%', 'border': 1})
    number_format = workbook.add_format({'num_format': '0.00', 'border': 1})
    input_format = workbook.add_format({'bg_color': '#E2EFDA', 'border': 1}) # Greenish for inputs
    input_currency = workbook.add_format({'bg_color': '#E2EFDA', 'num_format': '$#,##0.00', 'border': 1})
    input_percent = workbook.add_format({'bg_color': '#E2EFDA', 'num_format': '0.0%', 'border': 1})
    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})
    green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
    
    # --- SHEET: Settings & Costs ---
    ws_settings = workbook.add_worksheet('Settings & Costs')
    ws_settings.set_column('A:A', 35)
    ws_settings.set_column('B:B', 15)
    
    ws_settings.write('A1', 'Global Settings (Edit Green Cells)', header_format)
    ws_settings.write('B1', 'Value', header_format)
    
    settings_data = [
        ('Base Hourly Labor Rate', 20.00, input_currency),
        ('Labor Burden Multiplier (Taxes/Ins)', 1.25, input_format),
        ('Fuel Cost per Gallon', 3.50, input_currency),
        ('Vehicle MPG', 15, input_format),
        ('Equipment Wear & Tear per Job', 5.00, input_currency),
        ('Target Gross Margin', 0.60, input_percent),
        ('Monthly Fixed Overhead', 2000.00, input_currency)
    ]
    
    for row, (label, val, fmt) in enumerate(settings_data, start=1):
        ws_settings.write(row, 0, label, sub_header_format)
        ws_settings.write(row, 1, val, fmt)
            
    ws_settings.write('A10', 'Chemical Costs (Dynamic List)', header_format)
    ws_settings.write('B10', 'Cost per Unit', header_format)
    
    chem_data = [
        ('SH (Sodium Hypochlorite) / Gal', 4.50),
        ('Surfactant / Oz', 0.50),
        ('Degreaser / Oz', 0.75),
        ('Window Glide / Oz', 0.20),
        ('Oxalic Acid / Oz', 0.30)
    ]
    
    for i in range(20):
        row = 10 + i
        if i < len(chem_data):
            ws_settings.write(row, 0, chem_data[i][0], input_format)
            ws_settings.write(row, 1, chem_data[i][1], input_currency)
        else:
            ws_settings.write_blank(row, 0, '', input_format)
            ws_settings.write_blank(row, 1, '', input_currency)
            
    # --- SHEET: Job Calculator ---
    ws_job = workbook.add_worksheet('Job Calculator')
    ws_job.set_column('A:A', 35)
    ws_job.set_column('B:B', 20)
    ws_job.set_column('C:C', 20)
    
    ws_job.write('A1', 'Job Details (Edit Green Cells)', header_format)
    ws_job.write('B1', 'Input/Output', header_format)
    
    job_inputs = [
        ('Job Name/Address', '123 Main St', input_format),
        ('Estimated Time on Site (Hours)', 3.0, input_format),
        ('Number of Workers', 2, input_format),
        ('Round Trip Drive Time (Mins)', 45, input_format),
        ('Round Trip Distance (Miles)', 20, input_format),
        ('Quoted Price to Customer', 450.00, input_currency)
    ]
    
    for row, (label, val, fmt) in enumerate(job_inputs, start=1):
        ws_job.write(row, 0, label, sub_header_format)
        ws_job.write(row, 1, val, fmt)
        
    # Automated Pricing System
    ws_job.write('A9', 'Automated Pricing System', header_format)
    ws_job.write('B9', 'Input/Output', header_format)
    ws_job.write('A10', 'Area Size (Sq Ft)', sub_header_format)
    ws_job.write('B10', 2000, input_format)
    ws_job.write('A11', 'Productivity Rate (Sq Ft / Hour)', sub_header_format)
    ws_job.write('B11', 500, input_format)
    ws_job.write('A12', 'Estimated Hours (Calculated)', sub_header_format)
    ws_job.write_formula('B12', '=B10/B11', number_format)
    ws_job.write('A13', 'Supply Markup Percentage', sub_header_format)
    ws_job.write('B13', 0.20, input_percent)
    
    # Dynamic Chemical Usage Section
    ws_job.write('A15', 'Chemicals Used on Job', header_format)
    ws_job.write('B15', 'Quantity', header_format)
    ws_job.write('C15', 'Calculated Cost', header_format)
    
    for i in range(5):
        row = 15 + i
        ws_job.data_validation(row, 0, row, 0, {
            'validate': 'list',
            'source': '=\'Settings & Costs\'!$A$11:$A$30',
            'input_title': 'Select Chemical',
            'input_message': 'Choose a chemical from the Settings sheet'
        })
        if i == 0:
            ws_job.write(row, 0, 'SH (Sodium Hypochlorite) / Gal', input_format)
            ws_job.write(row, 1, 5, input_format)
        elif i == 1:
            ws_job.write(row, 0, 'Surfactant / Oz', input_format)
            ws_job.write(row, 1, 10, input_format)
        else:
            ws_job.write_blank(row, 0, '', input_format)
            ws_job.write_blank(row, 1, '', input_format)
            
        ws_job.write_formula(row, 2, f'=IF(A{row+1}<>"", VLOOKUP(A{row+1}, \'Settings & Costs\'!$A$11:$B$30, 2, FALSE) * B{row+1}, 0)', currency_format)
            
    ws_job.write('A22', 'Cost Breakdown', header_format)
    ws_job.write('B22', 'Amount', header_format)
    
    ws_job.write('A23', 'Total Labor Cost', sub_header_format)
    ws_job.write_formula('B23', '=(B3+(B5/60))*B4*\'Settings & Costs\'!B2*\'Settings & Costs\'!B3', currency_format)
    
    ws_job.write('A24', 'Fuel Cost', sub_header_format)
    ws_job.write_formula('B24', '=(B6/\'Settings & Costs\'!B5)*\'Settings & Costs\'!B4', currency_format)
    
    ws_job.write('A25', 'Chemical Cost (with Markup)', sub_header_format)
    ws_job.write_formula('B25', '=SUM(C16:C20)*(1+B13)', currency_format)
    
    ws_job.write('A26', 'Equipment Wear & Tear', sub_header_format)
    ws_job.write_formula('B26', '=\'Settings & Costs\'!B6', currency_format)
    
    ws_job.write('A27', 'TOTAL JOB COST', header_format)
    ws_job.write_formula('B27', '=SUM(B23:B26)', currency_format)
    
    # Price vs Labour breakdown
    ws_job.write('A29', 'Price vs Labour Breakdown', header_format)
    ws_job.write('B29', 'Metrics', header_format)
    ws_job.write('A30', 'Total Job Price', sub_header_format)
    ws_job.write_formula('B30', '=B7', currency_format)
    ws_job.write('A31', 'Total Labor Cost', sub_header_format)
    ws_job.write_formula('B31', '=B23', currency_format)
    ws_job.write('A32', 'Labor Cost % of Job Price', sub_header_format)
    ws_job.write_formula('B32', '=IF(B30>0, B31/B30, 0)', percent_format)
    
    ws_job.write('A34', 'Profitability Analysis', header_format)
    ws_job.write('B34', 'Metrics', header_format)
    
    ws_job.write('A35', 'Desired Profit Margin', sub_header_format)
    ws_job.write('B35', 0.60, input_percent)
    
    ws_job.write('A36', 'Gross Profit ($)', sub_header_format)
    ws_job.write_formula('B36', '=B7-B27', currency_format)
    
    ws_job.write('A37', 'Gross Margin (%)', sub_header_format)
    ws_job.write_formula('B37', '=IF(B7>0, B36/B7, 0)', percent_format)
    
    ws_job.write('A38', 'Target Price (Based on Desired Margin)', sub_header_format)
    ws_job.write_formula('B38', '=B27/(1-B35)', currency_format)
    
    ws_job.write('A39', 'Are You Undercharging?', sub_header_format)
    ws_job.write_formula('B39', '=IF(B7<B38, "YES - Increase Price!", "NO - Good Margin")', workbook.add_format({'border': 1, 'bold': True}))
    
    ws_job.conditional_format('B39', {'type': 'cell', 'criteria': '==', 'value': '"YES - Increase Price!"', 'format': red_format})
    ws_job.conditional_format('B39', {'type': 'cell', 'criteria': '==', 'value': '"NO - Good Margin"', 'format': green_format})
    
    ws_job.write('A41', 'Good/Better/Best Pricing Generator', header_format)
    ws_job.write('B41', 'Suggested Price', header_format)
    ws_job.write('A42', 'Good (Basic Wash - Target Margin)', sub_header_format)
    ws_job.write_formula('B42', '=B38', currency_format)
    ws_job.write('A43', 'Better (+20% Upsell e.g. Wax/Gutter)', sub_header_format)
    ws_job.write_formula('B43', '=B38*1.2', currency_format)
    ws_job.write('A44', 'Best (+40% Premium e.g. Full Property)', sub_header_format)
    ws_job.write_formula('B44', '=B38*1.4', currency_format)

    # --- SHEET: Callback Impact ---
    ws_call = workbook.add_worksheet('Callback Impact')
    ws_call.set_column('A:A', 45)
    ws_call.set_column('B:B', 15)
    
    ws_call.write('A1', 'The True Cost of a Callback', header_format)
    ws_call.write('B1', 'Amount', header_format)
    
    ws_call.write('A2', 'Original Job Profit', sub_header_format)
    ws_call.write_formula('B2', '=\'Job Calculator\'!B36', currency_format)
    
    ws_call.write('A3', 'Estimated Callback Hours', sub_header_format)
    ws_call.write('B3', 2.0, input_format)
    
    ws_call.write('A4', 'Callback Labor Cost', sub_header_format)
    ws_call.write_formula('B4', '=B3*\'Settings & Costs\'!B2*\'Settings & Costs\'!B3', currency_format)
    
    ws_call.write('A5', 'Callback Material/Chem Cost', sub_header_format)
    ws_call.write('B5', 15.00, input_currency)
    
    ws_call.write('A6', 'Callback Fuel Cost', sub_header_format)
    ws_call.write_formula('B6', '=\'Job Calculator\'!B24', currency_format)
    
    ws_call.write('A7', 'Total Callback Cost', header_format)
    ws_call.write_formula('B7', '=SUM(B4:B6)', currency_format)
    
    ws_call.write('A8', 'Net Profit After Callback', header_format)
    ws_call.write_formula('B8', '=B2-B7', currency_format)
    
    ws_call.write('A10', 'Jobs Needed to Recover', header_format)
    ws_call.write('B10', 'Count', header_format)
    ws_call.write('A11', 'How many NEW jobs at average profit to pay for this mistake?', sub_header_format)
    ws_call.write_formula('B11', '=IF(B2>0, B7/B2, 0)', number_format)

    # --- SHEET: Break Even Calculator ---
    ws_be = workbook.add_worksheet('Break Even Calculator')
    ws_be.set_column('A:A', 40)
    ws_be.set_column('B:B', 15)
    
    ws_be.write('A1', 'Break Even Calculator', header_format)
    ws_be.write('B1', 'Amount', header_format)
    
    ws_be.write('A2', 'Fixed Monthly Costs', header_format)
    ws_be.write('A3', 'Rent', sub_header_format)
    ws_be.write('B3', 500.00, input_currency)
    ws_be.write('A4', 'Insurance', sub_header_format)
    ws_be.write('B4', 200.00, input_currency)
    ws_be.write('A5', 'Vehicle Payments', sub_header_format)
    ws_be.write('B5', 400.00, input_currency)
    ws_be.write('A6', 'Software/Marketing', sub_header_format)
    ws_be.write('B6', 150.00, input_currency)
    ws_be.write('A7', 'Other Fixed Costs', sub_header_format)
    ws_be.write('B7', 100.00, input_currency)
    ws_be.write('A8', 'Total Fixed Monthly Costs', header_format)
    ws_be.write_formula('B8', '=SUM(B3:B7)', currency_format)
    
    ws_be.write('A10', 'Variable Costs Per Average Job', header_format)
    ws_be.write('A11', 'Average Labor Cost', sub_header_format)
    ws_be.write('B11', 100.00, input_currency)
    ws_be.write('A12', 'Average Chemical Cost', sub_header_format)
    ws_be.write('B12', 20.00, input_currency)
    ws_be.write('A13', 'Average Fuel/Other Direct Costs', sub_header_format)
    ws_be.write('B13', 15.00, input_currency)
    ws_be.write('A14', 'Total Variable Cost Per Job', header_format)
    ws_be.write_formula('B14', '=SUM(B11:B13)', currency_format)
    
    ws_be.write('A16', 'Average Revenue Per Job', sub_header_format)
    ws_be.write('B16', 400.00, input_currency)
    
    ws_be.write('A17', 'Contribution Margin Per Job', sub_header_format)
    ws_be.write_formula('B17', '=B16-B14', currency_format)
    
    ws_be.write('A19', 'Break Even Point', header_format)
    ws_be.write('A20', 'Jobs Required Per Month to Break Even', sub_header_format)
    ws_be.write_formula('B20', '=IF(B17>0, B8/B17, 0)', number_format)

    # --- SHEET: Job Logs ---
    ws_logs = workbook.add_worksheet('Job Logs')
    ws_logs.set_column('A:A', 15)
    ws_logs.set_column('B:B', 25)
    ws_logs.set_column('C:H', 15)
    
    # Summary at top
    ws_logs.write('A1', 'Total Revenue', header_format)
    ws_logs.write_formula('B1', '=SUM(C4:C100)', currency_format)
    ws_logs.write('C1', 'Total Gross Profit', header_format)
    ws_logs.write_formula('D1', '=SUM(H4:H100)', currency_format)
    
    log_headers = ['Date', 'Job Name', 'Revenue', 'Labor Cost', 'Chem Cost', 'Other Cost', 'Total Cost', 'Gross Profit']
    for col, header in enumerate(log_headers):
        ws_logs.write(2, col, header, header_format)
        
    # Add a sample row
    ws_logs.write(3, 0, '2024-01-01', input_format)
    ws_logs.write(3, 1, 'Sample Job 1', input_format)
    ws_logs.write(3, 2, 500.00, input_currency)
    ws_logs.write(3, 3, 150.00, input_currency)
    ws_logs.write(3, 4, 25.00, input_currency)
    ws_logs.write(3, 5, 15.00, input_currency)
    ws_logs.write_formula(3, 6, '=SUM(D4:F4)', currency_format)
    ws_logs.write_formula(3, 7, '=C4-G4', currency_format)
    
    # Add empty rows for user input
    for row in range(4, 100):
        ws_logs.write_blank(row, 0, '', input_format)
        ws_logs.write_blank(row, 1, '', input_format)
        ws_logs.write_blank(row, 2, '', input_currency)
        ws_logs.write_blank(row, 3, '', input_currency)
        ws_logs.write_blank(row, 4, '', input_currency)
        ws_logs.write_blank(row, 5, '', input_currency)
        ws_logs.write_formula(row, 6, f'=IF(C{row+1}="","",SUM(D{row+1}:F{row+1}))', currency_format)
        ws_logs.write_formula(row, 7, f'=IF(C{row+1}="","",C{row+1}-G{row+1})', currency_format)

    # --- SHEET: KPI Tracking ---
    ws_kpi = workbook.add_worksheet('KPI Tracking')
    ws_kpi.set_column('A:A', 40)
    ws_kpi.set_column('B:B', 15)
    
    ws_kpi.write('A1', 'KPI Tracking Dashboard', header_format)
    ws_kpi.write('B1', 'Value', header_format)
    
    ws_kpi.write('A3', 'Revenue Per Cleaner Per Day', header_format)
    ws_kpi.write('A4', 'Total Revenue (Period)', sub_header_format)
    ws_kpi.write('B4', 5000.00, input_currency)
    ws_kpi.write('A5', 'Total Cleaner Days Worked', sub_header_format)
    ws_kpi.write('B5', 10, input_format)
    ws_kpi.write('A6', 'Rev / Cleaner / Day', sub_header_format)
    ws_kpi.write_formula('B6', '=IF(B5>0, B4/B5, 0)', currency_format)
    
    ws_kpi.write('A8', 'Client Retention Rate', header_format)
    ws_kpi.write('A9', 'Total Clients at Start of Period', sub_header_format)
    ws_kpi.write('B9', 100, input_format)
    ws_kpi.write('A10', 'New Clients Acquired', sub_header_format)
    ws_kpi.write('B10', 20, input_format)
    ws_kpi.write('A11', 'Total Clients at End of Period', sub_header_format)
    ws_kpi.write('B11', 110, input_format)
    ws_kpi.write('A12', 'Retention Rate', sub_header_format)
    ws_kpi.write_formula('B12', '=IF(B9>0, (B11-B10)/B9, 0)', percent_format)
    
    ws_kpi.write('A14', 'Billable Hours Utilization', header_format)
    ws_kpi.write('A15', 'Total Hours Paid to Employees', sub_header_format)
    ws_kpi.write('B15', 400, input_format)
    ws_kpi.write('A16', 'Total Billable Hours Worked on Jobs', sub_header_format)
    ws_kpi.write('B16', 320, input_format)
    ws_kpi.write('A17', 'Utilization Rate', sub_header_format)
    ws_kpi.write_formula('B17', '=IF(B15>0, B16/B15, 0)', percent_format)
    
    ws_kpi.write('A19', 'Average Job Value', header_format)
    ws_kpi.write('A20', 'Total Revenue (From Logs)', sub_header_format)
    ws_kpi.write_formula('B20', '=\'Job Logs\'!B1', currency_format)
    ws_kpi.write('A21', 'Total Number of Jobs', sub_header_format)
    ws_kpi.write_formula('B21', '=COUNT(\'Job Logs\'!C4:C100)', number_format)
    ws_kpi.write('A22', 'Average Job Value', sub_header_format)
    ws_kpi.write_formula('B22', '=IF(B21>0, B20/B21, 0)', currency_format)

    # --- SHEET: Dashboard ---
    ws_dash = workbook.add_worksheet('Dashboard')
    ws_dash.set_column('A:A', 35)
    ws_dash.set_column('B:B', 20)
    
    ws_dash.write('A1', 'BUSINESS HEALTH DASHBOARD', header_format)
    ws_dash.write('B1', 'Status', header_format)
    
    ws_dash.write('A3', 'Current Job Margin (From Calculator)', sub_header_format)
    ws_dash.write_formula('B3', '=\'Job Calculator\'!B37', percent_format)
    
    ws_dash.write('A4', 'Target Margin', sub_header_format)
    ws_dash.write_formula('B4', '=\'Settings & Costs\'!B7', percent_format)
    
    ws_dash.write('A5', 'Pricing Status', sub_header_format)
    ws_dash.write_formula('B5', '=\'Job Calculator\'!B39', workbook.add_format({'border': 1, 'bold': True}))
    ws_dash.conditional_format('B5', {'type': 'cell', 'criteria': '==', 'value': '"YES - Increase Price!"', 'format': red_format})
    ws_dash.conditional_format('B5', {'type': 'cell', 'criteria': '==', 'value': '"NO - Good Margin"', 'format': green_format})
    
    ws_dash.write('A7', 'Break-Even Analysis (Based on Job Logs)', header_format)
    ws_dash.write('B7', 'Metrics', header_format)
    ws_dash.write('A8', 'Monthly Fixed Overhead', sub_header_format)
    ws_dash.write_formula('B8', '=\'Settings & Costs\'!B8', currency_format)
    
    ws_dash.write('A9', 'True Average Profit per Job', sub_header_format)
    ws_dash.write_formula('B9', '=IFERROR(AVERAGE(\'Job Logs\'!H4:H100), \'Job Calculator\'!B36)', currency_format)
    
    ws_dash.write('A10', 'Jobs Needed per Month to Break Even', sub_header_format)
    ws_dash.write_formula('B10', '=IF(B9>0, B8/B9, 0)', number_format)
    
    ws_dash.write('A11', 'Jobs Needed per Week', sub_header_format)
    ws_dash.write_formula('B11', '=B10/4.33', number_format)

    workbook.close()

if __name__ == "__main__":
    create_calculator()
