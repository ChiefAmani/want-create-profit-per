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
    
    # --- SHEET 3: Settings & Global Costs ---
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
            
    ws_settings.write('A10', 'Chemical Costs', header_format)
    ws_settings.write('B10', 'Cost per Gal/Oz', header_format)
    chem_data = [
        ('SH (Sodium Hypochlorite) / Gal', 4.50),
        ('Surfactant / Oz', 0.50),
        ('Degreaser / Oz', 0.75)
    ]
    for row, (label, val) in enumerate(chem_data, start=10):
        ws_settings.write(row, 0, label, sub_header_format)
        ws_settings.write(row, 1, val, input_currency)
        
    # --- SHEET 2: Job Calculator ---
    ws_job = workbook.add_worksheet('Job Calculator')
    ws_job.set_column('A:A', 35)
    ws_job.set_column('B:B', 20)
    
    ws_job.write('A1', 'Job Details (Edit Green Cells)', header_format)
    ws_job.write('B1', 'Input/Output', header_format)
    
    job_inputs = [
        ('Job Name/Address', '123 Main St', input_format),
        ('Estimated Time on Site (Hours)', 3.0, input_format),
        ('Number of Workers', 2, input_format),
        ('Round Trip Drive Time (Mins)', 45, input_format),
        ('Round Trip Distance (Miles)', 20, input_format),
        ('SH Used (Gallons)', 5, input_format),
        ('Surfactant Used (Oz)', 10, input_format),
        ('Quoted Price to Customer', 450.00, input_currency)
    ]
    
    for row, (label, val, fmt) in enumerate(job_inputs, start=1):
        ws_job.write(row, 0, label, sub_header_format)
        ws_job.write(row, 1, val, fmt)
            
    ws_job.write('A11', 'Cost Breakdown', header_format)
    ws_job.write('B11', 'Amount', header_format)
    
    ws_job.write('A12', 'Total Labor Cost', sub_header_format)
    ws_job.write_formula('B12', '=(B3+(B5/60))*B4*\'Settings & Costs\'!B2*\'Settings & Costs\'!B3', currency_format)
    
    ws_job.write('A13', 'Fuel Cost', sub_header_format)
    ws_job.write_formula('B13', '=(B6/\'Settings & Costs\'!B5)*\'Settings & Costs\'!B4', currency_format)
    
    ws_job.write('A14', 'Chemical Cost', sub_header_format)
    ws_job.write_formula('B14', '=(B7*\'Settings & Costs\'!B11)+(B8*\'Settings & Costs\'!B12)', currency_format)
    
    ws_job.write('A15', 'Equipment Wear & Tear', sub_header_format)
    ws_job.write_formula('B15', '=\'Settings & Costs\'!B6', currency_format)
    
    ws_job.write('A16', 'TOTAL JOB COST', header_format)
    ws_job.write_formula('B16', '=SUM(B12:B15)', currency_format)
    
    ws_job.write('A18', 'Profitability Analysis', header_format)
    ws_job.write('B18', 'Metrics', header_format)
    
    ws_job.write('A19', 'Gross Profit ($)', sub_header_format)
    ws_job.write_formula('B19', '=B9-B16', currency_format)
    
    ws_job.write('A20', 'Gross Margin (%)', sub_header_format)
    ws_job.write_formula('B20', '=IF(B9>0, B19/B9, 0)', percent_format)
    
    ws_job.write('A21', 'Target Price (Based on Target Margin)', sub_header_format)
    ws_job.write_formula('B21', '=B16/(1-\'Settings & Costs\'!B7)', currency_format)
    
    ws_job.write('A22', 'Are You Undercharging?', sub_header_format)
    ws_job.write_formula('B22', '=IF(B9<B21, "YES - Increase Price!", "NO - Good Margin")', workbook.add_format({'border': 1, 'bold': True}))
    
    # Conditional formatting for "Are You Undercharging?"
    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    ws_job.conditional_format('B22', {'type': 'cell', 'criteria': '==', 'value': '"YES - Increase Price!"', 'format': red_format})
    ws_job.conditional_format('B22', {'type': 'cell', 'criteria': '==', 'value': '"NO - Good Margin"', 'format': green_format})
    
    ws_job.write('A24', 'Good/Better/Best Pricing Generator', header_format)
    ws_job.write('B24', 'Suggested Price', header_format)
    ws_job.write('A25', 'Good (Basic Wash - Target Margin)', sub_header_format)
    ws_job.write_formula('B25', '=B21', currency_format)
    ws_job.write('A26', 'Better (+20% Upsell e.g. Wax/Gutter)', sub_header_format)
    ws_job.write_formula('B26', '=B21*1.2', currency_format)
    ws_job.write('A27', 'Best (+40% Premium e.g. Full Property)', sub_header_format)
    ws_job.write_formula('B27', '=B21*1.4', currency_format)

    # --- SHEET 4: Callback Cost Impact ---
    ws_call = workbook.add_worksheet('Callback Impact')
    ws_call.set_column('A:A', 45)
    ws_call.set_column('B:B', 15)
    
    ws_call.write('A1', 'The True Cost of a Callback', header_format)
    ws_call.write('B1', 'Amount', header_format)
    
    ws_call.write('A2', 'Original Job Profit', sub_header_format)
    ws_call.write_formula('B2', '=\'Job Calculator\'!B19', currency_format)
    
    ws_call.write('A3', 'Callback Labor Cost (Assumes 50% of orig)', sub_header_format)
    ws_call.write_formula('B3', '=\'Job Calculator\'!B12*0.5', currency_format)
    
    ws_call.write('A4', 'Callback Fuel Cost (100% of orig)', sub_header_format)
    ws_call.write_formula('B4', '=\'Job Calculator\'!B13', currency_format)
    
    ws_call.write('A5', 'Callback Chem Cost (Assumes 25% of orig)', sub_header_format)
    ws_call.write_formula('B5', '=\'Job Calculator\'!B14*0.25', currency_format)
    
    ws_call.write('A6', 'Total Callback Cost', header_format)
    ws_call.write_formula('B6', '=SUM(B3:B5)', currency_format)
    
    ws_call.write('A7', 'Net Profit After Callback', header_format)
    ws_call.write_formula('B7', '=B2-B6', currency_format)
    
    ws_call.write('A9', 'Jobs Needed to Recover', header_format)
    ws_call.write('B9', 'Count', header_format)
    ws_call.write('A10', 'How many NEW jobs at average profit to pay for this mistake?', sub_header_format)
    ws_call.write_formula('B10', '=IF(B2>0, B6/B2, 0)', number_format)

    # --- SHEET 1: Dashboard ---
    ws_dash = workbook.add_worksheet('Dashboard')
    ws_dash.set_column('A:A', 35)
    ws_dash.set_column('B:B', 20)
    
    ws_dash.write('A1', 'BUSINESS HEALTH DASHBOARD', header_format)
    ws_dash.write('B1', 'Status', header_format)
    
    ws_dash.write('A3', 'Current Job Margin', sub_header_format)
    ws_dash.write_formula('B3', '=\'Job Calculator\'!B20', percent_format)
    
    ws_dash.write('A4', 'Target Margin', sub_header_format)
    ws_dash.write_formula('B4', '=\'Settings & Costs\'!B7', percent_format)
    
    ws_dash.write('A5', 'Pricing Status', sub_header_format)
    ws_dash.write_formula('B5', '=\'Job Calculator\'!B22', workbook.add_format({'border': 1, 'bold': True}))
    ws_dash.conditional_format('B5', {'type': 'cell', 'criteria': '==', 'value': '"YES - Increase Price!"', 'format': red_format})
    ws_dash.conditional_format('B5', {'type': 'cell', 'criteria': '==', 'value': '"NO - Good Margin"', 'format': green_format})
    
    ws_dash.write('A7', 'Break-Even Analysis', header_format)
    ws_dash.write('B7', 'Metrics', header_format)
    ws_dash.write('A8', 'Monthly Fixed Overhead', sub_header_format)
    ws_dash.write_formula('B8', '=\'Settings & Costs\'!B8', currency_format)
    
    ws_dash.write('A9', 'Average Profit per Job', sub_header_format)
    ws_dash.write_formula('B9', '=\'Job Calculator\'!B19', currency_format)
    
    ws_dash.write('A10', 'Jobs Needed per Month to Break Even', sub_header_format)
    ws_dash.write_formula('B10', '=IF(B9>0, B8/B9, 0)', number_format)
    
    ws_dash.write('A11', 'Jobs Needed per Week', sub_header_format)
    ws_dash.write_formula('B11', '=B10/4.33', number_format)

    workbook.close()

if __name__ == "__main__":
    create_calculator()
