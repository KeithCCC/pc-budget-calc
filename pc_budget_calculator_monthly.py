import pandas as pd
import numpy as np
from datetime import datetime

def calculate_monthly_pc_budget(excel_file):
    """
    Calculate monthly PC rental costs with month-by-month breakdown.
    
    Expected Excel structure:
    Columns: PC Model | Unit Price | Jan | Feb | Mar | ... | Dec | Total (optional)
    
    - Sheet 1: Rental PC (monthly rental quantities)
    - Sheet 2: Returning PC (monthly return quantities)
    """
    
    def create_overall_summary(df_rental, df_returning, month_cols, 
                              price_col_rental, price_col_returning):
        """Create overall summary: Rental - Returning, matched by PC Model"""
        
        model_col = df_rental.columns[0] if len(df_rental) > 0 else None
        if not model_col:
            return pd.DataFrame()
        
        # Get all unique PC models from both sheets
        all_models = set()
        if len(df_rental) > 0:
            all_models.update(df_rental[model_col].dropna().tolist())
        if len(df_returning) > 0:
            returning_model_col = df_returning.columns[0]
            all_models.update(df_returning[returning_model_col].dropna().tolist())
        
        all_models = sorted(list(all_models))
        
        # Create result dataframe
        result_data = {
            'PC Model': [],
            'Unit Price': []
        }
        
        for month in month_cols:
            result_data[str(month)] = []
        
        # Process each PC model
        for model in all_models:
            # Get data for this model from each sheet
            rental_row = df_rental[df_rental[model_col] == model] if len(df_rental) > 0 else pd.DataFrame()
            
            returning_model_col = df_returning.columns[0] if len(df_returning) > 0 else None
            returning_row = df_returning[df_returning[returning_model_col] == model] if returning_model_col else pd.DataFrame()
            
            # Get unit price (prefer rental, then returning)
            if len(rental_row) > 0:
                unit_price = rental_row.iloc[0][price_col_rental]
            elif len(returning_row) > 0 and price_col_returning:
                unit_price = returning_row.iloc[0][price_col_returning]
            else:
                unit_price = 0
            
            result_data['PC Model'].append(model)
            result_data['Unit Price'].append(unit_price)
            
            # Calculate net cost for each month
            for month in month_cols:
                rental_cost = 0
                returning_cost = 0
                
                if len(rental_row) > 0 and month in rental_row.columns:
                    rental_qty = rental_row.iloc[0][month] if pd.notna(rental_row.iloc[0][month]) else 0
                    rental_cost = rental_qty * unit_price
                
                if len(returning_row) > 0 and month in returning_row.columns:
                    returning_qty = returning_row.iloc[0][month] if pd.notna(returning_row.iloc[0][month]) else 0
                    returning_cost = returning_qty * unit_price
                
                net_cost = rental_cost - returning_cost
                result_data[str(month)].append(net_cost)
        
        # Add TOTAL row
        result_data['PC Model'].append('TOTAL')
        result_data['Unit Price'].append('')
        
        for month in month_cols:
            month_total = sum([result_data[str(month)][i] for i in range(len(all_models))])
            result_data[str(month)].append(month_total)
        
        return pd.DataFrame(result_data)
    
    def create_monthly_cost_csv(df, month_cols, price_col):
        """Create a DataFrame for CSV export with monthly costs per PC model"""
        if len(df) == 0 or not price_col:
            return pd.DataFrame()
        
        # Get PC model column (usually first column)
        model_col = df.columns[0]
        
        # Create result dataframe
        result_data = {
            'PC Model': [],
            'Unit Price': []
        }
        
        # Add month columns
        for month in month_cols:
            result_data[str(month)] = []
        
        # Add each PC model row
        for idx, row in df.iterrows():
            result_data['PC Model'].append(row[model_col])
            result_data['Unit Price'].append(row[price_col])
            
            for month in month_cols:
                qty = row[month] if month in row else 0
                monthly_cost = qty * row[price_col]
                result_data[str(month)].append(monthly_cost)
        
        # Add TOTAL row
        result_data['PC Model'].append('TOTAL')
        result_data['Unit Price'].append('')
        
        for month in month_cols:
            month_total = (df[month] * df[price_col]).sum()
            result_data[str(month)].append(month_total)
        
        return pd.DataFrame(result_data)
    
    def display_summary_table(summary_df, month_cols):
        """Display the overall summary table"""
        if len(summary_df) == 0:
            print("  (No data)")
            return
        
        # Header
        header = f"{'PC Model':<40} {'Unit Price':>12}"
        for month in month_cols:
            header += f" {str(month):>12}"
        print(header)
        print("-" * len(header))
        
        # Each row
        for idx, row in summary_df.iterrows():
            model_name = str(row['PC Model'])[:38]
            unit_price = row['Unit Price'] if row['Unit Price'] != '' else ''
            
            if unit_price != '':
                line = f"{model_name:<40} {unit_price:>12,.0f}"
            else:
                line = f"{model_name:<40} {'':<12}"
            
            for month in month_cols:
                cost = row[str(month)]
                line += f" {cost:>12,.0f}"
            
            print(line)
    
    def display_monthly_cost_table(df, month_cols, price_col):
        """Display a table showing monthly costs per PC model"""
        if len(df) == 0:
            print("  (No data)")
            return
        
        # Get PC model column (usually first column)
        model_col = df.columns[0]
        
        # Header
        header = f"{'PC Model':<40} {'Unit Price':>12}"
        for month in month_cols:
            header += f" {str(month):>12}"
        print(header)
        print("-" * len(header))
        
        # Each row
        for idx, row in df.iterrows():
            model_name = str(row[model_col])[:38]
            unit_price = row[price_col]
            
            line = f"{model_name:<40} {unit_price:>12,.0f}"
            
            for month in month_cols:
                qty = row[month] if month in row else 0
                monthly_cost = qty * unit_price
                line += f" {monthly_cost:>12,.0f}"
            
            print(line)
        
        # Totals row
        total_line = f"{'TOTAL':<40} {'':<12}"
        for month in month_cols:
            month_total = (df[month] * df[price_col]).sum()
            total_line += f" {month_total:>12,.0f}"
        print("-" * len(header))
        print(total_line)
    
    print(f"\n{'='*80}")
    print(f"PC MONTHLY BUDGET CALCULATOR - {datetime.now().strftime('%Y-%m-%d')}")
    print(f"{'='*80}\n")
    
    # Read all sheets
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    sheet_names = list(excel_data.keys())
    
    if len(sheet_names) < 2:
        print("Error: Excel file must contain at least 2 sheets")
        return
    
    # Get the two tables
    df_rental = excel_data[sheet_names[0]].copy()
    df_returning = excel_data[sheet_names[1]].copy()
    
    # Identify month columns (Jan, Feb, Mar, ... or 1月, 2月, 3月, ...)
    month_patterns = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
                     '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']
    
    def get_month_columns(df):
        """Extract month column names from dataframe"""
        month_cols = []
        for col in df.columns:
            col_str = str(col).strip()
            if any(month in col_str for month in month_patterns):
                month_cols.append(col)
        return month_cols
    
    month_cols_rental = get_month_columns(df_rental)
    month_cols_returning = get_month_columns(df_returning)
    
    if not month_cols_rental:
        print("Error: No month columns found. Expected columns like: Jan, Feb, Mar, etc.")
        return
    
    # Get unit price column
    def get_price_col(df):
        for col in df.columns:
            if 'price' in str(col).lower() or 'unit' in str(col).lower():
                return col
        return None
    
    price_col_rental = get_price_col(df_rental)
    price_col_returning = get_price_col(df_returning)
    
    if not price_col_rental:
        print("Error: No 'Unit Price' column found")
        return
    
    print(f"Found {len(month_cols_rental)} month columns: {month_cols_rental[:3]}...{month_cols_rental[-1]}\n")
    
    # Clean data - remove empty rows
    df_rental = df_rental[df_rental[price_col_rental].notna()].reset_index(drop=True)
    df_returning = df_returning[df_returning[price_col_returning].notna()].reset_index(drop=True) if price_col_returning else pd.DataFrame()
    
    # Initialize monthly totals
    monthly_costs_rental = []
    monthly_costs_returning = []
    monthly_qty_rental = []
    monthly_qty_returning = []
    
    # Calculate monthly costs
    for month_col in month_cols_rental:
        # Rental PCs
        df_rental[month_col] = pd.to_numeric(df_rental[month_col], errors='coerce').fillna(0)
        cost = (df_rental[month_col] * df_rental[price_col_rental]).sum()
        qty = df_rental[month_col].sum()
        monthly_costs_rental.append(cost)
        monthly_qty_rental.append(qty)
    
    for month_col in month_cols_returning:
        # Returning PCs
        if len(df_returning) > 0 and price_col_returning:
            df_returning[month_col] = pd.to_numeric(df_returning[month_col], errors='coerce').fillna(0)
            cost = (df_returning[month_col] * df_returning[price_col_returning]).sum()
            qty = df_returning[month_col].sum()
        else:
            cost = 0
            qty = 0
        monthly_costs_returning.append(cost)
        monthly_qty_returning.append(qty)
    
    # Display summary
    print(f"{'='*80}")
    print(f"SHEET SUMMARY")
    print(f"{'='*80}")
    print(f"Sheet 1 - Rental PC:     {len(df_rental)} PC models")
    print(f"Sheet 2 - Returning PC:  {len(df_returning)} PC models")
    
    # Create CSV exports
    csv_rental = create_monthly_cost_csv(df_rental, month_cols_rental, price_col_rental)
    csv_returning = create_monthly_cost_csv(df_returning, month_cols_returning, price_col_returning)
    
    # Create overall summary by matching PC models
    csv_summary = create_overall_summary(df_rental, df_returning, 
                                         month_cols_rental, price_col_rental, 
                                         price_col_returning)
    
    # Save to Excel file with multiple sheets
    output_filename = 'PC_Budget_Results.xlsx'
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        csv_summary.to_excel(writer, sheet_name='Overall Summary', index=False)
        csv_rental.to_excel(writer, sheet_name='Rental PC', index=False)
        if len(csv_returning) > 0:
            csv_returning.to_excel(writer, sheet_name='Returning PC', index=False)
        
        # Format sheets
        workbook = writer.book
        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                adjusted_width = min(max_length + 2, 40)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Make header row bold
            for cell in worksheet[1]:
                cell.font = cell.font.copy(bold=True)
    
    print(f"\n✓ Excel file created: {output_filename}")
    print(f"  - Sheet 1: Overall Summary (Rental - Returning by PC Model)")
    print(f"  - Sheet 2: Rental PC (Monthly Costs)")
    if len(csv_returning) > 0:
        print(f"  - Sheet 3: Returning PC (Monthly Costs)")
    
    # Display overall summary table
    print(f"\n{'='*80}")
    print(f"OVERALL SUMMARY: NET MONTHLY COSTS BY PC MODEL")
    print(f"(Rental - Returning)")
    print(f"{'='*80}")
    display_summary_table(csv_summary, month_cols_rental)
    
    # Create detailed monthly cost tables
    print(f"\n{'='*80}")
    print(f"TABLE 1: RENTAL PC - MONTHLY COSTS")
    print(f"{'='*80}")
    display_monthly_cost_table(df_rental, month_cols_rental, price_col_rental)
    
    print(f"\n{'='*80}")
    print(f"TABLE 2: RETURNING PC - MONTHLY COSTS")
    print(f"{'='*80}")
    if len(df_returning) > 0 and price_col_returning:
        display_monthly_cost_table(df_returning, month_cols_returning, price_col_returning)
    else:
        print("  (No data)")
    
    # Monthly breakdown
    print(f"\n{'='*80}")
    print(f"MONTHLY BREAKDOWN")
    print(f"{'='*80}")
    print(f"{'Month':<10} {'Rental':<15} {'Returning':<15} {'Net Cost':<15} {'Net Qty':<10}")
    print(f"{'-'*80}")
    
    net_monthly_costs = []
    net_monthly_qty = []
    
    for i, month in enumerate(month_cols_rental):
        rental_cost = monthly_costs_rental[i]
        returning_cost = monthly_costs_returning[i] if i < len(monthly_costs_returning) else 0
        
        rental_qty = monthly_qty_rental[i]
        returning_qty = monthly_qty_returning[i] if i < len(monthly_qty_returning) else 0
        
        net_cost = rental_cost - returning_cost
        net_qty = rental_qty - returning_qty
        
        net_monthly_costs.append(net_cost)
        net_monthly_qty.append(net_qty)
        
        print(f"{str(month):<10} ${rental_cost:>12,.0f} -${returning_cost:>12,.0f} = ${net_cost:>12,.0f}  {int(net_qty):>6} PCs")
    
    # Annual summary
    print(f"{'='*80}")
    print(f"ANNUAL SUMMARY")
    print(f"{'='*80}")
    
    total_rental = sum(monthly_costs_rental)
    total_returning = sum(monthly_costs_returning)
    total_net = sum(net_monthly_costs)
    
    avg_monthly_cost = total_net / len(net_monthly_costs) if net_monthly_costs else 0
    avg_monthly_qty = sum(net_monthly_qty) / len(net_monthly_qty) if net_monthly_qty else 0
    
    print(f"Total Rental PC Costs:      ${total_rental:>15,.2f}")
    print(f"Total Returning PC Savings: ${total_returning:>15,.2f}")
    print(f"{'-'*80}")
    print(f"TOTAL ANNUAL COST:          ${total_net:>15,.2f}")
    print(f"Average Monthly Cost:       ${avg_monthly_cost:>15,.2f}")
    print(f"Average Monthly PCs:        {avg_monthly_qty:>15,.0f}")
    print(f"{'='*80}")
    
    # Export summary to Excel file
    summary_data = {
        'Month': month_cols_rental,
        'Rental PC Cost': monthly_costs_rental,
        'Returning PC Cost': monthly_costs_returning[:len(month_cols_rental)],
        'Net Monthly Cost': net_monthly_costs,
        'Net PC Quantity': net_monthly_qty
    }
    summary_df = pd.DataFrame(summary_data)
    
    # Append to existing Excel file
    with pd.ExcelWriter(output_filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        summary_df.to_excel(writer, sheet_name='Monthly Summary', index=False)
        
        # Format the new sheet
        workbook = writer.book
        worksheet = workbook['Monthly Summary']
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 40)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Make header row bold
        for cell in worksheet[1]:
            cell.font = cell.font.copy(bold=True)
    
    print(f"  - Sheet 4: Monthly Summary (Monthly Breakdown)")
    
    # Cost change analysis
    if len(net_monthly_costs) >= 2:
        first_month_cost = net_monthly_costs[0]
        last_month_cost = net_monthly_costs[-1]
        cost_change = last_month_cost - first_month_cost
        change_pct = (cost_change / first_month_cost * 100) if first_month_cost > 0 else 0
        
        print(f"\nCost Trend (First vs Last Month):")
        print(f"  First month: ${first_month_cost:,.2f}")
        print(f"  Last month:  ${last_month_cost:,.2f}")
        print(f"  Change:      ${cost_change:+,.2f} ({change_pct:+.1f}%)")
    
    return {
        'monthly_costs': net_monthly_costs,
        'monthly_quantities': net_monthly_qty,
        'total_annual': total_net,
        'average_monthly': avg_monthly_cost
    }


if __name__ == "__main__":
    import os
    import sys
    
    # Check if file specified as argument - THIS TAKES PRIORITY
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
        if not os.path.exists(excel_file):
            print(f"Error: File '{excel_file}' not found")
            sys.exit(1)
        print(f"Processing specified file: {excel_file}\n")
        
        try:
            result = calculate_monthly_pc_budget(excel_file)
            print("\n✓ Calculation complete!")
            
        except Exception as e:
            print(f"\nError: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
    else:
        # Find Excel files, prioritize template files
        excel_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
        
        if not excel_files:
            print("No Excel files found in current directory.")
            print("\nUsage: python pc_budget_calculator_monthly.py [filename.xlsx]")
            print("   OR: python pc_budget_calculator_monthly.py")
            print("       (will auto-detect Excel files, preferring templates)")
            print("\nCreate template first: python create_monthly_template.py")
            sys.exit(1)
        
        # Prefer template files when no file specified
        template_files = [f for f in excel_files if 'template' in f.lower() or 'monthly' in f.lower()]
        
        print(f"Found Excel files: {excel_files}\n")
        
        if template_files:
            excel_file = template_files[0]
            print(f"Auto-selected template file: {excel_file}\n")
        else:
            excel_file = excel_files[0]
            print(f"Auto-selected file: {excel_file}\n")
        
        try:
            result = calculate_monthly_pc_budget(excel_file)
            print("\n✓ Calculation complete!")
            
        except Exception as e:
            print(f"\nError: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
