import pandas as pd
from datetime import datetime
import os

def calculate_monthly_pc_budget(excel_file, cost_per_pc=None):
    """
    Calculate monthly PC rental costs.
    
    Parameters:
    - excel_file: Path to Excel file containing the data
    - cost_per_pc: Default cost per PC if not specified in tables
    
    Expected Excel structure:
    - Sheet 1 (or named "Current"): Current PC rentals
    - Sheet 2 (or named "Returning"): PCs being returned
    - Sheet 3 (or named "New"): New PCs to be rented
    """
    
    print(f"\n{'='*60}")
    print(f"PC Budget Calculator - {datetime.now().strftime('%Y-%m-%d')}")
    print(f"{'='*60}\n")
    
    # Read all sheets
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    sheet_names = list(excel_data.keys())
    
    print(f"Found {len(sheet_names)} sheets in Excel file:")
    for i, name in enumerate(sheet_names, 1):
        print(f"  Sheet {i}: {name}")
    print()
    
    # Get the three tables
    # Try to identify by sheet names or use first 3 sheets
    if len(sheet_names) >= 3:
        current_pcs = excel_data[sheet_names[0]]
        returning_pcs = excel_data[sheet_names[1]]
        new_pcs = excel_data[sheet_names[2]]
    else:
        print("Error: Excel file must contain at least 3 sheets")
        return
    
    print("="*60)
    print("TABLE 1: Current PC Rentals")
    print("="*60)
    print(current_pcs.head(10))
    print(f"\nTotal rows: {len(current_pcs)}")
    
    print("\n" + "="*60)
    print("TABLE 2: Returning PCs (Contract End)")
    print("="*60)
    print(returning_pcs.head(10))
    print(f"\nTotal rows: {len(returning_pcs)}")
    
    print("\n" + "="*60)
    print("TABLE 3: New PCs to be Rented")
    print("="*60)
    print(new_pcs.head(10))
    print(f"\nTotal rows: {len(new_pcs)}")
    
    # Calculate totals
    # Try to find quantity or count columns
    def get_quantity(df):
        """Extract quantity from dataframe, handling various column names"""
        possible_qty_cols = ['quantity', 'qty', 'count', 'number', 'total', 'pcs']
        
        for col in df.columns:
            if any(name in col.lower() for name in possible_qty_cols):
                return df[col].sum()
        
        # If no quantity column found, assume each row is 1 PC
        return len(df)
    
    def get_cost(df, default_cost=None):
        """Extract cost information from dataframe"""
        possible_cost_cols = ['cost', 'price', 'monthly', 'rental', 'fee']
        
        for col in df.columns:
            if any(name in col.lower() for name in possible_cost_cols):
                return df[col].sum()
        
        # If cost column not found, use default cost * quantity
        if default_cost:
            qty = get_quantity(df)
            return qty * default_cost
        
        return 0
    
    # Calculate quantities
    current_qty = get_quantity(current_pcs)
    returning_qty = get_quantity(returning_pcs)
    new_qty = get_quantity(new_pcs)
    
    print("\n" + "="*60)
    print("QUANTITY SUMMARY")
    print("="*60)
    print(f"Current PCs renting:        {current_qty:>10,.0f}")
    print(f"PCs returning to vendor:    {returning_qty:>10,.0f}")
    print(f"New PCs to be rented:       {new_qty:>10,.0f}")
    print(f"{'-'*60}")
    print(f"Net PC count:               {current_qty - returning_qty + new_qty:>10,.0f}")
    
    # Calculate costs
    current_cost = get_cost(current_pcs, cost_per_pc)
    returning_cost = get_cost(returning_pcs, cost_per_pc)
    new_cost = get_cost(new_pcs, cost_per_pc)
    
    print("\n" + "="*60)
    print("MONTHLY COST ANALYSIS")
    print("="*60)
    print(f"Current monthly cost:       ${current_cost:>15,.2f}")
    print(f"Returning PC cost savings:  ${returning_cost:>15,.2f}")
    print(f"New PC additional cost:     ${new_cost:>15,.2f}")
    print(f"{'-'*60}")
    total_monthly_cost = current_cost - returning_cost + new_cost
    print(f"NET MONTHLY COST:           ${total_monthly_cost:>15,.2f}")
    print("="*60)
    
    # Calculate change
    cost_change = total_monthly_cost - current_cost
    change_pct = (cost_change / current_cost * 100) if current_cost > 0 else 0
    
    print(f"\nCost change from current:   ${cost_change:>15,.2f} ({change_pct:+.1f}%)")
    print(f"Annual cost projection:     ${total_monthly_cost * 12:>15,.2f}")
    
    # Return summary dictionary
    return {
        'current_qty': current_qty,
        'returning_qty': returning_qty,
        'new_qty': new_qty,
        'net_qty': current_qty - returning_qty + new_qty,
        'current_cost': current_cost,
        'returning_cost': returning_cost,
        'new_cost': new_cost,
        'total_monthly_cost': total_monthly_cost,
        'cost_change': cost_change,
        'annual_cost': total_monthly_cost * 12
    }


def calculate_by_month(excel_file):
    """
    Calculate monthly budget if the data includes date/month columns.
    """
    print("\n" + "="*60)
    print("MONTHLY BREAKDOWN")
    print("="*60)
    
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    sheet_names = list(excel_data.keys())
    
    if len(sheet_names) < 2:
        print("Need at least 2 sheets for monthly analysis")
        return
    
    returning_pcs = excel_data[sheet_names[1]]
    new_pcs = excel_data[sheet_names[2]] if len(sheet_names) >= 3 else pd.DataFrame()
    
    # Try to find date columns
    date_cols = []
    for col in returning_pcs.columns:
        if 'date' in col.lower() or 'month' in col.lower() or 'return' in col.lower():
            date_cols.append(col)
    
    if date_cols:
        print(f"\nFound date column(s): {date_cols}")
        # Group by month if dates are available
        # This would require more specific implementation based on actual data structure
        print("\nFor detailed monthly breakdown, please review the data structure.")
    else:
        print("\nNo date columns found for monthly breakdown.")


if __name__ == "__main__":
    # Default configuration
    DEFAULT_COST_PER_PC = 50  # Adjust this default value as needed
    
    # Check for Excel files in current directory
    excel_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print("No Excel files found in current directory.")
        print("\nUsage: python pc_budget_calculator.py")
        print("Place your Excel file in the same directory with 3 sheets:")
        print("  Sheet 1: Current PC rentals")
        print("  Sheet 2: Returning PCs")
        print("  Sheet 3: New PCs")
    else:
        print(f"Found Excel files: {excel_files}\n")
        
        # Use first Excel file or prompt user
        excel_file = excel_files[0]
        print(f"Processing: {excel_file}\n")
        
        try:
            result = calculate_monthly_pc_budget(excel_file, DEFAULT_COST_PER_PC)
            
            # Optional: Calculate monthly breakdown if dates available
            # calculate_by_month(excel_file)
            
            print("\n" + "="*60)
            print("Calculation complete!")
            print("="*60)
            
        except Exception as e:
            print(f"\nError processing file: {e}")
            print("\nPlease ensure your Excel file has the correct structure:")
            print("  - Sheet 1: Current PC rentals")
            print("  - Sheet 2: Returning PCs")
            print("  - Sheet 3: New PCs")
            print("\nColumns should include quantity/count and cost/price information.")
