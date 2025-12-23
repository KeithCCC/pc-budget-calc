import pandas as pd
from datetime import datetime, timedelta

def create_excel_template(filename="PC_Budget_Template.xlsx"):
    """
    Create an Excel template with 3 sheets for PC budget calculation.
    """
    
    # Sheet 1: Current PC Rentals
    current_data = {
        'PC Model': ['Dell Latitude 5430', 'HP EliteBook 840', 'Lenovo ThinkPad T14', 'Dell Latitude 5430', 'HP EliteBook 840', ''],
        'Department': ['IT', 'Finance', 'HR', 'Sales', 'Marketing', ''],
        'Quantity': [50, 30, 25, 35, 20, ''],
        'Monthly Cost per PC': [50.00, 55.00, 52.00, 50.00, 55.00, ''],
        'Total Monthly Cost': [2500.00, 1650.00, 1300.00, 1750.00, 1100.00, ''],
        'Contract Start Date': ['2024-01-15', '2024-03-01', '2024-06-01', '2024-04-10', '2024-07-15', ''],
        'Notes': ['Batch 1', 'Premium model', 'Development team', 'Batch 2 - same model', 'Additional units', '']
    }
    
    # Sheet 2: Returning PCs (Contract End)
    # NOTE: Same model can appear multiple times with different end dates
    returning_data = {
        'PC Model': ['Dell Latitude 5430', 'Dell Latitude 5430', 'HP EliteBook 840', 'Lenovo ThinkPad T14', 'HP EliteBook 840', ''],
        'Department': ['IT', 'IT', 'Finance', 'HR', 'Finance', ''],
        'Quantity': [20, 30, 10, 15, 8, ''],
        'Monthly Cost per PC': [50.00, 50.00, 55.00, 52.00, 55.00, ''],
        'Total Monthly Cost': [1000.00, 1500.00, 550.00, 780.00, 440.00, ''],
        'Contract End Date': ['2025-01-31', '2025-04-30', '2025-02-28', '2025-03-31', '2025-05-31', ''],
        'Return Status': ['Confirmed', 'Confirmed', 'Pending', 'Confirmed', 'Planned', ''],
        'Notes': ['Batch 1 ending', 'Batch 2 ending later', 'Upgrade to new model', 'Contract renewal', 'Phase out', '']
    }
    
    # Sheet 3: New PCs to be Rented
    # NOTE: You can also have same model ordered at different times
    new_data = {
        'PC Model': ['Lenovo ThinkPad T16', 'Dell Latitude 7440', 'Lenovo ThinkPad T16', 'Dell Latitude 7440', ''],
        'Department': ['Engineering', 'Operations', 'Engineering', 'Sales', ''],
        'Quantity': [20, 15, 25, 10, ''],
        'Monthly Cost per PC': [58.00, 60.00, 58.00, 60.00, ''],
        'Total Monthly Cost': [1160.00, 900.00, 1450.00, 600.00, ''],
        'Expected Start Date': ['2025-02-01', '2025-03-01', '2025-04-01', '2025-03-15', ''],
        'Priority': ['High', 'Medium', 'High', 'Medium', ''],
        'Notes': ['Phase 1 - New team', 'Replacement units', 'Phase 2 - Expansion', 'Replace returning PCs', '']
    }
    
    # Create DataFrames
    df_current = pd.DataFrame(current_data)
    df_returning = pd.DataFrame(returning_data)
    df_new = pd.DataFrame(new_data)
    
    # Write to Excel with multiple sheets
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_current.to_excel(writer, sheet_name='1. Current Rentals', index=False)
        df_returning.to_excel(writer, sheet_name='2. Returning PCs', index=False)
        df_new.to_excel(writer, sheet_name='3. New PCs', index=False)
        
        # Access the workbook and worksheets to format
        workbook = writer.book
        
        # Format each sheet
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
                
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Make header row bold
            for cell in worksheet[1]:
                cell.font = cell.font.copy(bold=True)
    
    print(f"✓ Template created: {filename}")
    print(f"\nTemplate includes:")
    print(f"  Sheet 1: Current PC Rentals ({len(df_current)-1} example rows)")
    print(f"  Sheet 2: Returning PCs ({len(df_returning)-1} example rows)")
    print(f"  Sheet 3: New PCs ({len(df_new)-1} example rows)")
    print(f"\n📌 IMPORTANT: Same PC model can appear multiple times!")
    print(f"   Example: 'Dell Latitude 5430' appears twice in Returning PCs:")
    print(f"   - 20 units ending Jan 2025")
    print(f"   - 30 units ending Apr 2025")
    print(f"\nReplace the example data with your actual PC rental information.")
    print(f"Then run: python pc_budget_calculator.py")
    
    return filename


if __name__ == "__main__":
    create_excel_template()
