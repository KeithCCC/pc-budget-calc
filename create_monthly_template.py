import pandas as pd

def create_monthly_template(filename="PC_Budget_Monthly_Template.xlsx"):
    """
    Create an Excel template with monthly quantities for each PC model.
    Structure matches the user's format with monthly columns.
    """
    
    # Define months
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    # Sheet 1: Rental PC (Monthly)
    rental_base = {
        'PC Model': [
            'HP Elite SFF 800 G9 PC (4G087AV)',
            'LIFE BOOK U9313', 
            'SurfacePro8',
            'ThinkPad L13Gen2',
            'ThinkPad X13 Gen 4',
            'ThinkPad X13 Gen5',
            'Dell Latitude 7440',
            'Lenovo ThinkPad T16',
            ''
        ],
        'Unit Price': [3690, 4130, 6750, 4200, 4010, 3800, 4500, 5800, '']
    }
    
    # Add monthly columns with example data
    for month in months:
        if month in ['Jan', 'Feb', 'Mar']:
            rental_base[month] = [1, 80, 2, 18, 70, 0, 0, 0, '']
        elif month in ['Apr', 'May', 'Jun']:
            rental_base[month] = [1, 80, 2, 18, 70, 20, 15, 0, '']
        elif month in ['Jul', 'Aug', 'Sep']:
            rental_base[month] = [1, 80, 2, 18, 70, 50, 25, 25, '']
        else:
            rental_base[month] = [1, 80, 2, 18, 70, 50, 25, 25, '']
    
    # Add total column
    rental_base['Total'] = [12, 960, 24, 216, 840, 350, 200, 175, '']
    
    # Sheet 2: Returning PC (Monthly) - Shows when PCs are returned each month
    returning_base = {
        'PC Model': [
            'HP Elite SFF 800 G9 PC (4G087AV)',
            'LIFE BOOK U9313',
            'ThinkPad L13Gen2',
            'ThinkPad X13 Gen 4',
            ''
        ],
        'Unit Price': [3690, 4130, 4200, 4010, '']
    }
    
    # Monthly return quantities (example: some PCs return in specific months)
    for i, month in enumerate(months):
        if month == 'Mar':  # Some PCs return in March
            returning_base[month] = [0, 20, 5, 10, '']
        elif month == 'Jun':  # More PCs return in June
            returning_base[month] = [0, 15, 3, 15, '']
        elif month == 'Sep':  # More returns in September
            returning_base[month] = [1, 10, 2, 5, '']
        else:
            returning_base[month] = [0, 0, 0, 0, '']
    
    returning_base['Total'] = [1, 45, 10, 30, '']
    
    # Create DataFrames
    df_rental = pd.DataFrame(rental_base)
    df_returning = pd.DataFrame(returning_base)
    
    # Write to Excel
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_rental.to_excel(writer, sheet_name='1. Rental PC', index=False)
        df_returning.to_excel(writer, sheet_name='2. Returning PC', index=False)
        
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
                
                adjusted_width = min(max_length + 2, 40)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Make header row bold
            for cell in worksheet[1]:
                cell.font = cell.font.copy(bold=True)
    
    print(f"✓ Monthly template created: {filename}")
    print(f"\nTemplate structure:")
    print(f"  Columns: PC Model | Unit Price | Jan | Feb | ... | Dec | Total")
    print(f"\n  Sheet 1: Rental PC - Monthly rental quantities (includes all PCs)")
    print(f"  Sheet 2: Returning PC - Monthly return quantities")
    print(f"\n📌 How to use:")
    print(f"   - Enter the quantity of each PC model for each month")
    print(f"   - Unit Price is the monthly rental cost per PC")
    print(f"   - Each row can have different quantities per month")
    print(f"   - Calculator will compute: (Rental - Returning) × Unit Price for each month")
    print(f"\nRun: python pc_budget_calculator_monthly.py")
    
    return filename


if __name__ == "__main__":
    create_monthly_template()
