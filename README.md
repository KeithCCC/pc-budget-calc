# PC Budget Calculator

Calculate monthly PC rental costs based on rental PCs and returning PCs.

## Features

- Reads data from Excel files with multiple sheets
- Calculates net PC count and monthly costs
- Shows cost changes and projections
- Handles various column naming conventions

## Setup

1. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Prepare your Excel file with 2 sheets:
   - **Sheet 1**: Rental PC (all rented PCs including existing and new)
   - **Sheet 2**: Returning PC (PCs being returned/contract end)

2. Run the calculator:
```bash
python pc_budget_calculator_monthly.py
```

## Excel File Structure

The script automatically handles Excel files with monthly columns (Jan, Feb, Mar, etc.).

### Example Structure

**Sheet 1 - Rental PC:**
| PC Model | Unit Price | Jan | Feb | Mar | ... | Dec | Total |
|----------|-----------|-----|-----|-----|-----|-----|-------|
| Dell Latitude | 2500 | 50 | 50 | 45 | ... | 60 | 650 |
| HP EliteBook | 1800 | 30 | 30 | 30 | ... | 30 | 360 |
| Lenovo ThinkPad | 1200 | 0 | 20 | 20 | ... | 20 | 220 |

**Sheet 2 - Returning PC:**
| PC Model | Unit Price | Jan | Feb | Mar | ... | Dec | Total |
|----------|-----------|-----|-----|-----|-----|-----|-------|
| Dell Latitude | 2500 | 0 | 0 | 10 | ... | 0 | 20 |
| HP EliteBook | 1800 | 0 | 0 | 0 | ... | 5 | 10 |

## Calculation Formula

```
Net Monthly Cost = Rental Cost - Returning Cost
Net PC Count = Rental PCs - Returning PCs
Monthly Cost = PC Quantity × Unit Price
```

## Output

The script generates a single Excel file: **PC_Budget_Results.xlsx** with multiple sheets:

**Sheet 1: Overall Summary** - Net costs by PC model (Rental - Returning)
**Sheet 2: Rental PC** - Detailed monthly costs for all rental PCs
**Sheet 3: Returning PC** - Detailed monthly costs for returning PCs  
**Sheet 4: Monthly Summary** - Month-by-month breakdown with quantities

The console also displays:
- Detailed monthly cost tables
- Month-by-month breakdown showing rental, returning, and net costs
- Annual summary with total and average monthly costs
- Cost trend analysis

## Template Creation

To create a new Excel template:
```bash
python create_monthly_template.py
```

This will generate `PC_Budget_Monthly_Template.xlsx` with the correct structure.
