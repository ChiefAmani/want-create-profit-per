# TECHNICAL SPECIFICATION: Residential Cleaning Profitability Calculator

## 1. OVERVIEW
This project generates a highly marketable, premium Excel-based "Profit Per Job Calculator" tailored for residential cleaning businesses (window, gutter, siding, pressure washing). The product is delivered as an `.xlsx` file generated via a Python script using `xlsxwriter`.

## 2. TECH STACK
- **Language**: Python 3.10+
- **Library**: `xlsxwriter` (Chosen for its robust support for Excel formulas, formatting, data validation, and charts without requiring Excel to be installed during generation).

## 3. EXCEL WORKBOOK ARCHITECTURE
The generated Excel file (`Cleaning_Profit_Calculator.xlsx`) will contain the following worksheets:

### Sheet 1: Dashboard (The "Sell" Sheet)
- **Purpose**: Visual summary and high-level business health.
- **Features**:
  - **"Are You Undercharging?" Indicator**: Conditional formatting (Red/Yellow/Green) based on target gross margin vs. actual.
  - **Break-Even Analysis**: Visual chart showing jobs needed per week to cover fixed costs.
  - **Good/Better/Best Pricing Generator**: Auto-calculates 3-tier pricing based on base job costs.

### Sheet 2: Job Calculator (The Engine)
- **Purpose**: Per-job input and profitability calculation.
- **Inputs**:
  - Job Type (Dropdown: Window, Gutter, Siding, Pressure Washing)
  - Estimated Labor Hours
  - Drive Time (Minutes) & Distance (Miles)
  - Chemical Usage (Ounces/Gallons)
  - Upsell Add-ons
- **Outputs**:
  - Total Cost (Labor + Chemicals + Fuel + Wear & Tear)
  - Suggested Minimum Price
  - Actual Profit Margin

### Sheet 3: Settings & Global Costs
- **Purpose**: User customization (makes the sheet reusable for any business).
- **Inputs**:
  - Labor Rates (Base pay + taxes/insurance multiplier)
  - Chemical Costs (Cost per gallon/oz for SH, Surfactants, etc.)
  - Vehicle Costs (MPG, Gas Price)
  - Fixed Monthly Overhead (Insurance, Marketing, Software)

### Sheet 4: Callback Cost Impact
- **Purpose**: Demonstrates the hidden cost of mistakes.
- **Features**:
  - Calculates the true cost of a callback (Lost time + extra chemicals + fuel).
  - Shows how many *new* jobs are required to offset one callback.

## 4. STRATEGIC VALUE ADDITIONS (To "Sell Like Hot Cakes")
1. **Drive Time & Fuel Cost Analyzer**: Often ignored by owner-operators. Highlighting this cost increases the perceived value of the tool.
2. **Good/Better/Best Pricing Generator**: Transitions the tool from a simple "cost tracker" to a "revenue generator" by encouraging upsells.
3. **Equipment Wear & Tear Micro-Costs**: Adds a flat fee per job for pump/hose degradation, showing advanced financial maturity.
4. **Callback Offset Formula**: A psychological feature that trains owners to do it right the first time by showing the devastating margin impact of a return trip.

## 5. FILE TREE
.
|-- generate_calculator.py (Python script to build the Excel file)
|-- requirements.txt
|-- README.md
|-- TECHNICAL_SPEC.md
