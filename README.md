# Market Volume Forecasting System

## Overview

This system automates the collection, processing, and forecasting of market volume data across multiple states. It consists of two main scripts that work in sequence to produce accurate volume forecasts based on historical data and membership trends.

## Features

- **Data Collection & Aggregation**: Processes raw Excel files from multiple project directories
- **Intelligent Forecasting**: Uses specialized algorithms for different forecasting scenarios
- **Membership-Adjusted Predictions**: Accounts for membership changes when calculating forecasts
- **Project Year Transitions**: Special handling for new project years (August transitions)
- **Rich Visualizations**: Includes charts and color-coded Excel outputs
- **Comprehensive Reporting**: Detailed Excel reports with methodology explanations

## System Components

### 1. Chart_Actuals.py

The first script in the pipeline that processes raw data files:

- Reads Excel files from project directories (`Project2023`, `Project2024`, `Project2025`)
- Normalizes column names and aggregates data by state
- Combines with membership reference data
- Generates a consolidated Excel file with actual volume data
- Creates summary visualizations

### 2. Market_Volume_Forecast.py

The second script that builds upon the actuals to create forecasts:

- Takes the output from Chart_Actuals.py as input
- Identifies which months need forecasting
- Applies two different forecasting algorithms:
  - **Standard Algorithm**: `Current Value + (Last Year's Change * Membership Ratio)`
  - **New Project Year Algorithm**: Special calculation for August transitions
- Generates detailed Excel output with color-coding to distinguish actuals from forecasts
- Provides methodology explanations in the output file

## Directory Structure

```
/
├── Chart_Actuals.py            # First script (processes raw data)
├── Market_Volume_Forecast.py   # Second script (generates forecasts)
├── Project2023/                # Raw data folders organized by project year
│   └── *.xlsx                  # Monthly data files (format: YYYY-MM*.xlsx)
├── Project2024/
│   └── *.xlsx
├── Project2025/
│   └── *.xlsx
├── MEM_REF/
│   └── MP_MEM_REF.xlsx         # Membership reference data
└── output/                     # Generated output files
    ├── Market_Volume_Actuals.xlsx     # Output from first script
    ├── Market_Volume_Forecast.xlsx    # Output from second script
    └── forecast_calculation.log       # Logging information
```

## Required Data Format

### Input Excel Files

- Monthly data files in project directories must have:
  - Filename starting with YYYY-MM format
  - Columns for state, vendor, and chart count (exact names can vary)

### Membership Reference File

- Located at `MEM_REF/MP_MEM_REF.xlsx`
- Must contain a row with index "Total"
- Columns should be formatted as year-month (e.g., "2023M08")

## Usage Instructions

### Step 1: Process Actual Data

```bash
python Run_Actuals.py
```

This will:
- Process all Excel files in the project directories
- Generate `output/Market_Volume_Actuals.xlsx`
- Show a summary of processed files

### Step 2: Generate Forecasts

```bash
python Run_Forecast.py
```

This will:
- Read the actuals file generated in Step 1
- Calculate forecasts for future months
- Generate `output/Market_Volume_Forecast.xlsx`
- Display a summary of the forecast results

## Forecasting Methodology

### Standard Forecast (Most Months)

Formula: `Current Value + (Last Year's Change * Membership Ratio)`

Where:
- Current Value = Value from previous month
- Last Year's Change = Same month previous year minus its previous month
- Membership Ratio = Current membership / Previous year's membership

### Special Forecast (New Project Year - August)

Formula: `Previous Year's August * Membership Adjustment`

Where:
- Membership Adjustment = Estimated future membership / Previous August membership

## Dependencies

- Python 3.7+
- pandas
- numpy
- xlsxwriter
- matplotlib

## Installation

```bash
pip install pandas numpy xlsxwriter matplotlib
```

## Author

Your Name  
April 2025
