#!/usr/bin/env python3
"""
Market Volume Forecast Generator

This script calculates forecasted market volumes based on historical data.
It uses a specialized algorithm that accounts for membership growth and seasonal patterns.
This is the second script in the process, after Market_Volume_Actuals.py has been run.

Author: Mitchell Turner
Date: April 2025
"""

import pandas as pd
import os
import logging
from pathlib import Path
import numpy as np
from datetime import datetime

# --- CONFIGURATION ---
# File paths - match the structure from the first script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Input file - output from the first script
input_file = os.path.join(BASE_DIR, "output", "Market_Volume_Actuals.xlsx")

# Output enhanced summary file
output_dir = os.path.join(BASE_DIR, "output")
output_file = os.path.join(output_dir, "Market_Volume_Forecast.xlsx")

# Constants for special calculations
SPECIAL_FORECAST_MONTH = "2025-08"  # New project year month
SAMPLE_STATES = ["AL", "AR", "FL", "TX"]  # States to log for debugging

# Set up logging
def setup_logging():
    """Set up logging configuration"""
    # Ensure output directory exists before setting up file handler
    os.makedirs(output_dir, exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(os.path.join(output_dir, "forecast_calculation.log"), encoding='utf-8'),
            logging.StreamHandler()
        ])
    return logging.getLogger("forecast_calculator")

# Initialize logger
logger = setup_logging()

def get_previous_month(year, month):
    """
    Get the previous month as a (year, month) tuple.
    
    Args:
        year (int): The current year
        month (int): The current month (1-12)
        
    Returns:
        tuple: A (year, month) tuple for the previous month
    """
    if month == 1:  # January
        return (year - 1, 12)
    else:
        return (year, month - 1)

def get_year_month_string(year, month):
    """
    Convert year and month to string format YYYY-MM.
    
    Args:
        year (int): The year
        month (int): The month (1-12)
        
    Returns:
        str: Formatted string in YYYY-MM format
    """
    return f"{year}-{month:02d}"

def read_summary_file(file_path):
    """
    Read the summary Excel file and return the State Summary DataFrame.
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        pandas.DataFrame: DataFrame containing the State Summary data
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        Exception: For any other errors reading the file
    """
    try:
        logger.info(f"Reading summary file: {file_path}")
        # Read State Summary sheet
        state_summary = pd.read_excel(file_path, sheet_name="State Summary")
        logger.info(f"Successfully read State Summary: {len(state_summary)} states")
        return state_summary
    except FileNotFoundError:
        logger.error(f"Summary file not found: {file_path}")
        raise
    except Exception as e:
        logger.error(f"Error reading summary file: {e}", exc_info=True)
        raise

def identify_months_to_forecast(df):
    """Identify which months need forecasting by finding empty or NaN values in the TOTAL row"""
    state_col = df.columns[0]  # Usually "State"
    month_columns = [col for col in df.columns if col != state_col]
    # Sort the months chronologically
    month_columns.sort()
    # Get the TOTAL row
    total_row = df[df[state_col] == "TOTAL"]
    # Find months with valid data vs. months that need forecasting
    actual_months = []
    forecast_months = []
    for month in month_columns:
        if not total_row.empty:
            value = total_row[month].values[0]
            # Consider as needing forecasting if total is NaN, empty, or zero
            if pd.isna(value) or value == "" or value == 0:
                forecast_months.append(month)
            else:
                actual_months.append(month)
    # Find the latest month with actual data
    latest_actual = actual_months[-1] if actual_months else None
    logger.info(f"Latest month with actual data: {latest_actual}")
    logger.info(f"Months to forecast: {forecast_months}")
    return month_columns, actual_months, forecast_months, latest_actual

def calculate_forecasts(df, all_months, forecast_months, latest_actual):
    """Calculate forecasts for each state for the specified forecast months"""
    state_col = df.columns[0]  # Usually "State"
    forecasted_df = df.copy()
    # Convert all month columns to numeric
    for month in all_months:
        if month in forecasted_df.columns:
            forecasted_df[month] = pd.to_numeric(forecasted_df[month], errors='coerce').fillna(0)
    # Get membership row for ratio calculations
    membership_row = forecasted_df[forecasted_df[state_col] == "MEMBERSHIP"]
    # Process each forecast month in chronological order
    for forecast_month in sorted(forecast_months):
        # Parse forecast month into year and month numbers
        forecast_year, forecast_month_num = map(int, forecast_month.split('-'))
        # Special handling for 2025-08 (new project year)
        if forecast_month == SPECIAL_FORECAST_MONTH:
            logger.info(f"Using special logic for {forecast_month} as it's the start of a new project year")
            # Previous year's same month is 2024-08
            prev_year_same_month = "2024-08"
            # Check if we have the required data
            if prev_year_same_month not in all_months:
                logger.warning(f"Missing required data {prev_year_same_month} for {forecast_month}, skipping")
                continue
            # Calculate membership adjustment ratio
            if not membership_row.empty:
                # For 2025-08, we need the forecast of membership for that month
                # Since we don't have it yet, we'll use the most recent membership and apply growth
                # Get most recent actual membership
                current_membership = membership_row[latest_actual].values[0]
                # Get membership from previous year same month
                prev_year_membership = membership_row[prev_year_same_month].values[0]
                # Calculate adjustment ratio
                if prev_year_membership != 0 and not pd.isna(prev_year_membership):
                    # We can estimate future membership based on YoY growth
                    membership_ratio = current_membership / prev_year_membership
                    estimated_membership = prev_year_membership * membership_ratio
                    membership_adjustment = estimated_membership / prev_year_membership
                    logger.info(f"Membership adjustment for {forecast_month}: {membership_adjustment:.4f}")
                else:
                    membership_adjustment = 1.0
                    logger.warning(f"Cannot calculate membership adjustment for {forecast_month} - using default 1.0")
            else:
                membership_adjustment = 1.0
                logger.warning("No membership row found, using default membership adjustment of 1.0")
            # Calculate forecast for each state
            for idx, row in forecasted_df.iterrows():
                state = row[state_col]
                # Skip TOTAL and MEMBERSHIP rows
                if state in ["TOTAL", "MEMBERSHIP"]:
                    continue
                # Get previous year's value for same month
                prev_year_value = row[prev_year_same_month]
                # Skip if missing required data
                if pd.isna(prev_year_value):
                    logger.debug(f"Missing data for {state} in {prev_year_same_month}, using zero")
                    forecasted_df.at[idx, forecast_month] = 0
                    continue
                # Apply the special formula: previous_year_same_month * membership_adjustment
                forecast_value = prev_year_value * membership_adjustment
                # Ensure no negative values and round to integer
                forecast_value = max(0, round(forecast_value))
                # Update the DataFrame
                forecasted_df.at[idx, forecast_month] = forecast_value
                # Log for debugging
                if state in SAMPLE_STATES:  # Log a few sample states
                    logger.info(
                        f"{state} {forecast_month}: {prev_year_value} * {membership_adjustment} = {forecast_value}")
            # Recalculate TOTAL row for this forecast month
            total_idx = forecasted_df[forecasted_df[state_col] == "TOTAL"].index
            if len(total_idx) > 0:
                # Sum all state values, excluding TOTAL and MEMBERSHIP
                state_data = forecasted_df[~forecasted_df[state_col].isin(["TOTAL", "MEMBERSHIP"])]
                total_value = state_data[forecast_month].sum()
                forecasted_df.at[total_idx[0], forecast_month] = total_value
                logger.info(f"Total forecast for {forecast_month}: {total_value}")
        else:
            # Standard forecasting logic for other months
            # Get previous month in same year
            prev_year, prev_month_num = get_previous_month(forecast_year, forecast_month_num)
            prev_month_str = get_year_month_string(prev_year, prev_month_num)
            # Get same month in previous year
            prev_year_month = get_year_month_string(forecast_year - 1, forecast_month_num)
            # Get previous month in previous year
            prev_year_prev_year, prev_year_prev_month_num = get_previous_month(forecast_year - 1, forecast_month_num)
            prev_year_prev_month_str = get_year_month_string(prev_year_prev_year, prev_year_prev_month_num)
            logger.info(
                f"Forecasting {forecast_month} using data from {prev_month_str}, {prev_year_month}, and {prev_year_prev_month_str}")
            # Check if we have all required months
            required_months = [prev_month_str, prev_year_month, prev_year_prev_month_str]
            if not all(month in all_months for month in required_months):
                logger.warning(f"Missing required data for {forecast_month}, skipping")
                continue
            # Calculate membership ratio
            membership_ratio = 1.0  # Default (no change)
            if not membership_row.empty:
                current_membership = membership_row[prev_month_str].values[0]
                prev_year_membership = membership_row[prev_year_month].values[0]
                if prev_year_membership != 0 and not pd.isna(prev_year_membership):
                    membership_ratio = current_membership / prev_year_membership
                    logger.info(f"Membership ratio for {forecast_month}: {membership_ratio:.4f}")
            # Create forecast column if it doesn't exist
            if forecast_month not in forecasted_df.columns:
                forecasted_df[forecast_month] = 0
            # Calculate forecast for each state
            for idx, row in forecasted_df.iterrows():
                state = row[state_col]
                # Skip TOTAL and MEMBERSHIP rows - we'll calculate TOTAL after all states
                if state in ["TOTAL", "MEMBERSHIP"]:
                    continue
                # Get values needed for calculation
                current_value = row[prev_month_str]
                prev_year_value = row[prev_year_month]
                prev_year_prev_month_value = row[prev_year_prev_month_str]
                # Skip if missing required data
                if pd.isna(current_value) or pd.isna(prev_year_value) or pd.isna(prev_year_prev_month_value):
                    logger.debug(f"Missing data for {state} in {forecast_month}, using zero")
                    forecasted_df.at[idx, forecast_month] = 0
                    continue
                # Calculate last year's change
                last_year_change = prev_year_value - prev_year_prev_month_value
                # Apply the formula: current + (last_year_change * membership_ratio)
                forecast_value = current_value + (last_year_change * membership_ratio)
                # Ensure no negative values
                forecast_value = max(0, forecast_value)
                # Round to integer
                forecast_value = round(forecast_value)
                # Update the DataFrame
                forecasted_df.at[idx, forecast_month] = forecast_value
                # Log for debugging
                if state in SAMPLE_STATES:  # Log a few sample states
                    logger.info(
                        f"{state} {forecast_month}: {current_value} + ({last_year_change} * {membership_ratio}) = {forecast_value}")
            # Recalculate TOTAL row for this forecast month
            total_idx = forecasted_df[forecasted_df[state_col] == "TOTAL"].index
            if len(total_idx) > 0:
                # Sum all state values, excluding TOTAL and MEMBERSHIP
                state_data = forecasted_df[~forecasted_df[state_col].isin(["TOTAL", "MEMBERSHIP"])]
                total_value = state_data[forecast_month].sum()
                forecasted_df.at[total_idx[0], forecast_month] = total_value
                logger.info(f"Total forecast for {forecast_month}: {total_value}")
    # Make sure all columns are in correct order
    all_columns = [state_col] + sorted([col for col in forecasted_df.columns if col != state_col])
    forecasted_df = forecasted_df[all_columns]
    return forecasted_df

def write_forecast_excel(df, actual_months, forecast_months, output_path):
    """Write forecasted data to Excel with formatting to distinguish actuals from forecasts"""
    logger.info(f"Writing forecast Excel file: {output_path}")
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Add a worksheet for the forecasted data
        df.to_excel(writer, sheet_name="State Forecast", index=False)
        worksheet = writer.sheets["State Forecast"]
        # Create formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1
        })
        actual_format = workbook.add_format({
            'num_format': '#,##0',
            'bg_color': '#D9EAD3',  # Light green
            'border': 1
        })
        forecast_format = workbook.add_format({
            'num_format': '#,##0',
            'bg_color': '#FCE4D6',  # Light orange
            'border': 1
        })
        august_format = workbook.add_format({
            'num_format': '#,##0',
            'bg_color': '#D0E0E3',  # Light blue for 2025-08 special forecast
            'border': 1
        })
        total_actual_format = workbook.add_format({
            'bold': True,
            'num_format': '#,##0',
            'bg_color': '#B6D7A8',  # Darker green
            'border': 1
        })
        total_forecast_format = workbook.add_format({
            'bold': True,
            'num_format': '#,##0',
            'bg_color': '#F9CB9C',  # Darker orange
            'border': 1
        })
        total_august_format = workbook.add_format({
            'bold': True,
            'num_format': '#,##0',
            'bg_color': '#9FC5E8',  # Darker blue for 2025-08 special forecast
            'border': 1
        })
        membership_format = workbook.add_format({
            'italic': True,
            'bold': True,
            'num_format': '#,##0',
            'bg_color': '#EEEEEE',  # Light grey
            'border': 1
        })
        state_col = df.columns[0]  # Usually "State"
        # Format headers
        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, header_format)
        # Format data
        for row_idx, row in df.iterrows():
            state = row[state_col]
            worksheet.write(row_idx + 1, 0, state)  # Write state name
            # Format each month column
            for col_idx, month in enumerate(df.columns[1:], start=1):
                value = row[month]
                # Choose format based on row type and whether month is actual or forecast
                if state == "TOTAL":
                    if month in actual_months:
                        cell_format = total_actual_format
                    elif month == SPECIAL_FORECAST_MONTH:
                        cell_format = total_august_format
                    else:
                        cell_format = total_forecast_format
                elif state == "MEMBERSHIP":
                    cell_format = membership_format
                else:
                    if month in actual_months:
                        cell_format = actual_format
                    elif month == SPECIAL_FORECAST_MONTH:
                        cell_format = august_format
                    else:
                        cell_format = forecast_format
                # Write value with appropriate format
                if pd.isna(value) or value == "":
                    worksheet.write_string(row_idx + 1, col_idx, "N/A", cell_format)
                else:
                    try:
                        numeric_value = float(value)
                        worksheet.write_number(row_idx + 1, col_idx, numeric_value, cell_format)
                    except (ValueError, TypeError):
                        worksheet.write_string(row_idx + 1, col_idx, str(value), cell_format)
        # Add filtering and freeze panes
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
        worksheet.freeze_panes(1, 1)
        # Adjust column widths
        worksheet.set_column(0, 0, 20)  # State column
        worksheet.set_column(1, len(df.columns) - 1, 12)  # Month columns
        # Add a legend for the colors
        row_pos = len(df) + 3
        worksheet.write(row_pos, 0, "Legend:", workbook.add_format({'bold': True}))
        worksheet.write(row_pos + 1, 0, "Actual Data", actual_format)
        worksheet.write(row_pos + 1, 1, "Standard Forecast", forecast_format)
        worksheet.write(row_pos + 1, 2, f"New Project Year ({SPECIAL_FORECAST_MONTH})", august_format)
        # Add forecasting methodology explanation
        row_pos += 3
        worksheet.merge_range(row_pos, 0, row_pos, 3, "Forecasting Methodology:", workbook.add_format({'bold': True}))
        row_pos += 1
        # Standard forecast method
        worksheet.merge_range(row_pos, 0, row_pos, 5, "Standard forecast (most months):",
                              workbook.add_format({'italic': True, 'bold': True}))
        row_pos += 1
        worksheet.merge_range(row_pos, 0, row_pos, 5, "Current Value + (Last Year's Change * Membership Ratio)",
                              workbook.add_format({'bold': True}))
        row_pos += 1
        worksheet.merge_range(row_pos, 0, row_pos, 5, "Where:", workbook.add_format({'italic': True}))
        row_pos += 1
        worksheet.merge_range(row_pos, 0, row_pos, 5, "- Current Value = Value from previous month",
                              workbook.add_format())
        row_pos += 1
        worksheet.merge_range(row_pos, 0, row_pos, 5,
                              "- Last Year's Change = Same month previous year minus its previous month",
                              workbook.add_format())
        row_pos += 1
        worksheet.merge_range(row_pos, 0, row_pos, 5,
                              "- Membership Ratio = Current membership / Previous year's membership",
                              workbook.add_format())
        # August 2025 special method
        row_pos += 2
        worksheet.merge_range(row_pos, 0, row_pos, 5, "Special forecast for August 2025 (new project year):",
                              workbook.add_format({'italic': True, 'bold': True}))
        row_pos += 1
        worksheet.merge_range(row_pos, 0, row_pos, 5, "Previous Year's August * Membership Adjustment",
                              workbook.add_format({'bold': True}))
        row_pos += 1
        worksheet.merge_range(row_pos, 0, row_pos, 5,
                              "Where Membership Adjustment = Estimated future membership / Previous August membership",
                              workbook.add_format())
        # Add timestamp
        row_pos += 2
        worksheet.write(row_pos, 0, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    logger.info(f"Forecast Excel file successfully created: {output_path}")

def main():
    """
    Main execution function that orchestrates the forecasting process:
    1. Reads the input file with actual data
    2. Identifies which months need forecasting
    3. Calculates the forecasts using appropriate algorithms
    4. Writes the results to a formatted Excel file
    
    Returns:
        bool: True if successful, False if an error occurred
    """
    try:
        # Ensure output directory exists
        Path(output_dir).mkdir(exist_ok=True)
        
        # Check if input file exists
        if not os.path.exists(input_file):
            logger.error(f"Input file not found: {input_file}")
            print(f"\nError: The summary file '{input_file}' does not exist.")
            print(f"Please run the first script (Chart_Actuals.py) to generate the summary file first.")
            return False
        # Read the summary file
        state_df = read_summary_file(input_file)
        # Identify months for forecasting
        all_months, actual_months, forecast_months, latest_actual = identify_months_to_forecast(state_df)
        if not forecast_months:
            logger.info("No months found that need forecasting. All data appears to be present.")
            print("\nAll months already have data. No forecasting needed.")
            return True
        if not latest_actual:
            logger.error("Could not identify the latest month with actual data.")
            print("\nError: Could not identify any months with actual data.")
            return False
        # Calculate forecasts
        forecasted_df = calculate_forecasts(state_df, all_months, forecast_months, latest_actual)
        # Write to Excel with formatting
        write_forecast_excel(forecasted_df, actual_months, forecast_months, output_file)
        print(f"\nSUMMARY:")
        print(f"  - Forecast completed successfully")
        print(f"  - Latest actual data: {latest_actual}")
        print(f"  - Forecasted {len(forecast_months)} months")
        print(f"  - Special calculation used for August 2025 (new project year)")
        print(f"  - Output saved to: {output_file}")
        return True
    except Exception as e:
        logger.error(f"Error in main execution: {e}", exc_info=True)
        print(f"\nAn error occurred: {e}")
        return False

if __name__ == "__main__":
    # Print startup information
    print(f"\nMarket Volume Forecast Generator")
    print(f"--------------------------------")
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print(f"Logs: {os.path.join(output_dir, 'forecast_calculation.log')}")
    print(f"--------------------------------\n")
    
    # Run the main function
    success = main()
    
    if success:
        print(f"\nSUCCESS: Forecast generated successfully!")
        print(f"Output file: {output_file}")
    else:
        print(f"\nFAILED: Script encountered errors. Check the log for details.")
        
    print(f"--------------------------------")
