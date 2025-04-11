#!/usr/bin/env python3
"""
Market Volume Actuals Processor

This script processes Excel files from project directories and creates a summary
of actual chart volumes by state. This is the first script in the pipeline,
generating the input for the Market Volume Forecast script.

Author: Mitchell Turner
Date: April 2025
"""

import os
import pandas as pd
from collections import defaultdict
from datetime import datetime
import xlsxwriter
import matplotlib.pyplot as plt

# --- CONFIGURATION ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
root_dirs = ["Project2023", "Project2024", "Project2025"]  # Project directories relative to script
output_dir = os.path.join(BASE_DIR, "output")
output_file = os.path.join(output_dir, "Market_Volume_Actuals.xlsx")
start_month = "2023-08"
end_month = "2026-07"

# --- Helper: generate month list ---
def generate_months(start, end):
    """
    Generate a list of months between start and end dates.
    
    Args:
        start (str): Start month in format 'YYYY-MM'
        end (str): End month in format 'YYYY-MM'
        
    Returns:
        list: List of months in format 'YYYY-MM'
    """
    dates = pd.date_range(start=start + "-01", end=end + "-01", freq="MS")
    return [d.strftime("%Y-%m") for d in dates]

all_months = generate_months(start_month, end_month)

# --- Helper: standardize column names ---
def standardize_columns(df):
    """
    Standardize column names in the DataFrame for consistent processing.
    
    Args:
        df (pandas.DataFrame): DataFrame with original column names
        
    Returns:
        pandas.DataFrame: DataFrame with standardized column names
    """
    col_map = {}
    for col in df.columns:
        col_lower = col.strip().lower()
        if "vendor" in col_lower:
            col_map[col] = "Ret_Vendor"
        elif "chart" in col_lower and "count" in col_lower:
            col_map[col] = "Chartcount"
        elif "state" in col_lower:
            col_map[col] = "State"
    return df.rename(columns=col_map)

def main():
    """
    Main execution function that processes Excel files and generates summary output.
    
    Returns:
        bool: True if successful, False if an error occurred
    """
    try:
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        # --- Data holders ---
        state_data = defaultdict(lambda: {month: None for month in all_months})
        
        # --- Process Excel files ---
        for folder in root_dirs:
            folder_path = os.path.join(BASE_DIR, folder)
            if not os.path.exists(folder_path):
                print(f"Warning: Project directory not found: {folder_path}")
                continue
                
            print(f"Processing files in {folder}...")
            files_processed = 0
            
            for file in os.listdir(folder_path):
                if file.endswith(".xlsx") and file[:7] in all_months:
                    month = file[:7]
                    file_path = os.path.join(folder_path, file)
                    try:
                        df = pd.read_excel(file_path)
                        df = standardize_columns(df)
        
                        if {"Ret_Vendor", "Chartcount", "State"}.issubset(df.columns):
                            df["Chartcount"] = pd.to_numeric(df["Chartcount"], errors="coerce").fillna(0)
                            state_group = df.groupby("State")["Chartcount"].sum()
                            for state, count in state_group.items():
                                state_data[state][month] = count
                            files_processed += 1
                        else:
                            print(f"  Warning: Required columns missing in {file}")
                    except Exception as e:
                        print(f"  Error processing {file_path}: {e}")
            
            print(f"  Processed {files_processed} files from {folder}")
        
        # --- Convert to DataFrames ---
        state_df = pd.DataFrame.from_dict(state_data, orient="index", columns=all_months).fillna("")
        state_df.index.name = "State"
        state_df.reset_index(inplace=True)
        
        # --- Load Membership Reference ---
        mem_file = os.path.join(BASE_DIR, "MEM_REF/MP_MEM_REF.xlsx")
        if os.path.exists(mem_file):
            try:
                mem_df = pd.read_excel(mem_file, index_col=0)
                mem_df.columns = [pd.to_datetime(col, format="%YM%m").strftime("%Y-%m") for col in mem_df.columns]
                membership_row = mem_df.loc["Total"].reindex(all_months).to_frame().T
            except Exception as e:
                print(f"Error loading membership reference: {e}")
                # Create empty membership row as fallback
                membership_row = pd.DataFrame([[0] * len(all_months)], columns=all_months)
        else:
            print(f"Warning: Membership reference file not found: {mem_file}")
            # Create empty membership row as fallback
            membership_row = pd.DataFrame([[0] * len(all_months)], columns=all_months)
        
        # --- Define colors ---
        project_colors = {
            "2023": "#FFF2CC",
            "2024": "#D9EAD3",
            "2025": "#D0E0E3",
            "2026": "#F4CCCC",
        }
        
        def get_project_color(month):
            """
            Get color for a specific month based on project year.
            
            Args:
                month (str): Month in format 'YYYY-MM'
                
            Returns:
                str: Hex color code
            """
            if "2023-08" <= month <= "2024-07":
                return project_colors["2023"]
            elif "2024-08" <= month <= "2025-07":
                return project_colors["2024"]
            elif "2025-08" <= month <= "2026-07":
                return project_colors["2025"]
            else:
                return project_colors["2026"]
        
        # --- Generate Excel Output with Summary Tab ---
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            # Main sheet: State Summary
            sheet_name = "State Summary"
            totals = state_df[all_months].apply(pd.to_numeric, errors='coerce').sum().to_frame().T
            totals.insert(0, "State", "TOTAL")
        
            membership_clone = membership_row.copy()
            membership_clone.columns = all_months
            membership_clone.insert(0, "State", "MEMBERSHIP")
        
            final_df = pd.concat([state_df, totals, membership_clone], ignore_index=True)
            final_df.to_excel(writer, index=False, sheet_name=sheet_name)
        
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
        
            header_format = workbook.add_format({"bold": True, "bg_color": "#D9E1F2"})
            worksheet.freeze_panes(1, 1)
        
            for col_num, value in enumerate(final_df.columns):
                worksheet.write(0, col_num, value, header_format)
        
            worksheet.set_column(0, 0, 20)
            worksheet.set_column(1, len(all_months), 12)
        
            num_rows = len(final_df)
            for col_idx, month in enumerate(all_months, start=1):
                color = get_project_color(month)
                for row in range(1, num_rows + 1):
                    value = final_df.iloc[row - 1, col_idx]
                    label = final_df.iloc[row - 1, 0]
        
                    if pd.isna(value) or value == "":
                        value = ""
        
                    if label == "TOTAL":
                        cell_format = workbook.add_format({"bg_color": color, "bold": True, "num_format": "#,##0"})
                    elif label == "MEMBERSHIP":
                        cell_format = workbook.add_format({"bg_color": "#EAEAEA", "italic": True, "bold": True, "num_format": "#,##0"})
                    else:
                        cell_format = workbook.add_format({"bg_color": color, "num_format": "#,##0"})
        
                    worksheet.write(row, col_idx, value, cell_format)
        
            # Summary sheet with chart
            chart_sheet = workbook.add_worksheet("Summary")
        
            # Create a simple line chart using totals over time
            chart_data = totals[all_months].T
            chart_data.columns = ["Total"]
            chart_data.reset_index(inplace=True)
            chart_data.columns = ["Month", "Total"]
        
            for i, (month, value) in enumerate(zip(chart_data["Month"], chart_data["Total"])):  # Write raw data for chart
                chart_sheet.write(i, 0, month)
                chart_sheet.write(i, 1, value)
        
            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
                'name':       'Total Chart Volumes',
                'categories': ['Summary', 0, 0, len(chart_data)-1, 0],
                'values':     ['Summary', 0, 1, len(chart_data)-1, 1],
                'line':       {'color': 'blue'}
            })
            chart.set_title({'name': 'Total Chart Volumes Over Time'})
            chart.set_x_axis({'name': 'Month'})
            chart.set_y_axis({'name': 'Chart Count'})
        
            chart_sheet.insert_chart('D2', chart)
            
        print(f"\n✅ Excel file saved: {output_file}")
        return True
    
    except Exception as e:
        print(f"\n❌ Error in main execution: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    # Print startup information
    print(f"\nMarket Volume Actuals Processor")
    print(f"-------------------------------")
    print(f"Processing directories: {', '.join(root_dirs)}")
    print(f"Output file: {output_file}")
    print(f"-------------------------------\n")
    
    # Run the main function
    success = main()
    
    if success:
        print(f"\nSUCCESS: Actuals processing complete!")
        print(f"You can now run Market_Volume_Forecast.py to generate forecasts.")
    else:
        print(f"\nFAILED: Script encountered errors.")
        
    print(f"-------------------------------")
