"""
Wrangle IxI (Industry-by-Industry) tables from SUT-98-22.xlsx

This script separates the IxI tables by year into:
- Main matrix (industry-to-industry flows)
- Final demand matrix (consumption, investment, exports)
- GVA matrix (value added components)
"""

import pandas as pd
import os

# Configuration
INPUT_FILE = "SUT-98-22.xlsx"
OUTPUT_DIR = "output"

# Column indices (0-indexed)
COL_YEAR = 0
COL_SIC = 1
COL_NAME = 2
COL_MATRIX_START = 3
COL_MATRIX_END = 100  # Last industry column (inclusive)
COL_TOTAL_INTERMEDIATE = 101
COL_FINAL_DEMAND_START = 102
COL_FINAL_DEMAND_END = 116  # Total use for industry output

# SIC codes that represent GVA/input rows (not industries)
GVA_SIC_CODES = ['TDU', 'RUKImp', 'RoWImp', 'TIU', 'TlSPrds', 'TlSPrdn', 'CoE', 'GOS', 'GVA', 'TOut']

def load_ixi_sheet(filepath):
    """Load the IxI sheet from the Excel file."""
    print(f"Loading {filepath}...")
    xl = pd.ExcelFile(filepath)
    df = pd.read_excel(xl, 'IxI', header=None)
    print(f"Loaded sheet with shape: {df.shape}")
    return df

def get_column_headers(df):
    """Extract column headers from row 6."""
    headers = df.iloc[6, COL_MATRIX_START:].tolist()
    return headers

def extract_year_data(df, year):
    """Extract all data for a specific year."""
    year_mask = df.iloc[:, COL_YEAR] == year
    year_data = df[year_mask].copy()
    return year_data

def separate_matrices(year_data, industry_headers, final_demand_headers):
    """
    Separate a year's data into main matrix, final demand, and GVA components.

    Returns:
        main_matrix: DataFrame of industry-to-industry flows
        final_demand: DataFrame of final demand columns for industries
        gva_matrix: DataFrame of GVA/input rows
    """
    # Separate industry rows from GVA rows
    sic_codes = year_data.iloc[:, COL_SIC].astype(str)
    industry_mask = ~sic_codes.isin(GVA_SIC_CODES)

    industry_rows = year_data[industry_mask].copy()
    gva_rows = year_data[~industry_mask].copy()

    # Extract main matrix (industry x industry)
    main_matrix = industry_rows.iloc[:, COL_MATRIX_START:COL_MATRIX_END+1].copy()
    main_matrix.columns = industry_headers[:len(main_matrix.columns)]
    main_matrix.index = industry_rows.iloc[:, COL_NAME].values

    # Extract final demand (industry x final demand categories)
    final_demand = industry_rows.iloc[:, COL_FINAL_DEMAND_START:COL_FINAL_DEMAND_END+1].copy()
    final_demand.columns = final_demand_headers
    final_demand.index = industry_rows.iloc[:, COL_NAME].values

    # Extract GVA matrix (GVA components x industries)
    gva_matrix = gva_rows.iloc[:, COL_MATRIX_START:COL_MATRIX_END+1].copy()
    gva_matrix.columns = industry_headers[:len(gva_matrix.columns)]
    gva_matrix.index = gva_rows.iloc[:, COL_NAME].values

    return main_matrix, final_demand, gva_matrix

def save_matrices(year, main_matrix, final_demand, gva_matrix, output_dir):
    """Save matrices to CSV files."""
    year_dir = os.path.join(output_dir, str(year))
    os.makedirs(year_dir, exist_ok=True)

    main_matrix.to_csv(os.path.join(year_dir, "main_matrix.csv"))
    final_demand.to_csv(os.path.join(year_dir, "final_demand.csv"))
    gva_matrix.to_csv(os.path.join(year_dir, "gva_matrix.csv"))

    print(f"  Saved: main_matrix ({main_matrix.shape}), "
          f"final_demand ({final_demand.shape}), "
          f"gva_matrix ({gva_matrix.shape})")

def main():
    # Load data
    df = load_ixi_sheet(INPUT_FILE)

    # Get headers
    industry_headers = df.iloc[6, COL_MATRIX_START:COL_MATRIX_END+1].tolist()
    final_demand_headers = df.iloc[6, COL_FINAL_DEMAND_START:COL_FINAL_DEMAND_END+1].tolist()

    print(f"\nIndustry columns: {len(industry_headers)}")
    print(f"Final demand columns: {len(final_demand_headers)}")

    # Get unique years
    years = sorted(df.iloc[7:, COL_YEAR].dropna().unique())
    print(f"\nYears found: {years}")

    # Create output directory
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Process each year
    print("\nProcessing years...")
    for year in years:
        print(f"\nYear {year}:")
        year_data = extract_year_data(df, year)

        main_matrix, final_demand, gva_matrix = separate_matrices(
            year_data, industry_headers, final_demand_headers
        )

        save_matrices(year, main_matrix, final_demand, gva_matrix, OUTPUT_DIR)

    print(f"\nDone! Output saved to '{OUTPUT_DIR}/' directory.")
    print("\nStructure:")
    print("  output/")
    print("    └── {year}/")
    print("        ├── main_matrix.csv     (98x98 industry flows)")
    print("        ├── final_demand.csv    (98x15 final demand)")
    print("        └── gva_matrix.csv      (10x98 value added)")

if __name__ == "__main__":
    main()
