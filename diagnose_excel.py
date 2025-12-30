"""
Diagnostic Tool - Inspect Raw Excel Files
Use this to see exactly what's in the first rows of each file
"""
import pandas as pd
from pathlib import Path
from Config import Config


def inspect_file_raw(filename, num_rows=10):
    """
    Show raw contents of Excel file without any header assumptions
    Useful for seeing where the actual data starts
    """
    filepath = Path(Config.BASE_FOLDER) / filename
    
    if not filepath.exists():
        print(f"ERROR: File not found: {filepath}")
        return None
    
    print("=" * 80)
    print(f"RAW FILE INSPECTION: {filename}")
    print("=" * 80)
    
    # Load without assuming header location
    df_raw = pd.read_excel(filepath, header=None, nrows=num_rows)
    
    print(f"\nShowing first {len(df_raw)} rows (0-indexed):\n")
    
    # Display with row numbers
    for idx, row in df_raw.iterrows():
        print(f"Row {idx}:")
        # Show each cell
        for col_idx, value in enumerate(row):
            if pd.notna(value):
                value_str = str(value)[:50]  # Truncate long values
                print(f"  Col {col_idx}: {value_str}")
        print()
    
    # Try to auto-detect header
    print("=" * 80)
    print("AUTO-DETECTION ANALYSIS")
    print("=" * 80)
    
    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        filled_count = row.notna().sum()
        text_count = sum(1 for val in row if isinstance(val, str))
        unique_count = row.dropna().nunique()
        
        # Check for header keywords
        header_keywords = ['project', 'name', 'date', 'assigned', 'estimator', 'location']
        row_text = ' '.join([str(val).lower() for val in row if pd.notna(val)])
        keyword_matches = [kw for kw in header_keywords if kw in row_text]
        
        print(f"\nRow {idx} Analysis:")
        print(f"  - Filled cells: {filled_count}/{len(row)}")
        print(f"  - Text cells: {text_count}")
        print(f"  - Unique values: {unique_count}")
        if keyword_matches:
            print(f"  - Header keywords found: {', '.join(keyword_matches)}")
        
        # Likelihood assessment
        if filled_count > len(row) * 0.5 and text_count > len(row) * 0.5 and unique_count > 3:
            print(f"  HIGH likelihood this is the header row")
        elif filled_count < 3:
            print(f"  LOW likelihood (mostly empty)")
        elif unique_count == 1:
            print(f"  WARNING: Probably a title row (all same value)")
    
    print("\n" + "=" * 80)
    return df_raw


def inspect_all_configured_files():
    """Inspect all files configured in Config.REPORT_FILES"""
    
    print("=" * 80)
    print("INSPECTING ALL CONFIGURED FILES")
    print("=" * 80)
    
    for report_key, filename in Config.REPORT_FILES.items():
        inspect_file_raw(filename, num_rows=8)
        print("\n" + "=" * 80)
        input("Press ENTER to continue to next file...")


def test_header_detection():
    """
    Test the automatic header detection on all configured files
    """
    from Data_Loader import SAPDataLoader
    
    print("=" * 80)
    print("TESTING HEADER DETECTION")
    print("=" * 80)
    
    loader = SAPDataLoader()
    
    for report_key, filename in Config.REPORT_FILES.items():
        filepath = Path(Config.BASE_FOLDER) / filename
        
        if not filepath.exists():
            print(f"\nERROR: {filename} - File not found")
            continue
        
        print(f"\n{filename}")
        
        # Detect header row
        header_row = loader._find_header_row(filepath)
        print(f"   Detected header at row: {header_row}")
        
        # Load a few rows to verify
        df = pd.read_excel(filepath, header=header_row, nrows=3)
        print(f"   Columns detected: {', '.join(df.columns[:5].tolist())}")
        if len(df.columns) > 5:
            print(f"      ... and {len(df.columns) - 5} more")
        
        # Show first row of actual data
        if len(df) > 0:
            print(f"   First data row sample:")
            for col in df.columns[:5]:
                val = df[col].iloc[0]
                if pd.notna(val):
                    print(f"      {col}: {str(val)[:40]}")
    
    print("\n" + "=" * 80)
    print("Header detection test complete!")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        # Inspect specific file
        filename = sys.argv[1]
        inspect_file_raw(filename)
    else:
        # Show menu
        print("\n" + "=" * 80)
        print("EXCEL DIAGNOSTIC TOOL")
        print("=" * 80)
        print("\nOptions:")
        print("  1. Test header detection on all configured files")
        print("  2. Inspect specific file in detail")
        print("  3. Inspect all configured files (interactive)")
        
        choice = input("\nEnter choice (1-3): ").strip()
        
        if choice == "1":
            test_header_detection()
        elif choice == "2":
            filename = input("\nEnter filename (e.g., sd-09 Cost Estimating Schedule.xlsx): ").strip()
            inspect_file_raw(filename)
        elif choice == "3":
            inspect_all_configured_files()
        else:
            print("Invalid choice")

