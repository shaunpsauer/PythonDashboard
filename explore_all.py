"""
Utility to explore ALL Excel files in the SAP folder
Use this to discover what data is available in other reports
Now with automatic header detection!
"""
import pandas as pd
from pathlib import Path
from Config import Config


def find_header_row(filepath, max_rows=20):
    """Find the header row in an Excel file"""
    df_sample = pd.read_excel(filepath, header=None, nrows=max_rows, engine='openpyxl')
    
    header_keywords = [
        'project', 'name', 'date', 'status', 'phase', 'number',
        'assigned', 'estimator', 'location', 'region', 'pmo',
        'start', 'end', 'complete', 'due', 'schedule', 'id'
    ]
    
    best_row = 0
    best_score = 0
    
    for idx in range(min(max_rows, len(df_sample))):
        row = df_sample.iloc[idx]
        score = 0
        
        filled_ratio = row.notna().sum() / len(row)
        score += filled_ratio * 100
        
        text_values = sum(1 for val in row if isinstance(val, str))
        score += (text_values / len(row)) * 50
        
        if row.notna().any():
            row_text = ' '.join([str(val).lower() for val in row if pd.notna(val)])
            keyword_matches = sum(1 for keyword in header_keywords if keyword in row_text)
            score += keyword_matches * 30
        
        unique_values = row.dropna().nunique()
        if unique_values > 3:
            score += 20
        elif unique_values == 1:
            score -= 50
        
        if score > best_score:
            best_score = score
            best_row = idx
    
    return best_row


def explore_all_files():
    """Scan all Excel files and show their structure"""
    
    folder = Path(Config.BASE_FOLDER)
    
    if not folder.exists():
        print(f"ERROR: Folder not found: {folder}")
        return
    
    # Find all Excel files
    excel_files = sorted(folder.glob("*.xlsx"))
    
    print("=" * 80)
    print(f"EXPLORING ALL FILES IN: {folder.name}")
    print("=" * 80)
    print(f"\nFound {len(excel_files)} Excel files")
    print("(Now with automatic header detection!)\n")
    
    results = []
    
    for i, filepath in enumerate(excel_files, 1):
        try:
            print(f"\n[{i}/{len(excel_files)}] {filepath.name}")
            print("-" * 80)
            
            # Auto-detect header row
            header_row = find_header_row(filepath)
            if header_row > 0:
                print(f"   Header detected at row {header_row + 1} (skipping {header_row} title rows)")
            
            # Load with correct header
            df = pd.read_excel(filepath, header=header_row, nrows=5, engine='openpyxl')
            
            # Clean column names
            df.columns = df.columns.astype(str).str.strip()
            
            # Remove empty rows/columns
            df = df.dropna(how='all')
            unnamed_cols = [col for col in df.columns if 'Unnamed' in str(col) and df[col].isna().all()]
            df = df.drop(columns=unnamed_cols)
            
            # Store info
            info = {
                'filename': filepath.name,
                'header_row': header_row,
                'rows': len(df),
                'columns': len(df.columns),
                'column_list': list(df.columns)
            }
            results.append(info)
            
            # Display
            print(f"   Columns ({len(df.columns)}): {', '.join(df.columns[:8])}")
            if len(df.columns) > 8:
                print(f"   ... and {len(df.columns) - 8} more columns")
            
            # Show sample data for first few columns
            if len(df) > 0:
                print("\n   Sample data (first 2 rows, first 5 columns):")
                print(df.iloc[:2, :min(5, len(df.columns))].to_string(index=False))
            
        except Exception as e:
            print(f"   ERROR loading: {e}")
            continue
    
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    
    # Group by prefix (sd-01, sd-02, etc.)
    print("\nFiles by category:")
    for info in results:
        filename = info['filename']
        prefix = filename.split()[0] if ' ' in filename else filename
        print(f"   {prefix:20s} - {info['columns']:2d} columns - {filename}")
    
    print("\nExploration complete!")
    print("\nTip: Look for files with columns like:")
    print("   - 'Project Number', 'Project Name'")
    print("   - 'Assigned', 'Estimator', 'Owner'")
    print("   - 'Start Date', 'Due Date', 'Complete'")
    print("   - 'Location', 'Region', 'PMO'")
    
    return results


def find_files_with_column(search_term):
    """Find which files contain a specific column name"""
    
    folder = Path(Config.BASE_FOLDER)
    excel_files = folder.glob("*.xlsx")
    
    matches = []
    
    print(f"\nSearching for columns containing: '{search_term}'")
    print("-" * 80)
    
    for filepath in excel_files:
        try:
            # Auto-detect header
            header_row = find_header_row(filepath)
            
            # Load with correct header
            df = pd.read_excel(filepath, header=header_row, nrows=1, engine='openpyxl')
            df.columns = df.columns.astype(str).str.strip()
            
            # Check if any column contains the search term
            matching_cols = [col for col in df.columns 
                           if search_term.lower() in col.lower()]
            
            if matching_cols:
                matches.append((filepath.name, matching_cols, header_row))
                print(f"FOUND: {filepath.name}")
                if header_row > 0:
                    print(f"   (Header at row {header_row + 1})")
                print(f"   Columns: {', '.join(matching_cols)}")
        
        except Exception as e:
            continue
    
    if not matches:
        print(f"ERROR: No files found with columns containing '{search_term}'")
    else:
        print(f"\nFound {len(matches)} files with matching columns")
    
    return matches


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        # Search mode: python explore_all.py "assignee"
        search_term = sys.argv[1]
        find_files_with_column(search_term)
    else:
        # Full exploration mode
        results = explore_all_files()
        
        # Offer search
        print("\n" + "=" * 80)
        print("Want to search for specific columns?")
        print("Usage: python explore_all.py 'search_term'")
        print("\nExamples:")
        print("   python explore_all.py 'assign'")
        print("   python explore_all.py 'date'")
        print("   python explore_all.py 'location'")
        print("=" * 80)

