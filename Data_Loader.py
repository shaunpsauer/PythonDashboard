"""
Data Loader for SAP Reports
Handles loading and initial processing of Excel files
"""
import pandas as pd
from pathlib import Path
from datetime import datetime
from Config import Config


class SAPDataLoader:
    """Loads SAP report files and provides data access"""
    
    def __init__(self):
        self.config = Config()
        self.reports = {}
        self.load_timestamp = None
        
    def load_all_reports(self, verbose=True):
        """
        Load all configured SAP reports
        
        Returns:
            dict: Dictionary of DataFrames with report keys
        """
        print("=" * 60)
        print("SAP REPORT DATA LOADER")
        print("=" * 60)
        print(f"Loading reports from: {self.config.BASE_FOLDER}\n")
        
        file_paths = self.config.get_all_file_paths()
        
        for report_key, file_path in file_paths.items():
            df = self._load_single_report(report_key, file_path, verbose)
            if df is not None:
                self.reports[report_key] = df
        
        self.load_timestamp = datetime.now()
        
        print("\n" + "=" * 60)
        print(f"Successfully loaded {len(self.reports)}/{len(file_paths)} reports")
        print(f"Load completed at: {self.load_timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 60)
        
        return self.reports
    
    def _find_header_row(self, file_path, max_rows_to_check=20):
        """
        Automatically detect which row contains the column headers
        
        Strategies:
        1. Look for rows with high column fill rate (many non-empty cells)
        2. Look for common header keywords
        3. Avoid rows that are all the same value
        
        Returns:
            int: Row number (0-indexed) where headers are found, or 0 if not detected
        """
        # Load first N rows without assuming header position
        df_sample = pd.read_excel(
            file_path, 
            header=None,
            nrows=max_rows_to_check,
            engine=self.config.EXCEL_ENGINE
        )
        
        # Common header keywords
        header_keywords = [
            'project', 'name', 'date', 'status', 'phase', 'number',
            'assigned', 'estimator', 'location', 'region', 'pmo',
            'start', 'end', 'complete', 'due', 'schedule', 'id'
        ]
        
        best_row = 0
        best_score = 0
        
        for idx in range(min(max_rows_to_check, len(df_sample))):
            row = df_sample.iloc[idx]
            
            # Calculate score for this row being the header
            score = 0
            
            # Factor 1: How many cells are filled? (Headers usually have many columns)
            filled_ratio = row.notna().sum() / len(row)
            score += filled_ratio * 100
            
            # Factor 2: Are values text/strings? (Headers are usually text)
            text_values = sum(1 for val in row if isinstance(val, str))
            score += (text_values / len(row)) * 50
            
            # Factor 3: Contains header keywords?
            if row.notna().any():
                row_text = ' '.join([str(val).lower() for val in row if pd.notna(val)])
                keyword_matches = sum(1 for keyword in header_keywords if keyword in row_text)
                score += keyword_matches * 30
            
            # Factor 4: Not all the same value (avoid title rows like "PROJECT REPORT")
            unique_values = row.dropna().nunique()
            if unique_values > 3:  # Good header has multiple different column names
                score += 20
            elif unique_values == 1:  # Probably a title row
                score -= 50
            
            # Track best scoring row
            if score > best_score:
                best_score = score
                best_row = idx
        
        return best_row
    
    def _load_single_report(self, report_key, file_path, verbose=True):
        """Load a single Excel report file"""
        
        if verbose:
            print(f"\nLoading: {report_key}")
            print(f"   File: {file_path.name}")
        
        # Check if file exists
        if not file_path.exists():
            print(f"   ERROR: File not found!")
            print(f"   Path: {file_path}")
            return None
        
        try:
            # Step 1: Find the header row automatically
            header_row = self._find_header_row(file_path)
            
            if verbose and header_row > 0:
                print(f"   Auto-detected header row: {header_row + 1} (skipping {header_row} rows)")
            
            # Step 2: Load Excel file with correct header position
            df = pd.read_excel(
                file_path,
                header=header_row,
                engine=self.config.EXCEL_ENGINE
            )
            
            # Step 3: Clean up column names
            # Remove extra spaces
            df.columns = df.columns.astype(str).str.strip()
            
            # Remove any "Unnamed: X" columns that are completely empty
            unnamed_cols = [col for col in df.columns if 'Unnamed' in str(col)]
            for col in unnamed_cols:
                if df[col].isna().all():
                    df = df.drop(columns=[col])
            
            # Step 4: Remove any completely empty rows (sometimes after header)
            df = df.dropna(how='all')
            
            # Reset index after dropping rows
            df = df.reset_index(drop=True)
            
            # Display info
            if verbose:
                print(f"   Loaded: {len(df):,} rows x {len(df.columns)} columns")
                print(f"   Columns: {', '.join(df.columns[:5].tolist())}" + 
                      (f"... (+{len(df.columns)-5} more)" if len(df.columns) > 5 else ""))
            
            return df
            
        except Exception as e:
            print(f"   ERROR loading file: {str(e)}")
            return None
    
    def get_report(self, report_key):
        """Get a specific loaded report"""
        return self.reports.get(report_key)
    
    def explore_report(self, report_key):
        """
        Display detailed information about a specific report
        """
        if report_key not in self.reports:
            print(f"ERROR: Report '{report_key}' not loaded")
            return
        
        df = self.reports[report_key]
        
        print("\n" + "=" * 60)
        print(f"REPORT EXPLORER: {report_key.upper()}")
        print("=" * 60)
        
        # Basic info
        print(f"\nDataset Shape: {df.shape[0]:,} rows x {df.shape[1]} columns")
        
        # Column details
        print(f"\nAll Columns ({len(df.columns)}):")
        for i, col in enumerate(df.columns, 1):
            non_null = df[col].notna().sum()
            null_pct = (df[col].isna().sum() / len(df) * 100)
            dtype = df[col].dtype
            print(f"   {i:2d}. {col:40s} [{dtype}] - {non_null:,} values ({null_pct:.1f}% null)")
        
        # Sample data
        print(f"\nFirst 3 Rows:")
        print(df.head(3).to_string())
        
        # Data types summary
        print(f"\nData Types:")
        print(df.dtypes.value_counts())
        
        print("=" * 60)
    
    def find_your_assignments(self, name_column=None, user_name=None):
        """
        Find all rows in cost_estimating report assigned to you
        
        Args:
            name_column: Column name to search (will auto-detect if None)
            user_name: Your name to search for (uses Config.USER_NAME if None)
        """
        if 'cost_estimating' not in self.reports:
            print("ERROR: Cost Estimating Schedule not loaded")
            return None
        
        df = self.reports['cost_estimating']
        user_name = user_name or self.config.USER_NAME
        
        # If no column specified, try to find it
        if name_column is None:
            possible_cols = [col for col in df.columns 
                           if any(keyword in col.lower() 
                                 for keyword in ['assign', 'estimator', 'name', 'owner'])]
            
            if possible_cols:
                print(f"Found possible name columns: {possible_cols}")
                print(f"   Using: {possible_cols[0]}")
                name_column = possible_cols[0]
            else:
                print("ERROR: Could not auto-detect name column")
                print(f"   Available columns: {', '.join(df.columns)}")
                return None
        
        # Normalize names for flexible matching (handles order, case, punctuation)
        def normalize_name(name):
            """Normalize a name by lowercasing, splitting on delimiters, and sorting words"""
            import re
            # Convert to lowercase and split on common delimiters (comma, space, etc.)
            words = re.split(r'[,;\s]+', str(name).lower())
            # Remove empty strings and sort for consistent comparison
            words = sorted([w.strip() for w in words if w.strip()])
            return ' '.join(words)
        
        # Normalize the user name
        normalized_user_name = normalize_name(user_name)
        
        # Apply normalization to each value in the column and check for match
        mask = df[name_column].astype(str).apply(
            lambda x: normalize_name(x) == normalized_user_name
        )
        your_projects = df[mask]
        
        print(f"\nFound {len(your_projects)} projects assigned to {user_name}")
        
        return your_projects
    
    def get_summary_stats(self):
        """Get summary statistics for all loaded reports"""
        print("\n" + "=" * 60)
        print("SUMMARY STATISTICS")
        print("=" * 60)
        
        for report_key, df in self.reports.items():
            print(f"\n{report_key.upper()}:")
            print(f"  Total rows: {len(df):,}")
            print(f"  Columns: {len(df.columns)}")
            
            # Try to find date columns
            date_cols = [col for col in df.columns if 'date' in col.lower()]
            if date_cols:
                print(f"  Date columns: {', '.join(date_cols)}")


def main():
    """Main function to test the data loader"""
    
    # Initialize loader
    loader = SAPDataLoader()
    
    # Load all reports
    reports = loader.load_all_reports(verbose=True)
    
    # Show summary
    loader.get_summary_stats()
    
    # Explore the main file (Cost Estimating Schedule)
    print("\n\nEXPLORING PRIMARY FILE: COST ESTIMATING SCHEDULE")
    loader.explore_report('cost_estimating')
    
    # Try to find your assignments
    print("\n\nFINDING YOUR ASSIGNMENTS:")
    your_projects = loader.find_your_assignments()
    
    if your_projects is not None and len(your_projects) > 0:
        print("\nYour Active Projects:")
        print(your_projects.head(10).to_string())
    
    return loader


if __name__ == "__main__":
    loader = main()
