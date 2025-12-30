"""
Data Loader for SAP Reports
Handles loading and inital processing of Excel Files
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
from Config import Config

class SAPDataLoader:
    """Loads SAP report files and provides access"""

    def __init__(self):
        self.config = Config()
        self.reports = {}
        self.load_timestamp = None

    def load_all_reports(self, verbose=True):
        """
        Load all reports from the config

        Returns:
        dict: Dictionary of DataFrames with report keys
        """
        print("=" * 60)
        print("SAP REPORT DATA LOADER")
        print("=" * 60)
        print(f"Loading reports from: {self.config.BASE_FOLDER}\n")

        file_paths = self.config.get_all_file_paths()

        for report_key, file_path in file_paths.items():
            df = self._load_single_report(file_path, report_key, verbose)
            if df is not None:
                self.reports[report_key] = df

        self.load_timestamp = datetime.now()

        print("\n" + "=" * 60)
        print(f"Successfully loaded {len(self.reports)}/{len(file_paths)} reports")
        print(f"Last load: {self.load_timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 60)

        return self.reports
        
    def _load_single_report(self, file_path, report_key, verbose=True):
        """Load a single Excel report file"""

        if verbose:
            print(f"\nLoading {report_key}")
            print(f"From: {file_path}")

        if not file_path.exists():
            print(f"Error: File not found!")
            print(f"Path: {file_path}")
            return None

        try:
            #load the Excel file
            df = pd.read_excel(
                file_path,
                engine=self.config.EXCEL_ENGINE
            )

            #Clean up the column names (reomve extra spaces, etc.)
            df.columns = df.columns.str.strip()

            #Display info
            if verbose:
                print(f"Loaded: {len(df):,} rows x {len(df.columns)} columns")
                print(f"Columns: {', '.join(df.columns[:5].tolist())}" + 
                        (f"... (+{len(df.columns) - 5} more)" if len(df.columns) > 5 else ""))

            return df
        
        except Exception as e:
            print(f"Error loading {report_key}: {e}")
            return None
        
    def get_report(self, report_key):
        """Get a specific loaded report"""
        return self.reports.get(report_key)

    def explore_report(self, report_key):
        """
        Display detailed information about a specific report
        """

        if report_key not in self.reports:
            print(f"Error: Report '{report_key}' not loaded")
            return

        df = self.get_report(report_key)
        print(f"REPORT EXPLORER: {report_key.upper()}")
        print("=" * 60)

        # Basic info
        print(f"\n Dataset Shape: {df.shape[0]:,} rows x {df.shape[1]:,} columns")

        # Column details
        print(f"\n All Columns ({len(df.columns)}):")
        for i, col in enumerate(df.columns, 1):
            non_null = df[col].notna().sum()
            null_pct = (df[col].isna().sum() / len(df)) * 100
            dtype = df[col].dtype
            print(f" {i:2d}, {col:40s} [{dtype}] - {non_null:,} values ({null_pct:.1f}% null)")

        # Sample data
        print(f"\n First 3 Rows:")
        print(df.head(3).to_string())

        # Data types summary
        print(f"\n Data Types:")
        print(df.dtypes.value_counts())

        print("\n" + "=" * 60)

    def find_your_assignments(self, name_column=None, user_name=None):
        """
        Find all rows in cost estimating schedule for a specific user

        Args:
        name_column: Column name search (will auto-detect if None)
        user_name: Name to search for (Uses Config.USER_NAME if None)
        """

        if 'cost estimating' not in self.reports:
            print("Cost Estimating Schedule not loaded!")
            return None

        df = self.reports['cost estimating']
        user_name = user_name or self.config.USER_NAME

        # If no column name specified, try to find it
        if name_column is None:
            possible_cols = [col for col in df.columns
                            if any(keyword in col.lower()
                                for keyword in ['assign', 'estimator'])]
            if possible_cols:
                print(f"Found possible name column: {possible_cols}")
                print(f" Available columns: {', '.join(df.columns)}")
                return None
        
        # Filter for the user
        mask = df[name_column].astype(str).str.contains(user_name, case=False, na=False)
        user_projects = df[mask]

        print(f"\n Found {len(user_projects):,} projects assigned to {user_name}")

        return user_projects

    def get_summary_stats(self):
        """Get summary statistics for all loaded reports"""
        print("=" * 60)
        print("SUMMARY STATISTICS")
        print("=" * 60)

        for report_key, df in self.reports.items():
            print(f"\n Report: {report_key.upper()}:")
            print(f" * Total Rows: {len(df):,}")

            # Try to find data columns
            date_cols = [col for col in df.columns if 'data' in col.lower()]
            if date_cols:
                print(f" * Date Columns: {', '.join(date_cols)}")

def main():
    """Main function to testthe data loader"""

    # Initialize the data loader
    loader = SAPDataLoader()

    # Load all reports
    reports = loader.load_all_reports(verbose=True)

    # Show summary
    loader.get_summary_stats()

    # Explore the main file (SD-09 Cost Estimating Schedule)
    print("\n\nExploring Cost Estimating Schedule...")
    loader.explore_report('cost estimating')

    # Try to find user assignments
    print("\n\nFinding your assignments...")
    user_projects = loader.find_your_assignments()

    if user_projects is not None and len(user_projects) > 0:
        print("\n Your Active Projects:")
        print(user_projects.head(10).to_string())

    return loader

if __name__ == "__main__":
    loader = main()