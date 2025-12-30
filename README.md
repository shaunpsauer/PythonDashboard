# SAP Project Tracker - Data Ingestion Module

## Key Features

- Automatic Header Detection - Finds column headers even when they're not on row 1
- Smart Data Loading - Handles SAP export formatting quirks
- Assignment Discovery - Automatically finds your projects
- Flexible Configuration - Easy to add more files

## Quick Setup

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Update Configuration
Edit `Config.py` and set your actual file path:

```python
BASE_FOLDER = r"C:\Users\Shaun\OneDrive - PGE\00 Program (All Workstreams)"
```

### 3. Run Quick Test
```bash
python Quick_Start.py
```

This will:
- Load all 4 SAP report files
- Show you what columns exist in each file
- Find your assigned projects in sd-09 Cost Estimating Schedule
- Display sample data

## Files Overview

### Core Files
- **Config.py** - Configuration settings (UPDATE THIS FIRST!)
- **Data_Loader.py** - Main data loading logic
- **Quick_Start.py** - Test script to verify setup

### Utility Files
- **diagnose_excel.py** - Diagnostic tool to inspect Excel file structure
- **explore_all.py** - Explore all Excel files in the folder

### Data Files Being Loaded
1. **sd-01 Milestone Schedule.xlsx** - Key project milestones
2. **sd-02 Contract Schedule.xlsx** - Contract information
3. **sd-09 Cost Estimating Schedule.xlsx** - YOUR MAIN FILE (assignments)
4. **sd-17 PGE Gas Ops Order Data.xlsx** - Order data

## Usage Examples

### Basic Loading
```python
from Data_Loader import SAPDataLoader

# Initialize and load
loader = SAPDataLoader()
reports = loader.load_all_reports()

# Access specific report
cost_estimating_df = loader.get_report('cost_estimating')
milestone_df = loader.get_report('milestone')
```

### Find Your Assignments
```python
# Auto-detect column and find your projects
your_projects = loader.find_your_assignments()

# Or specify the exact column name
your_projects = loader.find_your_assignments(
    name_column='Assigned Estimator',
    user_name='Shaun'
)
```

### Explore Data Structure
```python
# See detailed info about any report
loader.explore_report('cost_estimating')
loader.explore_report('milestone')
```

## What's Next?

After verifying data loads correctly:
1. Week 1 Complete: Data ingestion working
2. Week 2: Build change detection (compare week-to-week)
3. Week 3: Create Streamlit dashboard

## Troubleshooting

**File not found errors?**
- Check that `BASE_FOLDER` path in `Config.py` is correct
- Make sure OneDrive is synced and files are available locally

**Wrong columns detected / Headers look weird?**
- Run the diagnostic tool: `python diagnose_excel.py`
- This shows you exactly where headers are being detected
- Choose option 1 to test all configured files
- If detection is wrong, you can manually specify header row in config

**Can't find assignments?**
- Run `loader.explore_report('cost_estimating')` to see all columns
- Update `USER_NAME` in Config.py if needed
- Manually specify column name in `find_your_assignments()`

**Wrong data loaded?**
- Check that files are up to date (should be from Friday)
- Verify file names match exactly in `Config.py`

### Using the Diagnostic Tool

The diagnostic tool helps you understand your Excel file structure:

```bash
# Test header detection on all files
python diagnose_excel.py
# Then select option 1

# Inspect a specific file in detail
python diagnose_excel.py
# Then select option 2 and enter filename
```

This will show you:
- Raw contents of first 10 rows
- Analysis of which row is likely the header
- Detection confidence scores
- Actual column names found

