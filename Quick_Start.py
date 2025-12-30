"""
Quick Start Script - Test Your Data Loading
Run this after updating the file path in Config.py
"""
from Data_Loader import SAPDataLoader
from Config import Config


def quick_test():
    """Quick test to verify everything works"""
    
    print("PROJECT TRACKER - QUICK START TEST")
    print("\n1. Checking configuration...")
    
    # Show configured paths
    print(f"   Base folder: {Config.BASE_FOLDER}")
    print(f"   Your name: {Config.USER_NAME}")
    print(f"\n   Files to load:")
    for key, filename in Config.REPORT_FILES.items():
        print(f"      - {key}: {filename}")
    
    input("\n   Press ENTER to continue with data loading...")
    
    # Load data
    print("\n2. Loading SAP reports...")
    loader = SAPDataLoader()
    reports = loader.load_all_reports(verbose=True)
    
    if len(reports) == 0:
        print("\nERROR: No files loaded successfully!")
        print("   Check that your file paths in Config.py are correct")
        return
    
    # Show what we got
    print("\n3. Exploring loaded data...")
    
    # Show the main file (Cost Estimating)
    if 'cost_estimating' in reports:
        print("\nCOST ESTIMATING SCHEDULE (Your main file):")
        loader.explore_report('cost_estimating')
        
        print("\n4. Finding your assignments...")
        your_projects = loader.find_your_assignments()
        
        if your_projects is not None and len(your_projects) > 0:
            print(f"\nSUCCESS! Found {len(your_projects)} projects assigned to you")
            print("\nFirst few projects:")
            print(your_projects.head().to_string())
    
    print("\n" + "=" * 60)
    print("TEST COMPLETE!")
    print("\nNext steps:")
    print("   1. Check that the column names look correct above")
    print("   2. Verify your assignments were found correctly")
    print("   3. Ready to build the dashboard!")
    print("=" * 60)
    
    return loader


if __name__ == "__main__":
    loader = quick_test()
