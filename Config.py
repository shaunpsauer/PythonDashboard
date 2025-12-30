"""
Configuration for SAP Report Data Ingestion
"""

from pathlib import Path

class Config:
    """Configuration settings for the scheduler"""

    #Base folder path
    BASE_FOLDER = r"C:\Users\Shaun\OneDrive - PGE\00 Program (All Workstreams)"

    #Key report files
    REPORT_FILES = {
        'cost_estimating': 'sd-09 Cost Estimating Schedule.xlsx',
        'milestone': 'sd-01 Milestone Schedule.xlsx',
        'contract': 'sd-02 Contract Schedule.xlsx',
        'order_data': 'sd-17 PGE Gas Ops Order Data.xlsx',
    }

    #Excel read settings
    EXCEL_ENGINE = 'openpyxl' #for xlsx files

    #Data storage
    DATABASE_PATH = Path(__file__).parent / 'data' / 'project_tracker.db'
    ARCHIVE_FOLDER = Path(__file__).parent / 'data' / 'archive'

    #User Name
    USER_NAME = 'Shaun Sauer'

    @classmethod
    def get_file_path(cls, report_key):
        """Get the full path to a report file"""
        return Path(cls.BASE_FOLDER) / cls.REPORT_FILES[report_key]

    @classmethod
    def get_all_file_paths(cls):
        """Get the full paths to all report files"""
        return {
            key: Path(cls.BASE_FOLDER) / file_name 
            for key, file_name in cls.REPORT_FILES.items()
            }