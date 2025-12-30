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
        'Cost Estimating': 'sd-09 Cost Estimating Schedule.xlsx',
        'Milestone Schedule': 'sd-01 Milestone Schedule.xlsx',
        'Contract Schedule': "sd-01 Contract Schedule.xlsx",
        'Order Data': 'sd-17 PGE Gas Ops Order Data.xlsx',
    }

    #Excel read settings
    EXCEL_ENGINE = 'openpyxl' #for xlsx files

    #Data storage
    DATABASE_PATH = Path(BASE_FOLDER) / 'Data' / 'project_tracker.db'
    ARCHIVE_FOLDER = Path(BASE_FOLDER) / 'Data' / 'Archive'

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