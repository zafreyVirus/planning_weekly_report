# planning_report.py
import pandas as pd
import os
import glob
from datetime import datetime, timedelta
from pathlib import Path

# ============================================
# CONFIGURATION
# ============================================

# Get the current script's directory
BASE_DIR = Path(__file__).parent

# Folder configuration
CSV_FOLDER = BASE_DIR / "data"
OUTPUT_FOLDER = BASE_DIR / "output"
LOG_FOLDER = BASE_DIR / "logs"

# Create folders if they don't exist
CSV_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)
LOG_FOLDER.mkdir(exist_ok=True)

OUTPUT_EXCEL = OUTPUT_FOLDER / "planning_report_output.xlsx"

# File patterns - ADJUST THESE TO MATCH YOUR ACTUAL FILE NAMES
FILE_PATTERNS = {
    # SAU (Attached Users) files
    "2G_SAU_LMB": "*2G*SAU*LMB*.csv",
    "2G_SAU_LL": "*2G*SAU*LLG*.csv",
    "3G_SAU_LMB": "*3G*SAU*LMB*.csv",
    "3G_SAU_LL": "*3G*SAU*LLG*.csv",
    "4G_SAU_LMB": "*4G*SAU*LMB*.csv",
    "4G_SAU_LL": "*4G*SAU*LLG*.csv",
    # PDP files
    "2G_PDP_LMB": "*2G*PDP*LMB*.csv",
    "2G_PDP_LL": "*2G*PDP*LLG*.csv",
    "3G_PDP_LMB": "*3G*PDP*LMB*.csv",
    "3G_PDP_LL": "*3G*PDP*LLG*.csv",
    "4G_PDP_LMB": "*4G*PDP*LMB*.csv",
    "4G_PDP_LL": "*4G*PDP*LLG*.csv",
}

# Metric column mappings
METRIC_COLUMNS = {
    "2G_SAU": "Gb mode maximum attached users (number)",
    "3G_SAU": "Iu mode maximum attached users (number)",
    "4G_SAU": "Maximum attached users (number)",
    "2G_PDP": "Gb mode maximum act PDP context (number)",
    "3G_PDP": "Iu mode maximum active PDP context (number)",
    "4G_PDP": "Maximum bearer number in S1 mode (number)",
}

# NE Name mapping
NE_NAMES = {
    "LMB": "LMB_vUSN01",
    "LL": "CLOUDUSN",
}


# ============================================
# FUNCTIONS
# ============================================

def setup_logging():
    """Setup logging configuration"""
    import logging
    log_file = LOG_FOLDER / f"processing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)


def find_data_start_row(file_path):
    """Find the row index where the actual data header starts."""
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
            for i, line in enumerate(lines):
                if 'Start Time' in line and 'Period (min)' in line:
                    return i
        return 0
    except Exception as e:
        logger.error(f"Error reading {file_path}: {e}")
        return 0


def read_csv_with_metadata(file_path):
    """Read CSV file that has metadata rows at the top."""
    start_row = find_data_start_row(file_path)
    df = pd.read_csv(file_path, skiprows=start_row, encoding='utf-8-sig')
    df.columns = df.columns.str.strip('"')
    return df


def get_max_from_last_7_days(df, ne_name, metric_column):
    """Get the maximum value from the last 7 days of data."""
    df_filtered = df[df['NE Name'] == ne_name]
    
    if df_filtered.empty:
        logger.warning(f"No data found for NE Name '{ne_name}'")
        return 0
    
    df_filtered = df_filtered.copy()
    df_filtered['Start Time'] = pd.to_datetime(df_filtered['Start Time'])
    
    max_date = df_filtered['Start Time'].max()
    cutoff_date = max_date - timedelta(days=7)
    df_last_7d = df_filtered[df_filtered['Start Time'] >= cutoff_date]
    
    if df_last_7d.empty:
        logger.warning(f"No data in last 7 days for '{ne_name}'")
        return 0
    
    return df_last_7d[metric_column].max()


def find_file(file_pattern):
    """Find the first file matching the pattern in the CSV folder."""
    pattern = str(CSV_FOLDER / file_pattern)
    files = glob.glob(pattern)
    
    if files:
        logger.info(f"Found file: {os.path.basename(files[0])}")
        return files[0]
    
    logger.warning(f"No file found for pattern: {file_pattern}")
    return None


def get_cloud_usn_value(file_path, region, metric_type):
    """Get the Cloud USN value for a specific region and metric type."""
    if not file_path:
        return 0
    
    try:
        df = read_csv_with_metadata(file_path)
        ne_name = NE_NAMES[region]
        metric_column = METRIC_COLUMNS[metric_type]
        value = get_max_from_last_7_days(df, ne_name, metric_column)
        return value if pd.notna(value) else 0
    except Exception as e:
        logger.error(f"Error processing {file_path}: {e}")
        return 0


def generate_calculation_sheet():
    """Generate the Calculation sheet with Cloud USN data."""
    
    results = {'LMB': {}, 'LL': {}}
    metrics = ['2G_SAU', '3G_SAU', '4G_SAU', '2G_PDP', '3G_PDP', '4G_PDP']
    
    logger.info("Starting Cloud USN data extraction...")
    logger.info(f"Looking for CSV files in: {CSV_FOLDER}")
    
    # Check if data folder has any CSV files
    csv_files = list(CSV_FOLDER.glob("*.csv"))
    if not csv_files:
        logger.error(f"No CSV files found in {CSV_FOLDER}")
        logger.info(f"Please place your CSV files in the '{CSV_FOLDER}' folder")
        return None
    
    logger.info(f"Found {len(csv_files)} CSV files in data folder")
    
    for region in ['LMB', 'LL']:
        logger.info(f"\nProcessing {region} Cloud USN...")
        
        for metric in metrics:
            pattern_key = f"{metric}_{region}"
            
            if pattern_key in FILE_PATTERNS:
                file_pattern = FILE_PATTERNS[pattern_key]
                file_path = find_file(file_pattern)
                value = get_cloud_usn_value(file_path, region, metric)
                results[region][metric] = value
                logger.info(f"  {metric}: {value:,.0f}")
            else:
                logger.warning(f"  {metric}: No file pattern defined")
                results[region][metric] = 0
    
    # Build DataFrame
    calculation_data = []
    
    for region in ['LMB', 'LL']:
        for metric, display_name in [
            ('2G_SAU', '2G SAU (per K sub)'),
            ('3G_SAU', '3G SAU (per K sub)'),
            ('4G_SAU', '4G SAU (per K sub)'),
            ('2G_PDP', '2G PDP (per K sub)'),
            ('3G_PDP', '3G PDP (per K sub)'),
            ('4G_PDP', '4G PDP (per K sub)')
        ]:
            calculation_data.append({
                'NE Type': 'Cloud USN',
                'Region': region,
                'Metric': display_name,
                'Value': results[region][metric],
                'Unit': 'count',
                'Data Source': 'Max from last 7 days'
            })
    
    return pd.DataFrame(calculation_data)


def generate_summary_sheet(df_calculation):
    """Generate a summary/pivot sheet for easier reading."""
    
    # Pivot the data for better readability
    pivot_df = df_calculation.pivot_table(
        index='Metric',
        columns='Region',
        values='Value',
        aggfunc='first'
    ).reset_index()
    
    # Add a total column if needed
    if 'LMB' in pivot_df.columns and 'LL' in pivot_df.columns:
        pivot_df['Total'] = pivot_df['LMB'] + pivot_df['LL']
    
    return pivot_df


# ============================================
# MAIN SCRIPT
# ============================================

if __name__ == "__main__":
    
    # Setup logging
    logger = setup_logging()
    
    print("=" * 60)
    print("CLOUD USN PLANNING REPORT GENERATOR")
    print("=" * 60)
    print(f"\nBase Directory: {BASE_DIR}")
    print(f"CSV Folder: {CSV_FOLDER}")
    print(f"Output Folder: {OUTPUT_FOLDER}")
    print(f"Logs Folder: {LOG_FOLDER}")
    
    # Check if data folder has CSV files
    if not any(CSV_FOLDER.glob("*.csv")):
        print("\n⚠️  WARNING: No CSV files found in 'data' folder!")
        print(f"\nPlease place your CSV files in: {CSV_FOLDER}")
        print("\nExpected file patterns (examples):")
        print("  - USN_SAU_2G_LMB_20260602.csv")
        print("  - USN_SAU_2G_LLG_20260602.csv")
        print("  - USN_SAU_3G_LMB_20260602.csv")
        print("  - USN_SAU_3G_LLG_20260602.csv")
        print("  - USN_SAU_4G_LMB_20260602.csv")
        print("  - USN_SAU_4G_LLG_20260602.csv")
        print("  - USN_2G_PDP_LMB_20260602.csv")
        print("  - USN_2G_PDP_LLG_20260602.csv")
        print("  - USN_3G_PDP_LMB_20260602.csv")
        print("  - USN_3G_PDP_LLG_20260602.csv")
        print("  - USN_4G_PDP_LMB_20260602.csv")
        print("  - USN_4G_PDP_LLG_20260602.csv")
        print("\n⚠️  Please add your CSV files and run again.")
        exit(1)
    
    print("\n" + "-" * 60)
    
    # Generate the Calculation sheet
    df_calculation = generate_calculation_sheet()
    
    if df_calculation is not None and not df_calculation.empty:
        # Generate summary sheet
        df_summary = generate_summary_sheet(df_calculation)
        
        # Write to Excel
        with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
            df_calculation.to_excel(writer, sheet_name='Calculation', index=False)
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
        
        print("\n" + "=" * 60)
        print(f"✅ SUCCESS! Report generated successfully!")
        print(f"📁 Output file: {OUTPUT_EXCEL}")
        print("=" * 60)
        
        # Print summary
        print("\n📊 SUMMARY OF EXTRACTED DATA:")
        print("-" * 60)
        print(df_summary.to_string(index=False))
        
    else:
        print("\n❌ ERROR: Failed to generate calculation sheet")
        print("Check the logs for more details.")