import pandas as pd
import os
from datetime import datetime, timedelta

# ============================================
# CONFIGURATION
# ============================================

CSV_FOLDER = "data/"
OUTPUT_EXCEL = "planning_report_output.xlsx"

# NE Name mapping
NE_NAMES = {
    "LMB": {
        "USN": "LMB_vUSN01",
        "CGW": "LMB_vCGW01",
        "DGW": "LMB_vDGW01",
        "UPCF": "dc2_LMBUPCF01",
        "MGW": "NEW LMBMGW",
        "VMGW": "LMBVMGW01"
    },
    "LL": {
        "USN": "CLOUDUSN",
        "CGW": "LLG_vCGW01",
        "DGW": "LLG_vDGW01",
        "UPCF": "dc1_LLGUPCF01",
        "MGW": "NEW LLGMGW",
        "VMGW": "LLWVMGW01"
    }
}

# File patterns
FILE_PATTERNS = {
    # Cloud USN files - SAU
    "2G_SAU_LMB": "USN SAU 2G LMB.csv",
    "2G_SAU_LL": "USN SAU 2G LLG.csv",
    "3G_SAU_LMB": "USN SAU 3G LMB.csv",
    "3G_SAU_LL": "USN SAU 3G LLG.csv",
    "4G_SAU_LMB": "USN SAU 4G LMB.csv",
    "4G_SAU_LL": "USN SAU 4G LLG.csv",
    # Cloud USN files - PDP
    "2G_PDP_LMB": "USN 2G PDP LMB.csv",
    "2G_PDP_LL": "USN 2G PDP LLG.csv",
    "3G_PDP_LMB": "vUSN 3G PDP LMB.csv",
    "3G_PDP_LL": "vUSN 3G PDP LLG.csv",
    "4G_PDP_LMB": "USN 4G PDP LMB.csv",
    "4G_PDP_LL": "USN 4G PDP LLG.csv",
    # Cloud UGW files
    "CGW_2G3G_PDP": "CGW 2G3G PDP.csv",
    "CGW_4G_PDP": "CGW 4G PDP.csv",
    "DGW_THROUGHPUT_3G": "vDGW Throughput 3G.csv",
    "DGW_THROUGHPUT_4G": "vDGW Throughput 4G.csv",
    # UPCF files
    "UPCF_LMB": "UPCF PDP LMB.csv",
    "UPCF_LL": "UPCF PDP LLG.csv",
    # MGW files
    "MGW": "MGW Traffic.csv",
    "VMGW": "vMGW Traffic.csv",
}

# Metric column mappings for Cloud USN
USN_METRIC_COLUMNS = {
    "2G_SAU": "Gb mode maximum attached users (number)",
    "3G_SAU": "Iu mode maximum attached users (number)",
    "4G_SAU": "Maximum attached users (number)",
    "2G_PDP": "Gb mode maximum act PDP context (number)",
    "3G_PDP": "Iu mode maximum active PDP context (number)",
    "4G_PDP": "Maximum bearer number in S1 mode (number)",
}

# Metric column mappings for Cloud UGW
UGW_METRIC_COLUMNS = {
    "2G3G_PDP": "PGW-C 2/3G Maximum simultaneously activated PDP contexts (number)",
    "4G_PDP": "SGW-C and PGW-C combined Max number of Active EPS Bearers (number)",
    "THROUGHPUT_3G": "PGW-U 2/3G Gi downlink peak throughput in MB/s (MB/s)",
    "THROUGHPUT_4G": "User Plane SGi downlink user traffic peak throughput in MB/s (MB/s) (MB/s)",
}

# UPCF metric column
UPCF_METRIC_COLUMN = "Maximum Number of Active Subscribers (number)"

# MGW metric column
MGW_METRIC_COLUMN = "Peak Licensed Traffic (number)"

# vMGW metric column
VMGW_METRIC_COLUMN = "License Resource Usage for Basic Software on the MGW (%)"


# ============================================
# HELPER FUNCTIONS
# ============================================

def find_data_start_row(file_path):
    """Find the row index where the actual data header starts."""
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        lines = f.readlines()
        for i, line in enumerate(lines):
            if 'Start Time' in line and 'Period (min)' in line:
                return i
    return 0


def read_csv_with_metadata(file_path):
    """Read CSV file that has metadata rows at the top."""
    try:
        start_row = find_data_start_row(file_path)
        df = pd.read_csv(file_path, skiprows=start_row, encoding='utf-8-sig')
        df.columns = df.columns.str.strip('"')
        return df
    except Exception as e:
        print(f"  Error reading {file_path}: {e}")
        return None


def get_max_from_last_7_days(df, ne_name, metric_column):
    """Get the maximum value from the last 7 days for a specific NE."""
    if df is None or df.empty:
        return 0
    
    # Filter by NE Name
    df_filtered = df[df['NE Name'] == ne_name]
    
    if df_filtered.empty:
        print(f"  Warning: No data found for NE Name '{ne_name}'")
        return 0
    
    # Convert Start Time to datetime
    df_filtered['Start Time'] = pd.to_datetime(df_filtered['Start Time'])
    
    # Get the last date available in the file
    max_date = df_filtered['Start Time'].max()
    cutoff_date = max_date - timedelta(days=7)
    
    # Filter to last 7 days
    df_last_7d = df_filtered[df_filtered['Start Time'] >= cutoff_date]
    
    if df_last_7d.empty:
        print(f"  Warning: No data in last 7 days for '{ne_name}'")
        return 0
    
    # Return the maximum value
    return df_last_7d[metric_column].max()


def get_cloud_usn_value(region, metric_type):
    """Get Cloud USN value for a specific region and metric."""
    file_pattern_key = f"{metric_type}_{region}"
    
    if file_pattern_key not in FILE_PATTERNS:
        return 0
    
    file_name = FILE_PATTERNS[file_pattern_key]
    file_path = os.path.join(CSV_FOLDER, file_name)
    
    if not os.path.exists(file_path):
        return 0
    
    try:
        df = read_csv_with_metadata(file_path)
        if df is None:
            return 0
        
        ne_name = NE_NAMES[region]["USN"]
        metric_column = USN_METRIC_COLUMNS[metric_type]
        return get_max_from_last_7_days(df, ne_name, metric_column)
    except Exception as e:
        print(f"  Error processing {file_name}: {e}")
        return 0


# ============================================
# CLOUD UGW FUNCTIONS
# ============================================

def get_cgw_2g3g_pdp(region):
    """Get vCGW 2G/3G PDP max value from last 7 days."""
    file_name = FILE_PATTERNS["CGW_2G3G_PDP"]
    file_path = os.path.join(CSV_FOLDER, file_name)
    
    if not os.path.exists(file_path):
        return 0
    
    try:
        df = read_csv_with_metadata(file_path)
        if df is None:
            return 0
        
        ne_name = NE_NAMES[region]["CGW"]
        metric_col = UGW_METRIC_COLUMNS["2G3G_PDP"]
        return get_max_from_last_7_days(df, ne_name, metric_col)
    except Exception as e:
        print(f"  Error processing {file_name}: {e}")
        return 0


def get_cgw_4g_pdp(region):
    """Get vCGW 4G PDP max value from last 7 days."""
    file_name = FILE_PATTERNS["CGW_4G_PDP"]
    file_path = os.path.join(CSV_FOLDER, file_name)
    
    if not os.path.exists(file_path):
        return 0
    
    try:
        df = read_csv_with_metadata(file_path)
        if df is None:
            return 0
        
        ne_name = NE_NAMES[region]["CGW"]
        metric_col = UGW_METRIC_COLUMNS["4G_PDP"]
        return get_max_from_last_7_days(df, ne_name, metric_col)
    except Exception as e:
        print(f"  Error processing {file_name}: {e}")
        return 0


def get_cgw_throughput(region):
    """Get vCGW Throughput = Gi (3G) + SGi (4G) throughput in MB/s."""
    
    # Get 3G Gi throughput
    file_3g_name = FILE_PATTERNS["DGW_THROUGHPUT_3G"]
    file_3g_path = os.path.join(CSV_FOLDER, file_3g_name)
    
    throughput_3g = 0
    if os.path.exists(file_3g_path):
        try:
            df_3g = read_csv_with_metadata(file_3g_path)
            if df_3g is not None:
                ne_name = NE_NAMES[region]["DGW"]
                metric_col = UGW_METRIC_COLUMNS["THROUGHPUT_3G"]
                throughput_3g = get_max_from_last_7_days(df_3g, ne_name, metric_col)
                print(f"  {region} 3G Gi throughput (max last 7 days): {throughput_3g:.2f} MB/s")
        except Exception as e:
            print(f"  Error getting 3G throughput: {e}")
    
    # Get 4G SGi throughput
    file_4g_name = FILE_PATTERNS["DGW_THROUGHPUT_4G"]
    file_4g_path = os.path.join(CSV_FOLDER, file_4g_name)
    
    throughput_4g = 0
    if os.path.exists(file_4g_path):
        try:
            df_4g = read_csv_with_metadata(file_4g_path)
            if df_4g is not None:
                ne_name = NE_NAMES[region]["DGW"]
                metric_col = UGW_METRIC_COLUMNS["THROUGHPUT_4G"]
                throughput_4g = get_max_from_last_7_days(df_4g, ne_name, metric_col)
                print(f"  {region} 4G SGi throughput (max last 7 days): {throughput_4g:.2f} MB/s")
        except Exception as e:
            print(f"  Error getting 4G throughput: {e}")
    
    total_throughput = throughput_3g + throughput_4g
    print(f"  {region} Total Throughput (Gi + SGi): {throughput_3g:.2f} + {throughput_4g:.2f} = {total_throughput:.2f} MB/s")
    
    return total_throughput


# ============================================
# UPCF FUNCTIONS
# ============================================

def get_upcf_value(region):
    """Get UPCF max active subscribers from last 7 days."""
    file_pattern_key = f"UPCF_{region}"
    
    if file_pattern_key not in FILE_PATTERNS:
        return 0
    
    file_name = FILE_PATTERNS[file_pattern_key]
    file_path = os.path.join(CSV_FOLDER, file_name)
    
    if not os.path.exists(file_path):
        print(f"  UPCF file not found: {file_path}")
        return 0
    
    try:
        df = read_csv_with_metadata(file_path)
        if df is None:
            return 0
        
        ne_name = NE_NAMES[region]["UPCF"]
        return get_max_from_last_7_days(df, ne_name, UPCF_METRIC_COLUMN)
    except Exception as e:
        print(f"  Error processing UPCF for {region}: {e}")
        return 0


# ============================================
# MGW FUNCTIONS
# ============================================

def get_mgw_value(region):
    """Get MGW peak licensed traffic from last 7 days."""
    file_name = FILE_PATTERNS["MGW"]
    file_path = os.path.join(CSV_FOLDER, file_name)
    
    if not os.path.exists(file_path):
        print(f"  MGW file not found: {file_path}")
        return 0
    
    try:
        df = read_csv_with_metadata(file_path)
        if df is None:
            return 0
        
        ne_name = NE_NAMES[region]["MGW"]
        return get_max_from_last_7_days(df, ne_name, MGW_METRIC_COLUMN)
    except Exception as e:
        print(f"  Error processing MGW for {region}: {e}")
        return 0


def get_vmgw_value(region):
    """Get vMGW license usage percentage from last 7 days."""
    file_name = FILE_PATTERNS["VMGW"]
    file_path = os.path.join(CSV_FOLDER, file_name)
    
    if not os.path.exists(file_path):
        print(f"  vMGW file not found: {file_path}")
        return 0
    
    try:
        df = read_csv_with_metadata(file_path)
        if df is None:
            return 0
        
        ne_name = NE_NAMES[region]["VMGW"]
        return get_max_from_last_7_days(df, ne_name, VMGW_METRIC_COLUMN)
    except Exception as e:
        print(f"  Error processing vMGW for {region}: {e}")
        return 0


# ============================================
# GENERATE CALCULATION SHEET
# ============================================

def generate_calculation_sheet():
    """Generate the Calculation sheet with all network elements."""
    
    results = []
    
    # ========================================
    # SECTION 1: Cloud USN
    # ========================================
    print("\n" + "="*50)
    print("PROCESSING CLOUD USN")
    print("="*50)
    
    usn_metrics = ['2G_SAU', '3G_SAU', '4G_SAU', '2G_PDP', '3G_PDP', '4G_PDP']
    usn_metric_names = {
        '2G_SAU': '2G SAU (per K sub)',
        '3G_SAU': '3G SAU (per K sub)',
        '4G_SAU': '4G SAU (per K sub)',
        '2G_PDP': '2G PDP (per K sub)',
        '3G_PDP': '3G PDP (per K sub)',
        '4G_PDP': '4G PDP (per K sub)'
    }
    
    for region in ['LMB', 'LL']:
        print(f"\n--- {region} Cloud USN ---")
        for metric in usn_metrics:
            value = get_cloud_usn_value(region, metric)
            
            results.append({
                'NE Type': 'Cloud USN',
                'Region': region,
                'Metric': usn_metric_names[metric],
                'Value': value
            })
            print(f"  {usn_metric_names[metric]}: {value:,.0f}")
    
    # ========================================
    # SECTION 2: Cloud UGW (vCGW)
    # ========================================
    print("\n" + "="*50)
    print("PROCESSING CLOUD UGW (vCGW)")
    print("="*50)
    
    for region in ['LMB', 'LL']:
        print(f"\n--- {region} Cloud UGW ---")
        
        # vCGW 2G/3G PDP
        value = get_cgw_2g3g_pdp(region)
        results.append({
            'NE Type': 'Cloud UGW',
            'Region': region,
            'Metric': 'vCGW 2G/3G PDP',
            'Value': value
        })
        print(f"  vCGW 2G/3G PDP: {value:,.0f}")
        
        # vCGW 4G PDP
        value = get_cgw_4g_pdp(region)
        results.append({
            'NE Type': 'Cloud UGW',
            'Region': region,
            'Metric': 'vCGW 4G PDP',
            'Value': value
        })
        print(f"  vCGW 4G PDP: {value:,.0f}")
        
        # vCGW Throughput (Gi + SGi)
        value = get_cgw_throughput(region)
        results.append({
            'NE Type': 'Cloud UGW',
            'Region': region,
            'Metric': 'vCGW Throughput (MB/s)',
            'Value': round(value, 2)
        })
        print(f"  vCGW Throughput: {value:.2f} MB/s")
    
    # ========================================
    # SECTION 3: UPCF
    # ========================================
    print("\n" + "="*50)
    print("PROCESSING UPCF")
    print("="*50)
    
    for region in ['LMB', 'LL']:
        print(f"\n--- {region} UPCF ---")
        value = get_upcf_value(region)
        
        results.append({
            'NE Type': 'UPCF',
            'Region': region,
            'Metric': 'Maximum Active Subscribers',
            'Value': value
        })
        print(f"  Maximum Active Subscribers: {value:,.0f}")
    
    # ========================================
    # SECTION 4: MGW
    # ========================================
    print("\n" + "="*50)
    print("PROCESSING MGW")
    print("="*50)
    
    for region in ['LMB', 'LL']:
        print(f"\n--- {region} MGW ---")
        
        # Physical MGW
        value = get_mgw_value(region)
        results.append({
            'NE Type': 'MGW',
            'Region': region,
            'Metric': 'Peak Licensed Traffic',
            'Value': value
        })
        print(f"  Peak Licensed Traffic: {value:,.0f}")
        
        # Cloud vMGW
        value = get_vmgw_value(region)
        results.append({
            'NE Type': 'vMGW',
            'Region': region,
            'Metric': 'License Usage (%)',
            'Value': value
        })
        print(f"  License Usage (%): {value:.2f}%")
    
    return pd.DataFrame(results)


# ============================================
# MAIN
# ============================================

if __name__ == "__main__":
    
    if not os.path.exists(CSV_FOLDER):
        print(f"Creating folder: {CSV_FOLDER}")
        os.makedirs(CSV_FOLDER)
        print(f"Please place your CSV files in '{CSV_FOLDER}' folder and run again.")
    else:
        print("="*50)
        print("GENERATING CALCULATION SHEET")
        print("="*50)
        
        # List all CSV files for debugging
        print("\nCSV files found in data folder:")
        for f in os.listdir(CSV_FOLDER):
            if f.endswith('.csv'):
                print(f"  - {f}")
        
        df_calc = generate_calculation_sheet()
        
        # Write to Excel
        with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
            df_calc.to_excel(writer, sheet_name='Calculation', index=False)
        
        print("\n" + "="*50)
        print(f"✅ Calculation sheet saved to: {OUTPUT_EXCEL}")
        print("="*50)
        
        print("\nFINAL RESULTS:")
        print(df_calc.to_string(index=False))