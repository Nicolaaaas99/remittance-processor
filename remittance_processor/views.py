# remittance_processor_project/remittance_processor/views.py
import pandas as pd
import numpy as np
from datetime import datetime
import io
from django.http import HttpResponseRedirect
from django.urls import reverse
from django.shortcuts import render
from django.http import HttpResponse
from django.contrib import messages
from django.db import connections
import logging
import time

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Constants
DEBUG_MODE = True
SAMPLE_MODE = False
MAX_DISPLAY_ROWS = 5

def load_remittance(file):
    """Load the remittance file (Excel, no headers, 13 columns, drop last 3 rows)."""
    start_time = time.time()
    logs = []
    try:
        columns = ['DocCode', 'DocType', 'Doc', 'ContractNo', 'Date', 'BranchCode', 
                   'BranchName', 'InvoiceClaimNo', 'Date1', 'DtAmount', 'CtAmount', 
                   'Discount Amnt', 'DiscRate']
        df_all = pd.read_excel(file, header=None, names=columns, engine='openpyxl')
        df_all = df_all.iloc[:-3]
        logs.append("Dropped last 3 rows from remittance file (assumed totals).")
        
        logs.append("Unique DocCode values before filtering:")
        logs.append(str(df_all['DocCode'].astype(str).unique()[:20]))
        
        df = df_all.dropna(how='all').copy()
        dropped_rows = len(df_all) - len(df)
        if dropped_rows > 0:
            logs.append(f"Note: {dropped_rows} empty or all-NaN rows removed from remittance.")
            logs.append("Sample of dropped rows:")
            logs.append(str(df_all[df_all.isna().all(axis=1)][['DocCode', 'DocType', 'BranchCode']].head(MAX_DISPLAY_ROWS)))
        
        if SAMPLE_MODE:
            df = df.head(100)
            logs.append("SAMPLE_MODE: Processing only first 100 rows.")
        
        numeric_cols = ['DtAmount', 'CtAmount', 'Discount Amnt', 'DiscRate']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            if df[col].isna().any():
                logs.append(f"Warning: Non-numeric values found in {col} (converted to 0):")
                logs.append(str(df[df[col].isna()][['DocCode', 'DocType', 'BranchCode', col]].head(MAX_DISPLAY_ROWS)))
        
        # Keep BranchCode as-is and ensure InvoiceClaimNo is text without .0
        df['BranchCode'] = df['BranchCode'].astype(str).str.strip()
        df['InvoiceClaimNo'] = df['InvoiceClaimNo'].astype(str).str.strip().str.replace(r'\.0+$', '', regex=True)
        for col in ['BranchCode', 'InvoiceClaimNo']:
            if df[col].isna().any() or (df[col] == '').any():
                logs.append(f"Warning: Missing or empty {col} values:")
                logs.append(str(df[df[col].isna() | (df[col] == '')][['DocCode', 'DocType', col]].head(MAX_DISPLAY_ROWS)))
        
        logs.append(f"Loaded remittance file with {len(df)} clean records in {time.time() - start_time:.2f} seconds.")
        logs.append(str(df.head(MAX_DISPLAY_ROWS)))
        logs.append("Summary of numeric columns and key fields:")
        for col in numeric_cols + ['BranchCode', 'InvoiceClaimNo']:
            logs.append(f"{col} unique values: {df[col].dropna().unique()[:10]}")
        
        totals = pd.DataFrame()
        return df, totals, logs
    except Exception as e:
        logs.append(f"Error loading remittance file: {e}")
        raise

def load_customers(company_filter=None):
    logs = []
    with connections['default'].cursor() as cursor:
        query = "SELECT StoreCode, Account, Branch, Name, Company FROM CorporateClients"
        params = []
        if company_filter:
            query += " WHERE Company LIKE %s"
            params.append(f"%{company_filter}%")
        cursor.execute(query, params)
        rows = cursor.fetchall()
        df = pd.DataFrame(rows, columns=['StoreCode', 'Account', 'Branch', 'Name', 'Company'])
        df['StoreCode'] = df['StoreCode'].astype(str).str.zfill(5)
        duplicates = df[df['StoreCode'].duplicated(keep=False)]
        if not duplicates.empty:
            logs.append(f"Warning: Duplicate StoreCode values in CorporateClients:")
            logs.append(str(duplicates[['StoreCode', 'Account', 'Branch', 'Name', 'Company']].head(5)))
            df = df.drop_duplicates(subset='StoreCode', keep='first').reset_index(drop=True)
            logs.append(f"Deduplicated CorporateClients, now has {len(df)} records.")
        logs.append(f"Loaded CorporateClients with {len(df)} records (filtered by {company_filter if company_filter else 'all companies'}).")
        logs.append(str(df.head(5)))
    return df, logs

def create_pivot_from_raw_data(pmt_date):
    logs = []
    with connections['default'].cursor() as cursor:
        date_str = datetime.strptime(pmt_date, '%Y-%m-%d').strftime('%Y-%m-%d')
        query = """
        SELECT  
            MatchRef, 
            Reference, 
            CAST(TxDate AS DATE) AS TxDate,
            MatchRef + '-' + CONVERT(varchar(10), CAST(TxDate AS DATE), 120) AS MergedDate,
            MatchRef + '-' + FORMAT(ABS(ROUND(Amnt, 2)), '0.00') + (CASE WHEN Amnt < 0 THEN '-' ELSE '' END) AS MergedAmnt,
            ROUND(Amnt, 2) AS Amnt
        FROM CheckersMain
        WHERE MatchRef IS NOT NULL
        AND TxDate <= CAST(%s AS DATE)
        ORDER BY MatchRef
        """
        cursor.execute(query, [date_str])
        rows = cursor.fetchall()
        # Log the raw column names from the cursor description
        columns = [desc[0] for desc in cursor.description]
        logs.append(f"Query returned columns: {columns}")
        df = pd.DataFrame(rows, columns=columns)
        df['MatchRef'] = df['MatchRef'].astype(str).str.strip()
        df['Reference'] = df['Reference'].astype(str).str.strip()
        logs.append(f"Created pivot DataFrame with columns: {df.columns.tolist()}")
        logs.append(f"Sample pivot_df: {df.head().to_string()}")
    return df, logs


def perform_checks(df, totals, customers_df):
    """Perform validation checks and insert new rows for unmatched BranchCodes."""
    start_time = time.time()
    logs = []
    errors = []
    
    required_columns = ['DocCode', 'DocType', 'ContractNo', 'BranchCode', 'InvoiceClaimNo', 
                        'DtAmount', 'CtAmount', 'Discount Amnt', 'DiscRate']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Missing column in remittance file: {col}")
    
    for col in required_columns:
        if col == 'ContractNo' and df[col].isna().any():
            logs.append(f"Warning: Missing values in {col}:")
            logs.append(str(df[df[col].isna()][['DocCode', 'DocType', 'BranchCode', col]].head(MAX_DISPLAY_ROWS)))
        elif col != 'ContractNo' and df[col].isna().any():
            errors.append(f"Missing values in column: {col}")
            logs.append(f"Rows with missing {col}:")
            logs.append(str(df[df[col].isna()][['DocCode', 'DocType', 'BranchCode', col]].head(MAX_DISPLAY_ROWS)))
    
    for col in ['DtAmount', 'CtAmount', 'Discount Amnt', 'DiscRate']:
        if df[col].isna().any():
            errors.append(f"Non-numeric or missing values in {col} after conversion")
    
    if (~df['BranchCode'].apply(lambda x: isinstance(x, str))).any():
        errors.append(f"Non-string values in BranchCode")
        logs.append(f"Rows with non-string BranchCode:")
        logs.append(str(df[~df['BranchCode'].apply(lambda x: isinstance(x, str))][['DocCode', 'DocType', 'BranchCode']].head(MAX_DISPLAY_ROWS)))
    
    logs.append("Skipped totals check — totals dropped from remittance file.")
    
    df['StoreCode'] = df['BranchCode'].str[-5:]  # Use last 5 characters for StoreCode
    unmatched_branches = df[~df['StoreCode'].isin(customers_df['StoreCode'])]
    if not unmatched_branches.empty:
        logs.append(f"Unmatched BranchCodes found (last 5 chars): {unmatched_branches['StoreCode'].unique().tolist()}")
        new_rows = []
        unmatched_no_match = []
        logs.append(f"Available StoreCode values in customers_df: {customers_df['StoreCode'].astype(str).unique().tolist()}")
        logs.append(f"Available Branch values in customers_df: {customers_df['Branch'].astype(str).unique().tolist()}")
        for branch_code in unmatched_branches['BranchCode'].unique():
            store_code = branch_code[-5:]  # Use last 5 chars as StoreCode
            logs.append(f"Checking BranchCode: {branch_code} (StoreCode: {store_code})")
            if store_code in customers_df['StoreCode'].values:
                logs.append(f"StoreCode {store_code} found in customers_df['StoreCode']")
                continue
            search_code = store_code.lstrip('0')  # Remove leading zeros for search
            logs.append(f"Searching for Branch: {search_code} in customers_df['Branch']")
            match = customers_df[customers_df['Branch'].astype(str).str.strip() == search_code]
            if not match.empty:
                matched_row = match.iloc[0]
                new_row = {
                    'StoreCode': store_code,
                    'Account': matched_row['Account'],
                    'Branch': matched_row['Branch'],
                    'Name': matched_row['Name'],
                    'Company': 'Checkers'  # Default value
                }
                new_rows.append(new_row)
                logs.append(f"Inserting new row in customers_df for BranchCode {branch_code}: {new_row}")
            else:
                branch_name = df[df['BranchCode'] == branch_code]['BranchName'].iloc[0] if not df[df['BranchCode'] == branch_code].empty else ''
                if branch_name:
                    logs.append(f"Searching for BranchName: {branch_name} for BranchCode: {branch_code}")
                    fuzzy_match = customers_df[customers_df['Name'].astype(str).str.contains(branch_name, case=False, na=False)]
                    if not fuzzy_match.empty:
                        matched_row = fuzzy_match.iloc[0]
                        new_row = {
                            'StoreCode': store_code,
                            'Account': matched_row['Account'],
                            'Branch': matched_row['Branch'],
                            'Name': matched_row['Name'],
                            'Company': 'Checkers'  # Default value
                        }
                        new_rows.append(new_row)
                        logs.append(f"Inserting new row in customers_df for BranchCode {branch_code} via fuzzy match: {new_row}")
                    else:
                        unmatched_no_match.append(branch_code)
                        errors.append(f"No matching customer found for BranchCode: {branch_code} (search code: {search_code}, branch name: {branch_name})")
                else:
                    unmatched_no_match.append(branch_code)
                    errors.append(f"No BranchName found for BranchCode: {branch_code} (search code: {search_code})")
        
        if new_rows:
            customers_df_new = customers_df.copy()
            for new_row in new_rows:
                search_code = new_row['Branch']
                match_idx = customers_df_new.index[customers_df_new['Branch'].astype(str).str.strip() == search_code].tolist()
                insert_idx = match_idx[0] + 1 if match_idx else len(customers_df_new)
                customers_df_new = pd.concat([
                    customers_df_new.iloc[:insert_idx],
                    pd.DataFrame([new_row]),
                    customers_df_new.iloc[insert_idx:]
                ]).reset_index(drop=True)
            for col in ['StoreCode', 'Account', 'Branch', 'Name', 'Company']:
                customers_df_new[col] = customers_df_new[col].astype(str).str.strip()
            customers_df = customers_df_new
            logs.append("Updated customers_df with new rows:")
            logs.append(str(customers_df[customers_df['StoreCode'].isin([row['StoreCode'] for row in new_rows])].head(MAX_DISPLAY_ROWS)))
        
        if unmatched_no_match:
            logs.append(f"BranchCodes with no matching StoreCode, Branch, or Name in customers_df: {unmatched_no_match}")
        
        unmatched_after_update = df[~df['StoreCode'].isin(customers_df['StoreCode'])]
        if not unmatched_after_update.empty:
            errors.append(f"Unmatched BranchCodes after update (last 5 chars): {unmatched_after_update['StoreCode'].unique().tolist()}")
    
    if errors:
        logs.append("Validation errors found:")
        for error in errors:
            logs.append(f"- {error}")
        if not DEBUG_MODE:
            raise ValueError("Processing stopped due to validation errors.")
        else:
            logs.append("Continuing in debug mode despite errors...")
    
    logs.append(f"All checks passed in {time.time() - start_time:.2f} seconds." if not errors else f"Checks completed with errors in {time.time() - start_time:.2f} seconds.")
    return df, customers_df, logs, errors

def compute_columns(df, customers_df, pivot_df, pmt_date, pmt_reference):
    """Compute formula-based columns and create processed_df with optimized Reference logic including Amnt match."""
    start_time = time.time()
    logs = []
    
    # Generate MatchRef using vectorized operation, ensuring string format
    df['MatchRef'] = np.where(
        df['DocCode'] == 21,
        df['InvoiceClaimNo'].astype(str).str.strip().str.replace(r'\.0+$', ''),
        df['BranchCode'].str[-5:] + df['InvoiceClaimNo'].astype(str).str.strip().str.replace(r'\.0+$', '')
    )
    
    # Compute Amnt as CtAmount - DtAmount
    df['Amnt'] = df['CtAmount'] - df['DtAmount']
    
    # Convert dates for comparison, handling NaT
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce').fillna(pd.Timestamp('1900-01-01')).dt.strftime('%Y-%m-%d')
    df['Date1'] = pd.to_datetime(df['Date1'], errors='coerce').fillna(pd.Timestamp('1900-01-01')).dt.strftime('%Y-%m-%d')
    pivot_df['TxDate'] = pd.to_datetime(pivot_df['TxDate'], errors='coerce').fillna(pd.Timestamp('1900-01-01')).dt.strftime('%Y-%m-%d')
    
    # Debug: Log sample data and pivot_df columns
    logs.append("Sample MatchRef, Date, Date1, Amnt:")
    logs.append(str(df[['MatchRef', 'Date', 'Date1', 'Amnt']].head(10).to_string()))
    logs.append(f"pivot_df columns: {pivot_df.columns.tolist()}")
    logs.append(f"Sample pivot_df [MatchRef, MergedAmnt, Reference]: {pivot_df[['MatchRef', 'MergedAmnt', 'Reference']].head(10).to_string()}")
    
    # Vectorized Reference assignment with optimized merge
    date_keys = pd.DataFrame({
        'index': df.index,
        'MatchRef_Date': df['MatchRef'] + '-' + df['Date'],
        'MatchRef_Date1': df['MatchRef'] + '-' + df['Date1'],
        'MatchRef_PmtDate': df['MatchRef'] + '-' + pmt_date,
        'MatchRef_Amnt': df['MatchRef'] + '-' + df['Amnt'].round(2).astype(str).str.replace(r'\.0+$', '')
    })
    
    # Check and adjust pivot_expanded based on available columns
    merge_cols = ['MergedDate', 'MergedAmnt', 'Reference']
    if not all(col in pivot_df.columns for col in ['MergedDate', 'MergedAmnt']):
        logs.append("Warning: MergedDate or MergedAmnt not found, falling back to Merged.")
        merge_cols = ['Merged', 'Merged', 'Reference']  # Temporary fallback if new columns missing
    pivot_expanded = pivot_df[merge_cols].copy()
    merge_cols[0] = 'MergedDate'  # Use MergedDate as the primary merge key
    
    merged_df = date_keys.merge(pivot_expanded, left_on='MatchRef_Date', right_on=merge_cols[0], how='left', suffixes=('', '_date'))
    merged_df = merged_df.merge(pivot_expanded, left_on='MatchRef_Date1', right_on=merge_cols[0], how='left', suffixes=('_date', '_date1'))
    merged_df = merged_df.merge(pivot_expanded, left_on='MatchRef_PmtDate', right_on=merge_cols[0], how='left', suffixes=('_date1', '_pmt'))
    merged_df = merged_df.merge(pivot_expanded, left_on='MatchRef_Amnt', right_on=merge_cols[1], how='left', suffixes=('_pmt', '_amnt'))
    
    # Debug: Log sample of merge keys and results
    logs.append("Sample date_keys [MatchRef_Amnt]:")
    logs.append(str(date_keys[['MatchRef_Amnt']].head(10).to_string()))
    logs.append("Sample pivot_expanded [MergedAmnt]:")
    logs.append(str(pivot_expanded[['MergedAmnt']].head(10).to_string()))
    
    # Ensure all Reference columns are present, filling with NaN if missing
    reference_cols = ['Reference_date', 'Reference_date1', 'Reference_pmt', 'Reference_amnt']
    for col in reference_cols:
        if col not in merged_df.columns:
            merged_df[col] = np.nan
    
    # Keep only the first match per original index to avoid duplication
    merged_df = merged_df.drop_duplicates(subset='index', keep='first')
    
    # Join back to original df using index
    df = df.reset_index().merge(merged_df[['index'] + reference_cols], on='index', how='left').set_index('index')
    
    # Assign Reference based on priority: PmtDate, Date, Date1, Amnt, then fallback
    mask_pmt = df['Reference_pmt'].notna()
    mask_date = ~mask_pmt & df['Reference_date'].notna()
    mask_date1 = ~mask_pmt & ~mask_date & df['Reference_date1'].notna()
    mask_amnt = ~mask_pmt & ~mask_date & ~mask_date1 & df['Reference_amnt'].notna()
    df['Reference'] = df['Reference_pmt'].fillna(df['Reference_date']).fillna(df['Reference_date1']).fillna(df['Reference_amnt'])
    df.loc[~mask_pmt & ~mask_date & ~mask_date1 & ~mask_amnt, 'Reference'] = df.loc[~mask_pmt & ~mask_date & ~mask_date1 & ~mask_amnt, 'BranchCode'].str.strip() + '-' + df.loc[~mask_pmt & ~mask_date & ~mask_date1 & ~mask_amnt, 'InvoiceClaimNo'].str.strip()
    
    # Clean up temporary columns
    df.drop(columns=reference_cols, errors='ignore', inplace=True)
    
    # Use last 5 chars of BranchCode for StoreCode and map EvoAccount
    df['StoreCode'] = df['BranchCode'].str[-5:]
    customers_df['StoreCode'] = customers_df['StoreCode'].astype(str)
    df['EvoAccount'] = df['StoreCode'].map(customers_df.set_index('StoreCode')['Account'])
    
    if df['EvoAccount'].isna().any():
        logs.append("Warning: NaN values in EvoAccount after mapping:")
        logs.append(str(df[df['EvoAccount'].isna()][['BranchCode', 'StoreCode', 'EvoAccount']].head(MAX_DISPLAY_ROWS)))
    
    df['PmtDate'] = pmt_date
    df['PmtReference'] = pmt_reference
    df['Module'] = 'AR'
    
    df['TrnCode'] = np.where(df['CtAmount'] - df['DtAmount'] > 0, 'PMT', 'PMTDT')
    
    processed_df = df[['MatchRef', 'EvoAccount', 'Reference', 'PmtDate', 'PmtReference', 'Module', 'TrnCode', 
                      'DtAmount', 'CtAmount', 'Discount Amnt']].copy()
    
    logs.append("TrnCode distribution in processed_df:")
    logs.append(str(processed_df['TrnCode'].value_counts()))
    logs.append("Sample of PMTDT rows with Discount Amnt:")
    logs.append(str(processed_df[processed_df['TrnCode'] == 'PMTDT'][['EvoAccount', 'TrnCode', 'Discount Amnt']].head(MAX_DISPLAY_ROWS)))
    
    logs.append(f"Computed columns and created processed_df in {time.time() - start_time:.2f} seconds.")
    logs.append(str(processed_df.head(MAX_DISPLAY_ROWS)))
    return df, processed_df, logs

def generate_total_import(processed_df, remittance_totals):
    """Generate the total per customer import file from processed_df."""
    start_time = time.time()
    logs = []
    
    processed_df['NettAmnt'] = processed_df['CtAmount'] - processed_df['DtAmount']
    processed_df['NettAmnt'] = processed_df['NettAmnt'].round(2)
    
    total_df = processed_df.groupby('EvoAccount').agg({
        'PmtDate': 'first',
        'Module': 'first',
        'PmtReference': 'first',
        'NettAmnt': 'sum',
        'CtAmount': 'sum',
        'DtAmount': 'sum',
        'Discount Amnt': 'sum'
    }).reset_index()
    
    logs.append("Per-customer sums (before SHO477 rebate and overrides):")
    logs.append(str(total_df[['EvoAccount', 'CtAmount', 'DtAmount', 'Discount Amnt', 'NettAmnt']].head(MAX_DISPLAY_ROWS)))
    
    total_df['Description'] = 'Shoprite Payment'
    total_df['TrnCode'] = total_df['NettAmnt'].apply(lambda x: 'PMT' if x > 0 else 'PMTDT')
    total_df['Amount'] = total_df['NettAmnt'].abs()
    
    pmt_dt_discount = processed_df['Discount Amnt'].sum().round(2)
    logs.append(f"SHO477 Rebate Amount (sum of all Discount Amnt): {pmt_dt_discount:.2f}")
    if pmt_dt_discount != 0:
        rebate_row = pd.DataFrame({
            'PmtDate': [processed_df['PmtDate'].iloc[0]],
            'EvoAccount': ['SHO477'],
            'Module': ['AR'],
            'PmtReference': [processed_df['PmtReference'].iloc[0]],
            'Description': ['Shoprite Rebate'],
            'TrnCode': ['PMTDT'],
            'NettAmnt': [-pmt_dt_discount],
            'Amount': [abs(pmt_dt_discount)],
            'CtAmount': [0],
            'DtAmount': [0],
            'Discount Amnt': [pmt_dt_discount]
        })
        total_df = pd.concat([total_df, rebate_row], ignore_index=True)
    
    total_import_net_sum = total_df['NettAmnt'].sum()
    remittance_net_total = remittance_totals['NettAmnt']
    logs.append("\nTotal Import Amount Check:")
    logs.append(f"Sum of signed NettAmnt in total_import (including SHO477 rebate): {total_import_net_sum:.2f}")
    logs.append(f"Remittance NettAmnt Total: {remittance_net_total:.2f}")
    if abs(total_import_net_sum - remittance_net_total) > 0.01:
        logs.append("Warning: Total Import NettAmnt does not match Remittance NettAmnt Total.")
    
    total_df = total_df[['PmtDate', 'EvoAccount', 'Module', 'TrnCode', 'PmtReference', 'Description', 'Amount']]
    total_df.rename(columns={'PmtReference': 'Reference'}, inplace=True)
    
    logs.append(f"Total import preview (generated in {time.time() - start_time:.2f} seconds):")
    logs.append(str(total_df.head(MAX_DISPLAY_ROWS)))
    return total_df, logs

def generate_detail_import(processed_df):
    """Generate the detail import file from processed_df."""
    start_time = time.time()
    logs = []
    MAX_DISPLAY_ROWS = 5

    processed_df['Amnt'] = processed_df['DtAmount'] - processed_df['CtAmount']

    detail_df = processed_df[['PmtDate', 'EvoAccount', 'Reference', 'PmtReference', 
                             'DtAmount', 'CtAmount', 'Amnt']].copy()
    detail_df.rename(columns={
        'DtAmount': 'Sum of DtAmount',
        'CtAmount': 'Sum of CtAmount',
        'Amnt': 'Sum of Amnt'
    }, inplace=True)

    logs.append("Validating Reference in detail import:")
    logs.append(str(detail_df[['PmtDate', 'EvoAccount', 'Reference', 'PmtReference']].head(MAX_DISPLAY_ROWS)))
    logs.append(f"Detail import generated in {time.time() - start_time:.2f} seconds.")
    return detail_df, logs

def process_remittance_view(request):
    logs = []
    errors = []
    totals = None
    show_download = False

    if request.method == 'POST':
        remittance_file = request.FILES.get('remittance_file')
        pmt_date = request.POST.get('pmt_date')
        pmt_reference = request.POST.get('pmt_reference')
        nett_amount_input = request.POST.get('nett_amount')
        company_filter = request.POST.get('company_filter', 'Checkers')  # Default to Checkers

        if not all([remittance_file, pmt_date, pmt_reference, nett_amount_input]):
            messages.error(request, "Remittance file, payment date, payment reference, and nett amount are required.")
            return render(request, 'remittance_processor/index.html', {'logs': logs, 'errors': errors})

        # Load remittance
        remittance_df, totals_df, remittance_logs = load_remittance(remittance_file)
        logs.extend(remittance_logs)

        # Load customers with company filter
        customers_df, customers_logs = load_customers(company_filter)
        logs.extend(customers_logs)

        # Load pivot from CheckersMain view with date filter
        pivot_df, pivot_logs = create_pivot_from_raw_data(pmt_date)
        logs.extend(pivot_logs)

        # Display totals to get NettAmnt
        totals, totals_logs = display_remittance_totals(remittance_df)
        logs.extend(totals_logs)

        # Validate nett amount
        try:
            nett_amount_input = float(nett_amount_input)
            nett_amount_calculated = totals.get('NettAmnt', 0)
            if abs(nett_amount_input - nett_amount_calculated) > 0.01:
                return HttpResponseRedirect(f"{reverse('remittance_processor:index')}?mismatch=true&input={nett_amount_input:.2f}&calculated={nett_amount_calculated:.2f}")
        except ValueError:
            messages.error(request, "Invalid nett amount format. Please enter a number.")
            return render(request, 'remittance_processor/index.html', {'logs': logs, 'errors': errors, 'totals': totals})

        # Rest of the processing
        remittance_df, customers_df, check_logs, check_errors = perform_checks(remittance_df, totals_df, customers_df)
        logs.extend(check_logs)
        errors.extend(check_errors)

        remittance_df, processed_df, compute_logs = compute_columns(remittance_df, customers_df, pivot_df, pmt_date, pmt_reference)
        logs.extend(compute_logs)

        total_df, total_logs = generate_total_import(processed_df, totals)
        logs.extend(total_logs)
        detail_df, detail_logs = generate_detail_import(processed_df)
        logs.extend(detail_logs)

        request.session['total_df'] = total_df.to_csv(index=False)
        request.session['detail_df'] = detail_df.to_csv(index=False)
        show_download = True
        messages.success(request, "Processing completed successfully. Download the output files below.")

    return render(request, 'remittance_processor/index.html', {
        'logs': logs,
        'errors': errors,
        'totals': totals,
        'show_download': show_download
    })
def display_remittance_totals(df):
    """Calculate totals for DtAmount, CtAmount, Discount Amnt, and NettAmnt."""
    totals = {
        'DtAmount': df['DtAmount'].sum(),
        'CtAmount': df['CtAmount'].sum(),
        'Discount Amnt': df['Discount Amnt'].sum(),
        'NettAmnt': (df['CtAmount'] - df['DtAmount'] - df['Discount Amnt']).sum()
    }
    logs = []
    logs.append("\nRemittance Totals:")
    for col, total in totals.items():
        logs.append(f"Total {col}: {total:.2f}")
    return totals, logs

def download_total_import(request):
    """Serve total_import.csv for download."""
    total_csv = request.session.get('total_df')
    if not total_csv:
        return HttpResponse("No total import file available.", status=400)
    
    response = HttpResponse(total_csv, content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="total_import.csv"'
    return response

def download_detail_import(request):
    """Serve detail_import.csv for download."""
    detail_csv = request.session.get('detail_df')
    if not detail_csv:
        return HttpResponse("No detail import file available.", status=400)
    
    response = HttpResponse(detail_csv, content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="detail_import.csv"'
    return response