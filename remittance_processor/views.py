# remittance_processor_project/remittance_processor/views.py
import pandas as pd
import numpy as np
from datetime import datetime
import io
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
        
        df['BranchCode'] = df['BranchCode'].astype(str).str.strip()
        df['BranchCode'] = df['BranchCode'].apply(
            lambda x: ('0' + x.split('-')[-1].lstrip('0')).zfill(5) if 'G' in x.split('-')[-1]
            else x.split('-')[-1].zfill(5) if x.split('-')[-1].replace('G', '').isdigit()
            else x.zfill(5) if x.replace('G', '').isdigit() or ('G' in x and x[1:].replace('G', '').isdigit())
            else x
        )
        df['InvoiceClaimNo'] = df['InvoiceClaimNo'].astype(str).str.strip()
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

def load_customers():
    logs = []
    with connections['default'].cursor() as cursor:
        cursor.execute("SELECT StoreCode, Account, Branch, Name, Company FROM CorporateClients")
        rows = cursor.fetchall()
        df = pd.DataFrame(rows, columns=['StoreCode', 'Account', 'Branch', 'Name', 'Company'])
        df['StoreCode'] = df['StoreCode'].astype(str).str.zfill(5)  # Padding with zeros
        duplicates = df[df['StoreCode'].duplicated(keep=False)]
        if not duplicates.empty:
            logs.append(f"Warning: Duplicate StoreCode values in CorporateClients:")
            logs.append(str(duplicates[['StoreCode', 'Account', 'Branch', 'Name', 'Company']].head(5)))
            df = df.drop_duplicates(subset='StoreCode', keep='first').reset_index(drop=True)
            logs.append(f"Deduplicated CorporateClients, now has {len(df)} records.")
        logs.append(f"Loaded CorporateClients with {len(df)} records.")
        logs.append(str(df.head(5)))
    return df, logs

def create_pivot_from_raw_data(pmt_date):
    logs = []
    with connections['default'].cursor() as cursor:
        # Convert pmt_date to a SQL-compatible date string
        date_str = datetime.strptime(pmt_date, '%Y-%m-%d').strftime('%Y-%m-%d')
        query = f"SELECT MatchRef, Reference FROM CheckersMain WHERE TxDate <= CAST('{date_str}' AS DATE)"
        cursor.execute(query)
        rows = cursor.fetchall()
        df = pd.DataFrame(rows, columns=['MatchRef', 'Reference'])
        df['MatchRef'] = df['MatchRef'].astype(str).replace('nan', '')
        df['Reference'] = df['Reference'].astype(str).replace('nan', '')
        duplicates = df[df['MatchRef'].duplicated(keep=False)]
        if not duplicates.empty:
            logs.append(f"Warning: Duplicate MatchRef values in CheckersMain:")
            logs.append(str(duplicates[['MatchRef', 'Reference']].head(5)))
            logs.append(f"Deduplicating CheckersMain by keeping first occurrence of MatchRef.")
            df = df.drop_duplicates(subset='MatchRef', keep='first')
        logs.append(f"Created pivot DataFrame from CheckersMain with {len(df)} records.")
        logs.append(str(df.head(5)))
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
    
    df['StoreCode'] = df['BranchCode']
    unmatched_branches = df[~df['StoreCode'].isin(customers_df['StoreCode'])]
    if not unmatched_branches.empty:
        logs.append(f"Unmatched BranchCodes found: {unmatched_branches['BranchCode'].unique().tolist()}")
        new_rows = []
        unmatched_no_match = []
        logs.append(f"Available StoreCode values in customers_df: {customers_df['StoreCode'].astype(str).unique().tolist()}")
        logs.append(f"Available Branch values in customers_df: {customers_df['Branch'].astype(str).unique().tolist()}")
        for branch_code in unmatched_branches['BranchCode'].unique():
            if not branch_code or not isinstance(branch_code, str) or len(branch_code) != 5:
                errors.append(f"Invalid BranchCode: {branch_code} (skipped)")
                unmatched_no_match.append(branch_code)
                continue
            store_code = branch_code
            logs.append(f"Checking BranchCode: {branch_code} (StoreCode: {store_code})")
            if store_code in customers_df['StoreCode'].values:
                logs.append(f"StoreCode {store_code} found in customers_df['StoreCode']")
                continue
            search_code = branch_code[1:] if branch_code.startswith('0') and 'G' in branch_code else branch_code.lstrip('0')
            logs.append(f"Searching for Branch: {search_code} in customers_df['Branch']")
            match = customers_df[customers_df['Branch'].astype(str).str.strip() == search_code]
            if not match.empty:
                matched_row = match.iloc[0]
                new_row = {
                    'StoreCode': store_code,
                    'Account': matched_row['Account'],
                    'Branch': matched_row['Branch'],
                    'Name': matched_row['Name'],
                    # Removed 'Company' or set a default
                    'Company': 'Checkers'  # Default value since it's not in the view
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
            errors.append(f"Unmatched BranchCodes after update: {unmatched_after_update['BranchCode'].unique().tolist()}")
    
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
    """Compute formula-based columns and create processed_df."""
    start_time = time.time()
    logs = []
    
    df['BranchCode'] = df['BranchCode'].astype(str).replace('nan', '')
    df['InvoiceClaimNo'] = df['InvoiceClaimNo'].astype(str).replace('nan', '')
    df['DocCode'] = pd.to_numeric(df['DocCode'], errors='coerce').fillna(0).astype(int)
    
    logs.append("Checking BranchCode and InvoiceClaimNo types:")
    for col in ['BranchCode', 'InvoiceClaimNo']:
        logs.append(f"{col} types: {df[col].apply(type).unique()}")
    
    df['MatchRef'] = np.where(
        df['DocCode'] == 21,
        df['InvoiceClaimNo'],
        df['BranchCode'] + df['InvoiceClaimNo']
    )
    
    df['Match'] = df['MatchRef'].apply(
        lambda x: pivot_df.index[pivot_df['MatchRef'] == x].tolist()[0] + 1
        if x in pivot_df['MatchRef'].values else pd.NA
    )
    
    df['StoreCode'] = df['BranchCode']
    customers_df['StoreCode'] = customers_df['StoreCode'].astype(str)
    df['EvoAccount'] = df['StoreCode'].map(customers_df.set_index('StoreCode')['Account'])
    
    if df['EvoAccount'].isna().any():
        logs.append("Warning: NaN values in EvoAccount after mapping:")
        logs.append(str(df[df['EvoAccount'].isna()][['BranchCode', 'StoreCode', 'EvoAccount']].head(MAX_DISPLAY_ROWS)))
    
    pivot_df = pivot_df.drop_duplicates(subset='MatchRef', keep='first')
    df['Reference'] = np.where(
        df['MatchRef'].isin(pivot_df['MatchRef']),
        df['MatchRef'].map(pivot_df.set_index('MatchRef')['Reference']),
        df['BranchCode'] + '-' + df['InvoiceClaimNo']
    )
    
    df['PmtDate'] = pmt_date
    df['PmtReference'] = pmt_reference
    df['Module'] = 'AR'
    
    df['TrnCode'] = (df['CtAmount'] - df['DtAmount']).apply(lambda x: 'PMT' if pd.notna(x) and x > 0 else 'PMTDT')
    
    processed_df = df[['MatchRef', 'Match', 'EvoAccount', 'Reference', 'PmtDate', 'PmtReference', 'Module', 'TrnCode', 
                      'DtAmount', 'CtAmount', 'Discount Amnt']].copy()
    
    logs.append("TrnCode distribution in processed_df:")
    logs.append(str(processed_df['TrnCode'].value_counts()))
    logs.append("Sample of PMTDT rows with Discount Amnt:")
    logs.append(str(processed_df[processed_df['TrnCode'] == 'PMTDT'][['EvoAccount', 'TrnCode', 'Discount Amnt']].head(MAX_DISPLAY_ROWS)))
    
    logs.append(f"Computed columns and created processed_df in {time.time() - start_time:.2f} seconds.")
    logs.append(str(processed_df.head(MAX_DISPLAY_ROWS)))
    return df, processed_df, logs

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
    
    processed_df['Amnt'] = processed_df['DtAmount'] - processed_df['CtAmount']
    
    detail_df = processed_df[['PmtDate', 'EvoAccount', 'Reference', 'PmtReference', 
                             'DtAmount', 'CtAmount', 'Amnt']].copy()
    detail_df.rename(columns={
        'DtAmount': 'Sum of DtAmount',
        'CtAmount': 'Sum of CtAmount',
        'Amnt': 'Sum of Amnt'
    }, inplace=True)
    
    logs.append(f"Detail import preview (generated in {time.time() - start_time:.2f} seconds):")
    logs.append(str(detail_df.head(MAX_DISPLAY_ROWS)))
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

        if not all([remittance_file, pmt_date, pmt_reference]):
            messages.error(request, "Remittance file, payment date, and payment reference are required.")
            return render(request, 'remittance_processor/index.html', {'logs': logs, 'errors': errors})

        # Load remittance (unchanged)
        remittance_df, totals_df, remittance_logs = load_remittance(remittance_file)
        logs.extend(remittance_logs)

        # Load customers from CorporateClients view
        customers_df, customers_logs = load_customers()
        logs.extend(customers_logs)

        # Load pivot from CheckersMain view with date filter
        pivot_df, pivot_logs = create_pivot_from_raw_data(pmt_date)
        logs.extend(pivot_logs)

        # Rest of the processing (perform_checks, compute_columns, etc.) remains unchanged
        remittance_df, customers_df, check_logs, check_errors = perform_checks(remittance_df, totals_df, customers_df)
        logs.extend(check_logs)
        errors.extend(check_errors)

        totals, totals_logs = display_remittance_totals(remittance_df)
        logs.extend(totals_logs)

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