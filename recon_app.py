import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from io import BytesIO
import traceback

st.set_page_config(page_title="Job Reconciliation", layout="wide")
st.title("🌳 Job Reconciliation Tool")

# Default excluded TPs
DEFAULT_EXCLUDED_TPS = [
    "Peter Dubiez Tree Solutions",
    "Auger",
    "Zane Dubiez Tree Solutions",
    "Jorden Pontin Tree Solutions"
]

def normalize_tp_name(name):
    """Normalize TP name by removing common suffixes and cleaning up"""
    if pd.isna(name):
        return ''
    
    name = str(name).lower().strip()
    
    # Common suffixes to remove
    suffixes = [
        ' limited', ' ltd', ' ltd.', ' llp',
        ' tree services', ' tree service', ' tree surgery', ' tree surgeons',
        ' tree care', ' tree solutions', ' trees',
        ' arboricultural', ' arboriculture', ' arborists',
        ' services', ' consultancy', ' contractors',
        ' (east midlands)', ' (midlands)', ' (south)', ' (north)',
        ' uk', ' group'
    ]
    
    for suffix in suffixes:
        if name.endswith(suffix):
            name = name[:-len(suffix)].strip()
    
    # Remove common prefixes that might differ
    prefixes = ['dc ', 'tcr ']
    for prefix in prefixes:
        if name.startswith(prefix):
            name = name[len(prefix):].strip()
    
    return name.strip()

def fuzzy_match(name1, name2, threshold=80):
    """Check if two names match using fuzzy matching with normalization"""
    try:
        if pd.isna(name1) or pd.isna(name2):
            return False
        
        n1 = str(name1).lower().strip()
        n2 = str(name2).lower().strip()
        
        if n1 == '' or n2 == '':
            return False
        
        # Exact match after lowercasing
        if n1 == n2:
            return True
        
        # Normalize and compare
        n1_norm = normalize_tp_name(name1)
        n2_norm = normalize_tp_name(name2)
        
        # Exact match after normalization
        if n1_norm == n2_norm and n1_norm != '':
            return True
        
        # Substring check on normalized names - but only if both names are similar length
        # This avoids "kw" matching "kw edge" or "watson" matching "watson & price"
        if n1_norm and n2_norm:
            shorter = n1_norm if len(n1_norm) <= len(n2_norm) else n2_norm
            longer = n2_norm if len(n1_norm) <= len(n2_norm) else n1_norm
            
            # Only do substring matching if:
            # 1. Shorter name is at least 2 words AND
            # 2. Shorter name is at least 70% the length of longer name
            word_count = len(shorter.split())
            length_ratio = len(shorter) / len(longer) if len(longer) > 0 else 0
            
            if word_count >= 2 and length_ratio >= 0.7:
                if shorter in longer:
                    return True
        
        # Use token_set_ratio on ORIGINAL names (not normalized) for better accuracy
        token_score = fuzz.token_set_ratio(n1, n2)
        if token_score >= 90:  # Higher threshold
            return True
        
        # Partial ratio - good when one is clearly a substring
        partial_score = fuzz.partial_ratio(n1, n2)
        if partial_score >= 95:
            return True
        
        # Standard ratio as fallback
        standard_score = fuzz.ratio(n1, n2)
        if standard_score >= threshold:
            return True
        
        return False
    except Exception:
        return False

def cost_matches(cost1, cost2, tolerance=0.01):
    """Check if two costs match within tolerance (1%)"""
    try:
        if pd.isna(cost1) or pd.isna(cost2):
            return False
        cost1 = float(cost1)
        cost2 = float(cost2)
        if cost1 == 0 and cost2 == 0:
            return True
        if cost1 == 0 or cost2 == 0:
            return False
        diff = abs(cost1 - cost2) / max(cost1, cost2)
        return diff <= tolerance
    except Exception:
        return False

def extract_tm_number(value):
    """Extract TM number, handling various formats"""
    try:
        if pd.isna(value):
            return None
        val = str(value).strip()
        if val == '' or val.lower() == 'nan':
            return None
        val = val.upper()
        if val.startswith("TM"):
            return val
        else:
            return f"TM{val}"
    except Exception:
        return None

def is_valid_tm_number(tm_no):
    """Check if TM number is valid (not None, not empty)"""
    if tm_no is None:
        return False
    if str(tm_no).strip() == '' or str(tm_no).strip().upper() == 'TM':
        return False
    return True

def parse_date(date_val):
    """Parse date and extract month-year"""
    try:
        if pd.isna(date_val):
            return None
        if isinstance(date_val, str):
            if date_val.strip() == '':
                return None
            dt = pd.to_datetime(date_val, format='%d/%m/%Y', errors='coerce')
            if pd.isna(dt):
                dt = pd.to_datetime(date_val, errors='coerce')
        else:
            dt = pd.to_datetime(date_val, errors='coerce')
        return dt
    except Exception:
        return None

def safe_strftime(dt, fmt='%b %Y', default='No Date'):
    """Safely format datetime"""
    try:
        if pd.isna(dt) or dt is None:
            return default
        return dt.strftime(fmt)
    except Exception:
        return default

def to_excel(df):
    """Convert dataframe to Excel bytes"""
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()
    except Exception as e:
        st.error(f"Error creating Excel file: {str(e)}")
        return None

def find_column(df, candidates):
    """Find a column in dataframe using case-insensitive matching"""
    try:
        col_map = {c.lower().strip(): c for c in df.columns}
        for candidate in candidates:
            if candidate.lower() in col_map:
                return col_map[candidate.lower()]
        return None
    except Exception:
        return None

def safe_get_unique(series):
    """Safely get unique values from a series"""
    try:
        return sorted([str(x) for x in series.dropna().unique() if str(x).strip() != ''])
    except Exception:
        return []

# File uploaders
st.sidebar.header("📁 Upload Files")
tracker_file = st.sidebar.file_uploader("Job Tracker (.xlsx)", type=['xlsx'])
tm_file = st.sidebar.file_uploader("TM Report (.xlsx)", type=['xlsx'])
xero_file = st.sidebar.file_uploader("Xero Report (.csv)", type=['csv'])

# Main logic
if tracker_file:
    try:
        # Load tracker data - detect the right sheet
        xlsx = pd.ExcelFile(tracker_file)
        sheet_names = xlsx.sheet_names
        
        # Try to find the main tracker sheet
        tracker_sheet = None
        for name in ['Master Tracker', 'Sheet1', 'Tracker', 'Jobs']:
            if name in sheet_names:
                tracker_sheet = name
                break
        
        if tracker_sheet is None:
            tracker_sheet = sheet_names[0]  # Default to first sheet
        
        tracker_df = pd.read_excel(xlsx, sheet_name=tracker_sheet)
        st.info(f"📋 Using sheet: '{tracker_sheet}'")
        
        # Check if first row is a section header row (common pattern in formatted Excel files)
        # If first column is something like "GENERAL JOB INFORMATION", the real headers are in row 2
        first_col = str(tracker_df.columns[0]).strip().upper()
        if 'GENERAL' in first_col or 'INFORMATION' in first_col or first_col.startswith('UNNAMED'):
            # Re-read with header on row 1 (0-indexed)
            tracker_df = pd.read_excel(xlsx, sheet_name=tracker_sheet, header=1)
            st.info("📋 Detected header row format - using row 2 as column names")
        
        # Remove completely empty columns (Unnamed columns that are all NaN)
        cols_to_drop = [col for col in tracker_df.columns if str(col).startswith('Unnamed') or pd.isna(col)]
        if cols_to_drop:
            tracker_df = tracker_df.drop(columns=cols_to_drop, errors='ignore')
        
        # Show tracker columns for debugging
        with st.expander("🔍 Tracker Columns (for debugging)"):
            st.write(list(tracker_df.columns))
            st.write(tracker_df.head())
        
        # Find required columns in tracker
        tm_no_col = find_column(tracker_df, ['REPORT TM NO.', 'REPORT TM NO', 'TM NO', 'TM NO.', 'TMNO'])
        tp_name_col = find_column(tracker_df, ['REPORT TP/DC NAME (IF APPLICABLE)', 'REPORT TP/DC NAME', 'TP NAME', 'TP/DC NAME'])
        ff_date_col = find_column(tracker_df, ['FF INSPECTION DATE', 'FF DATE', 'INSPECTION DATE'])
        po_type_col = find_column(tracker_df, ['PO TYPE', 'PO_TYPE', 'POTYPE'])
        status_col = find_column(tracker_df, ['STATUS'])
        client_col = find_column(tracker_df, ['CLIENT NAME', 'CLIENT_NAME', 'CLIENTNAME', 'CLIENT'])
        
        # Validate required columns
        missing_cols = []
        if not tm_no_col:
            missing_cols.append("REPORT TM NO.")
        if not tp_name_col:
            missing_cols.append("REPORT TP/DC NAME (IF APPLICABLE)")
        
        if missing_cols:
            st.error(f"❌ Missing required columns in Tracker: {', '.join(missing_cols)}")
            st.info(f"Available columns: {list(tracker_df.columns)}")
            st.stop()
        
        # Process tracker data
        tracker_df['_TM_NO'] = tracker_df[tm_no_col].apply(extract_tm_number)
        tracker_df['_TP_NAME'] = tracker_df[tp_name_col]
        
        if ff_date_col:
            tracker_df['_FF_DATE'] = tracker_df[ff_date_col].apply(parse_date)
            tracker_df['_MONTH'] = tracker_df['_FF_DATE'].apply(lambda x: safe_strftime(x))
        else:
            tracker_df['_MONTH'] = 'No Date'
        
        # Filter out blank TM numbers
        valid_tm_mask = tracker_df['_TM_NO'].apply(is_valid_tm_number)
        blank_tm_count = (~valid_tm_mask).sum()
        tracker_df = tracker_df[valid_tm_mask].copy()
        
        if blank_tm_count > 0:
            st.warning(f"⚠️ Ignored {blank_tm_count} rows with blank REPORT TM NO.")
        
        st.success(f"✅ Tracker loaded: {len(tracker_df)} jobs with valid TM numbers")
        
        # Get unique values for filters
        all_months = safe_get_unique(tracker_df['_MONTH'])
        if 'No Date' not in all_months:
            all_months.append('No Date')
        
        all_po_types = safe_get_unique(tracker_df[po_type_col]) if po_type_col else []
        all_statuses = safe_get_unique(tracker_df[status_col]) if status_col else []
        all_clients = safe_get_unique(tracker_df[client_col]) if client_col else []
        all_tracker_tps = safe_get_unique(tracker_df['_TP_NAME'])
        
        # Sidebar controls
        st.sidebar.header("🔧 Matching Mode")
        match_mode = st.sidebar.radio(
            "Select comparison",
            ["Tracker vs TM", "TM vs Xero", "3-way Full"],
            index=0
        )
        
        st.sidebar.header("🎛️ Filters")
        
        selected_months = st.sidebar.multiselect("Month (FF Inspection Date)", all_months, default=all_months)
        selected_po_types = st.sidebar.multiselect("PO Type", all_po_types, default=all_po_types) if all_po_types else []
        selected_statuses = st.sidebar.multiselect("Status", all_statuses, default=all_statuses) if all_statuses else []
        selected_clients = st.sidebar.multiselect("Client Name", all_clients, default=all_clients) if all_clients else []
        
        st.sidebar.header("🚫 Exclude TPs")
        excluded_tps = st.sidebar.multiselect(
            "TPs to exclude from analysis",
            options=all_tracker_tps,
            default=[tp for tp in DEFAULT_EXCLUDED_TPS if tp in all_tracker_tps]
        )
        
        # Apply filters to tracker
        try:
            filter_mask = tracker_df['_MONTH'].isin(selected_months)
            
            if po_type_col and selected_po_types:
                filter_mask &= tracker_df[po_type_col].astype(str).isin(selected_po_types)
            
            if status_col and selected_statuses:
                filter_mask &= tracker_df[status_col].astype(str).isin(selected_statuses)
            
            if client_col and selected_clients:
                filter_mask &= tracker_df[client_col].astype(str).isin(selected_clients)
            
            if excluded_tps:
                filter_mask &= ~tracker_df['_TP_NAME'].isin(excluded_tps)
            
            filtered_tracker = tracker_df[filter_mask].copy()
        except Exception as e:
            st.error(f"Error applying filters: {str(e)}")
            filtered_tracker = tracker_df.copy()
        
        st.info(f"📊 Filtered tracker: {len(filtered_tracker)} jobs (from {len(tracker_df)} total)")
        
        # Load TM Report if uploaded
        tm_df = None
        if tm_file:
            try:
                tm_df = pd.read_excel(tm_file)
                
                with st.expander("🔍 TM Report Columns (for debugging)"):
                    st.write(list(tm_df.columns))
                    st.write(tm_df.head())
                
                job_col = find_column(tm_df, ['jobno', 'job no', 'job_no', 'jobnumber', 'job number', 'job'])
                tp_col = find_column(tm_df, ['treeprofessional', 'tree professional', 'tpname', 'tp name', 'tp_name', 'contractor'])
                cost_col = find_column(tm_df, ['tpcost', 'tp cost', 'tp_cost', 'cost', 'amount', 'total'])
                
                if job_col and tp_col and cost_col:
                    tm_df['_TM_NO'] = tm_df[job_col].apply(extract_tm_number)
                    tm_df['_TP_NAME'] = tm_df[tp_col]
                    tm_df['_COST'] = pd.to_numeric(tm_df[cost_col], errors='coerce').fillna(0)
                    
                    # Filter out invalid TM numbers
                    valid_mask = tm_df['_TM_NO'].apply(is_valid_tm_number)
                    tm_df = tm_df[valid_mask].copy()
                    
                    st.success(f"✅ TM Report mapped: Job={job_col}, TP={tp_col}, Cost={cost_col} ({len(tm_df)} rows)")
                else:
                    st.error(f"❌ Could not auto-detect TM Report columns")
                    st.warning(f"Looking for: JobNo (found: {job_col}), treeprofessional (found: {tp_col}), TPCost (found: {cost_col})")
                    st.info(f"Available columns: {list(tm_df.columns)}")
                    tm_df = None
            except Exception as e:
                st.error(f"Error loading TM Report: {str(e)}")
                tm_df = None
        
        # Load Xero if uploaded
        xero_df = None
        if xero_file:
            try:
                xero_df = pd.read_csv(xero_file)
                
                with st.expander("🔍 Xero Report Columns (for debugging)"):
                    st.write(list(xero_df.columns))
                    st.write(xero_df.head())
                
                inv_col = find_column(xero_df, ['invoicenumber', 'invoice number', 'invoice_number', 'invoice no', 'invno'])
                contact_col = find_column(xero_df, ['contactname', 'contact name', 'contact_name', 'name', 'supplier'])
                total_col = find_column(xero_df, ['total', 'amount', 'invoicetotal', 'invoice total'])
                
                if inv_col and contact_col and total_col:
                    xero_df['_TM_NO'] = xero_df[inv_col].apply(extract_tm_number)
                    xero_df['_TP_NAME'] = xero_df[contact_col]
                    xero_df['_COST'] = pd.to_numeric(xero_df[total_col], errors='coerce').fillna(0)
                    
                    # Filter out invalid TM numbers
                    valid_mask = xero_df['_TM_NO'].apply(is_valid_tm_number)
                    xero_df = xero_df[valid_mask].copy()
                    
                    st.success(f"✅ Xero Report mapped: Invoice={inv_col}, Contact={contact_col}, Total={total_col} ({len(xero_df)} rows)")
                else:
                    st.error(f"❌ Could not auto-detect Xero Report columns")
                    st.warning(f"Looking for: InvoiceNumber (found: {inv_col}), ContactName (found: {contact_col}), Total (found: {total_col})")
                    st.info(f"Available columns: {list(xero_df.columns)}")
                    xero_df = None
            except Exception as e:
                st.error(f"Error loading Xero Report: {str(e)}")
                xero_df = None
        
        # Run reconciliation
        if st.button("🔍 Run Reconciliation", type="primary"):
            try:
                results = {
                    'matched': [],
                    'missing_in_tm': [],
                    'missing_in_xero': [],
                    'tp_mismatch_tm': [],
                    'tp_mismatch_xero': [],
                    'no_quote_in_tm': [],
                    'cost_mismatch': []
                }
                
                if match_mode == "Tracker vs TM":
                    if tm_df is None:
                        st.error("Please upload TM Report for this comparison")
                        st.stop()
                    
                    for _, row in filtered_tracker.iterrows():
                        try:
                            tm_no = row['_TM_NO']
                            tracker_tp = row['_TP_NAME']
                            
                            # Find all TM rows with this job number
                            tm_matches = tm_df[tm_df['_TM_NO'] == tm_no]
                            
                            if tm_matches.empty:
                                results['missing_in_tm'].append({
                                    'TM NO': tm_no,
                                    'Tracker TP': tracker_tp,
                                    'Client': row.get(client_col, '') if client_col else '',
                                    'PO Type': row.get(po_type_col, '') if po_type_col else '',
                                    'Status': row.get(status_col, '') if status_col else '',
                                    'FF Date': row.get(ff_date_col, '') if ff_date_col else ''
                                })
                            else:
                                # Find the TM row where TP name matches tracker TP (fuzzy)
                                matched_tm_row = None
                                for _, tm_row in tm_matches.iterrows():
                                    if fuzzy_match(tracker_tp, tm_row['_TP_NAME']):
                                        matched_tm_row = tm_row
                                        break
                                
                                if matched_tm_row is None:
                                    # No TP match found among TM rows - list all TM TPs for reference
                                    all_tm_tps = ', '.join(tm_matches['_TP_NAME'].dropna().unique().tolist())
                                    results['tp_mismatch_tm'].append({
                                        'TM NO': tm_no,
                                        'Tracker TP': tracker_tp,
                                        'TM TP(s)': all_tm_tps,
                                        'TM Rows Found': len(tm_matches),
                                        'Client': row.get(client_col, '') if client_col else '',
                                        'Status': row.get(status_col, '') if status_col else ''
                                    })
                                else:
                                    tm_tp = matched_tm_row['_TP_NAME']
                                    tm_cost = matched_tm_row['_COST']
                                    
                                    if tm_cost == 0 or pd.isna(tm_cost):
                                        results['no_quote_in_tm'].append({
                                            'TM NO': tm_no,
                                            'Tracker TP': tracker_tp,
                                            'TM TP': tm_tp,
                                            'TM Cost': tm_cost,
                                            'Client': row.get(client_col, '') if client_col else '',
                                            'Status': row.get(status_col, '') if status_col else '',
                                            'FF Inspection Date': row.get(ff_date_col, '') if ff_date_col else ''
                                        })
                                    else:
                                        results['matched'].append({
                                            'TM NO': tm_no,
                                            'TP': tracker_tp,
                                            'TM Cost': tm_cost
                                        })
                        except Exception as e:
                            st.warning(f"Error processing row {tm_no}: {str(e)}")
                            continue
                
                elif match_mode == "TM vs Xero":
                    if tm_df is None or xero_df is None:
                        st.error("Please upload both TM Report and Xero Report for this comparison")
                        st.stop()
                    
                    # Filter TM by excluded TPs
                    filtered_tm = tm_df[~tm_df['_TP_NAME'].isin(excluded_tps)].copy() if excluded_tps else tm_df.copy()
                    
                    # Process unique TM NO + TP combinations
                    processed_combinations = set()
                    
                    for _, row in filtered_tm.iterrows():
                        try:
                            tm_no = row['_TM_NO']
                            tm_tp = row['_TP_NAME']
                            tm_cost = row['_COST']
                            
                            # Skip if we've already processed this TM NO + TP combination
                            combo_key = (tm_no, str(tm_tp).lower().strip() if pd.notna(tm_tp) else '')
                            if combo_key in processed_combinations:
                                continue
                            processed_combinations.add(combo_key)
                            
                            # Find all Xero rows with this job number
                            xero_matches = xero_df[xero_df['_TM_NO'] == tm_no]
                            
                            if xero_matches.empty:
                                results['missing_in_xero'].append({
                                    'TM NO': tm_no,
                                    'TM TP': tm_tp,
                                    'TM Cost': tm_cost
                                })
                            else:
                                # Find Xero row where TP matches
                                matched_xero_row = None
                                for _, xero_row in xero_matches.iterrows():
                                    if fuzzy_match(tm_tp, xero_row['_TP_NAME']):
                                        matched_xero_row = xero_row
                                        break
                                
                                if matched_xero_row is None:
                                    all_xero_tps = ', '.join(xero_matches['_TP_NAME'].dropna().unique().tolist())
                                    results['tp_mismatch_xero'].append({
                                        'TM NO': tm_no,
                                        'TM TP': tm_tp,
                                        'Xero TP(s)': all_xero_tps,
                                        'TM Cost': tm_cost,
                                        'Xero Rows Found': len(xero_matches)
                                    })
                                else:
                                    xero_tp = matched_xero_row['_TP_NAME']
                                    xero_cost = matched_xero_row['_COST']
                                    
                                    if not cost_matches(tm_cost, xero_cost):
                                        diff = abs(float(tm_cost) - float(xero_cost))
                                        max_cost = max(float(tm_cost), float(xero_cost))
                                        diff_pct = (diff / max_cost * 100) if max_cost > 0 else 0
                                        results['cost_mismatch'].append({
                                            'TM NO': tm_no,
                                            'TP': tm_tp,
                                            'TM Cost': tm_cost,
                                            'Xero Total': xero_cost,
                                            'Difference': diff,
                                            'Diff %': f"{diff_pct:.1f}%"
                                        })
                                    else:
                                        results['matched'].append({
                                            'TM NO': tm_no,
                                            'TP': tm_tp,
                                            'Cost': tm_cost
                                        })
                        except Exception as e:
                            st.warning(f"Error processing TM row {tm_no}: {str(e)}")
                            continue
                
                elif match_mode == "3-way Full":
                    if tm_df is None or xero_df is None:
                        st.error("Please upload both TM Report and Xero Report for 3-way comparison")
                        st.stop()
                    
                    for _, row in filtered_tracker.iterrows():
                        try:
                            tm_no = row['_TM_NO']
                            tracker_tp = row['_TP_NAME']
                            
                            # Find all TM rows with this job number
                            tm_matches = tm_df[tm_df['_TM_NO'] == tm_no]
                            
                            if tm_matches.empty:
                                results['missing_in_tm'].append({
                                    'TM NO': tm_no,
                                    'Tracker TP': tracker_tp,
                                    'Client': row.get(client_col, '') if client_col else '',
                                    'Status': row.get(status_col, '') if status_col else ''
                                })
                                continue
                            
                            # Find the TM row where TP name matches tracker TP (fuzzy)
                            matched_tm_row = None
                            for _, tm_row in tm_matches.iterrows():
                                if fuzzy_match(tracker_tp, tm_row['_TP_NAME']):
                                    matched_tm_row = tm_row
                                    break
                            
                            if matched_tm_row is None:
                                # No TP match found among TM rows
                                all_tm_tps = ', '.join(tm_matches['_TP_NAME'].dropna().unique().tolist())
                                results['tp_mismatch_tm'].append({
                                    'TM NO': tm_no,
                                    'Tracker TP': tracker_tp,
                                    'TM TP(s)': all_tm_tps,
                                    'TM Rows Found': len(tm_matches),
                                    'Client': row.get(client_col, '') if client_col else ''
                                })
                                continue
                            
                            tm_tp = matched_tm_row['_TP_NAME']
                            tm_cost = matched_tm_row['_COST']
                            
                            if tm_cost == 0 or pd.isna(tm_cost):
                                results['no_quote_in_tm'].append({
                                    'TM NO': tm_no,
                                    'Tracker TP': tracker_tp,
                                    'TM TP': tm_tp,
                                    'Client': row.get(client_col, '') if client_col else '',
                                    'FF Inspection Date': row.get(ff_date_col, '') if ff_date_col else ''
                                })
                                continue
                            
                            # Find in Xero - also match by TP name
                            xero_matches = xero_df[xero_df['_TM_NO'] == tm_no]
                            
                            if xero_matches.empty:
                                results['missing_in_xero'].append({
                                    'TM NO': tm_no,
                                    'TP': tracker_tp,
                                    'TM Cost': tm_cost,
                                    'Client': row.get(client_col, '') if client_col else ''
                                })
                                continue
                            
                            # Find the Xero row where TP name matches TM TP (fuzzy)
                            matched_xero_row = None
                            for _, xero_row in xero_matches.iterrows():
                                if fuzzy_match(tm_tp, xero_row['_TP_NAME']):
                                    matched_xero_row = xero_row
                                    break
                            
                            if matched_xero_row is None:
                                all_xero_tps = ', '.join(xero_matches['_TP_NAME'].dropna().unique().tolist())
                                results['tp_mismatch_xero'].append({
                                    'TM NO': tm_no,
                                    'TM TP': tm_tp,
                                    'Xero TP(s)': all_xero_tps,
                                    'TM Cost': tm_cost,
                                    'Xero Rows Found': len(xero_matches)
                                })
                            else:
                                xero_tp = matched_xero_row['_TP_NAME']
                                xero_cost = matched_xero_row['_COST']
                                
                                if not cost_matches(tm_cost, xero_cost):
                                    diff = abs(float(tm_cost) - float(xero_cost))
                                    max_cost = max(float(tm_cost), float(xero_cost))
                                    diff_pct = (diff / max_cost * 100) if max_cost > 0 else 0
                                    results['cost_mismatch'].append({
                                        'TM NO': tm_no,
                                        'TP': tracker_tp,
                                        'TM Cost': tm_cost,
                                        'Xero Total': xero_cost,
                                        'Difference': diff,
                                        'Diff %': f"{diff_pct:.1f}%"
                                    })
                                else:
                                    results['matched'].append({
                                        'TM NO': tm_no,
                                        'TP': tracker_tp,
                                        'Cost': tm_cost
                                    })
                        except Exception as e:
                            st.warning(f"Error processing 3-way row {tm_no}: {str(e)}")
                            continue
                
                # Calculate summary
                total_matched = len(results['matched'])
                total_mismatches = (
                    len(results['missing_in_tm']) +
                    len(results['missing_in_xero']) +
                    len(results['tp_mismatch_tm']) +
                    len(results['tp_mismatch_xero']) +
                    len(results['no_quote_in_tm']) +
                    len(results['cost_mismatch'])
                )
                total_jobs = total_matched + total_mismatches
                match_pct = (total_matched / total_jobs * 100) if total_jobs > 0 else 0
                
                # Display summary
                st.header("📈 Summary")
                col1, col2, col3 = st.columns(3)
                col1.metric("Match %", f"{match_pct:.1f}%")
                col2.metric("Jobs Matching", total_matched)
                col3.metric("Not Matching", total_mismatches)
                
                # Mismatch breakdown
                st.header("🔎 Mismatch Breakdown")
                
                mismatch_summary = {}
                if results['missing_in_tm']:
                    mismatch_summary['Missing in TM'] = len(results['missing_in_tm'])
                if results['missing_in_xero']:
                    mismatch_summary['Missing in Xero'] = len(results['missing_in_xero'])
                if results['tp_mismatch_tm']:
                    mismatch_summary['TP Mismatch (Tracker vs TM)'] = len(results['tp_mismatch_tm'])
                if results['tp_mismatch_xero']:
                    mismatch_summary['TP Mismatch (TM vs Xero)'] = len(results['tp_mismatch_xero'])
                if results['no_quote_in_tm']:
                    mismatch_summary['No Quote in TM (Cost=0)'] = len(results['no_quote_in_tm'])
                if results['cost_mismatch']:
                    mismatch_summary['Cost Mismatch (>1%)'] = len(results['cost_mismatch'])
                
                if mismatch_summary:
                    num_cols = min(len(mismatch_summary), 6)
                    cols = st.columns(num_cols)
                    for i, (reason, count) in enumerate(mismatch_summary.items()):
                        cols[i % num_cols].metric(reason, count)
                else:
                    st.success("🎉 All jobs matched!")
                
                # Drill-down tables
                st.header("📋 Drill-Down Details")
                
                if results['missing_in_tm']:
                    with st.expander(f"❌ Missing in TM ({len(results['missing_in_tm'])})"):
                        df = pd.DataFrame(results['missing_in_tm'])
                        st.dataframe(df, use_container_width=True)
                        excel_data = to_excel(df)
                        if excel_data:
                            st.download_button("📥 Download", excel_data, "missing_in_tm.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_missing_tm")
                
                if results['missing_in_xero']:
                    with st.expander(f"❌ Missing in Xero ({len(results['missing_in_xero'])})"):
                        df = pd.DataFrame(results['missing_in_xero'])
                        st.dataframe(df, use_container_width=True)
                        excel_data = to_excel(df)
                        if excel_data:
                            st.download_button("📥 Download", excel_data, "missing_in_xero.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_missing_xero")
                
                if results['tp_mismatch_tm']:
                    with st.expander(f"⚠️ TP Mismatch - Tracker vs TM ({len(results['tp_mismatch_tm'])})"):
                        df = pd.DataFrame(results['tp_mismatch_tm'])
                        st.dataframe(df, use_container_width=True)
                        excel_data = to_excel(df)
                        if excel_data:
                            st.download_button("📥 Download", excel_data, "tp_mismatch_tm.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_tp_tm")
                
                if results['tp_mismatch_xero']:
                    with st.expander(f"⚠️ TP Mismatch - TM vs Xero ({len(results['tp_mismatch_xero'])})"):
                        df = pd.DataFrame(results['tp_mismatch_xero'])
                        st.dataframe(df, use_container_width=True)
                        excel_data = to_excel(df)
                        if excel_data:
                            st.download_button("📥 Download", excel_data, "tp_mismatch_xero.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_tp_xero")
                
                if results['no_quote_in_tm']:
                    with st.expander(f"💰 No Quote in TM - Cost=0 ({len(results['no_quote_in_tm'])})"):
                        df = pd.DataFrame(results['no_quote_in_tm'])
                        st.dataframe(df, use_container_width=True)
                        excel_data = to_excel(df)
                        if excel_data:
                            st.download_button("📥 Download", excel_data, "no_quote_in_tm.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_no_quote")
                
                if results['cost_mismatch']:
                    with st.expander(f"💸 Cost Mismatch >1% ({len(results['cost_mismatch'])})"):
                        df = pd.DataFrame(results['cost_mismatch'])
                        st.dataframe(df, use_container_width=True)
                        excel_data = to_excel(df)
                        if excel_data:
                            st.download_button("📥 Download", excel_data, "cost_mismatch.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cost")
                
                # Full export
                st.header("📥 Export All Mismatches")
                all_mismatches = []
                for key, items in results.items():
                    if key != 'matched' and items:
                        for item in items:
                            item_copy = item.copy()
                            item_copy['Mismatch Type'] = key.replace('_', ' ').title()
                            all_mismatches.append(item_copy)
                
                if all_mismatches:
                    all_df = pd.DataFrame(all_mismatches)
                    excel_data = to_excel(all_df)
                    if excel_data:
                        st.download_button(
                            "📥 Download All Mismatches (Excel)",
                            excel_data,
                            "all_mismatches.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            key="dl_all"
                        )
            
            except Exception as e:
                st.error(f"Error during reconciliation: {str(e)}")
                st.code(traceback.format_exc())
    
    except Exception as e:
        st.error(f"Error loading Tracker file: {str(e)}")
        st.code(traceback.format_exc())

else:
    st.info("👈 Please upload the Job Tracker file to begin")
