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

def fuzzy_match(name1, name2, threshold=80):
    """Check if two names match using fuzzy matching"""
    try:
        if pd.isna(name1) or pd.isna(name2):
            return False
        if str(name1).strip() == '' or str(name2).strip() == '':
            return False
        return fuzz.ratio(str(name1).lower().strip(), str(name2).lower().strip()) >= threshold
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
        # Load tracker data
        tracker_df = pd.read_excel(tracker_file)
        
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
                            
                            tm_match = tm_df[tm_df['_TM_NO'] == tm_no]
                            
                            if tm_match.empty:
                                results['missing_in_tm'].append({
                                    'TM NO': tm_no,
                                    'Tracker TP': tracker_tp,
                                    'Client': row.get(client_col, '') if client_col else '',
                                    'PO Type': row.get(po_type_col, '') if po_type_col else '',
                                    'Status': row.get(status_col, '') if status_col else '',
                                    'FF Date': row.get(ff_date_col, '') if ff_date_col else ''
                                })
                            else:
                                tm_row = tm_match.iloc[0]
                                tm_tp = tm_row['_TP_NAME']
                                tm_cost = tm_row['_COST']
                                
                                if tm_cost == 0 or pd.isna(tm_cost):
                                    results['no_quote_in_tm'].append({
                                        'TM NO': tm_no,
                                        'Tracker TP': tracker_tp,
                                        'TM TP': tm_tp,
                                        'TM Cost': tm_cost,
                                        'Client': row.get(client_col, '') if client_col else '',
                                        'Status': row.get(status_col, '') if status_col else ''
                                    })
                                elif not fuzzy_match(tracker_tp, tm_tp):
                                    results['tp_mismatch_tm'].append({
                                        'TM NO': tm_no,
                                        'Tracker TP': tracker_tp,
                                        'TM TP': tm_tp,
                                        'TM Cost': tm_cost,
                                        'Client': row.get(client_col, '') if client_col else '',
                                        'Status': row.get(status_col, '') if status_col else ''
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
                    
                    for _, row in filtered_tm.iterrows():
                        try:
                            tm_no = row['_TM_NO']
                            tm_tp = row['_TP_NAME']
                            tm_cost = row['_COST']
                            
                            xero_match = xero_df[xero_df['_TM_NO'] == tm_no]
                            
                            if xero_match.empty:
                                results['missing_in_xero'].append({
                                    'TM NO': tm_no,
                                    'TM TP': tm_tp,
                                    'TM Cost': tm_cost
                                })
                            else:
                                xero_row = xero_match.iloc[0]
                                xero_tp = xero_row['_TP_NAME']
                                xero_cost = xero_row['_COST']
                                
                                if not fuzzy_match(tm_tp, xero_tp):
                                    results['tp_mismatch_xero'].append({
                                        'TM NO': tm_no,
                                        'TM TP': tm_tp,
                                        'Xero TP': xero_tp,
                                        'TM Cost': tm_cost,
                                        'Xero Total': xero_cost
                                    })
                                elif not cost_matches(tm_cost, xero_cost):
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
                            
                            # Find in TM
                            tm_match = tm_df[tm_df['_TM_NO'] == tm_no]
                            
                            if tm_match.empty:
                                results['missing_in_tm'].append({
                                    'TM NO': tm_no,
                                    'Tracker TP': tracker_tp,
                                    'Client': row.get(client_col, '') if client_col else '',
                                    'Status': row.get(status_col, '') if status_col else ''
                                })
                                continue
                            
                            tm_row = tm_match.iloc[0]
                            tm_tp = tm_row['_TP_NAME']
                            tm_cost = tm_row['_COST']
                            
                            if tm_cost == 0 or pd.isna(tm_cost):
                                results['no_quote_in_tm'].append({
                                    'TM NO': tm_no,
                                    'Tracker TP': tracker_tp,
                                    'TM TP': tm_tp,
                                    'Client': row.get(client_col, '') if client_col else ''
                                })
                                continue
                            
                            if not fuzzy_match(tracker_tp, tm_tp):
                                results['tp_mismatch_tm'].append({
                                    'TM NO': tm_no,
                                    'Tracker TP': tracker_tp,
                                    'TM TP': tm_tp,
                                    'TM Cost': tm_cost,
                                    'Client': row.get(client_col, '') if client_col else ''
                                })
                                continue
                            
                            # Find in Xero
                            xero_match = xero_df[xero_df['_TM_NO'] == tm_no]
                            
                            if xero_match.empty:
                                results['missing_in_xero'].append({
                                    'TM NO': tm_no,
                                    'TP': tracker_tp,
                                    'TM Cost': tm_cost,
                                    'Client': row.get(client_col, '') if client_col else ''
                                })
                                continue
                            
                            xero_row = xero_match.iloc[0]
                            xero_tp = xero_row['_TP_NAME']
                            xero_cost = xero_row['_COST']
                            
                            if not fuzzy_match(tm_tp, xero_tp):
                                results['tp_mismatch_xero'].append({
                                    'TM NO': tm_no,
                                    'TM TP': tm_tp,
                                    'Xero TP': xero_tp,
                                    'TM Cost': tm_cost,
                                    'Xero Total': xero_cost
                                })
                            elif not cost_matches(tm_cost, xero_cost):
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
