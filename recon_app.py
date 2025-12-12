import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
from io import BytesIO

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
    if pd.isna(name1) or pd.isna(name2):
        return False
    return fuzz.ratio(str(name1).lower().strip(), str(name2).lower().strip()) >= threshold

def cost_matches(cost1, cost2, tolerance=0.01):
    """Check if two costs match within tolerance (1%)"""
    if pd.isna(cost1) or pd.isna(cost2):
        return False
    if cost1 == 0 and cost2 == 0:
        return True
    if cost1 == 0 or cost2 == 0:
        return False
    diff = abs(cost1 - cost2) / max(cost1, cost2)
    return diff <= tolerance

def extract_tm_number(value):
    """Extract TM number, handling various formats"""
    if pd.isna(value):
        return None
    val = str(value).strip().upper()
    if val.startswith("TM"):
        return val
    else:
        return f"TM{val}"

def parse_date(date_val):
    """Parse date and extract month-year"""
    if pd.isna(date_val):
        return None
    try:
        if isinstance(date_val, str):
            # Try dd/MM/YYYY format
            dt = pd.to_datetime(date_val, format='%d/%m/%Y', errors='coerce')
            if pd.isna(dt):
                dt = pd.to_datetime(date_val, errors='coerce')
        else:
            dt = pd.to_datetime(date_val, errors='coerce')
        return dt
    except:
        return None

def to_excel(df):
    """Convert dataframe to Excel bytes"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# File uploaders
st.sidebar.header("📁 Upload Files")
tracker_file = st.sidebar.file_uploader("Job Tracker (.xlsx)", type=['xlsx'])
tm_file = st.sidebar.file_uploader("TM Report (.xlsx)", type=['xlsx'])
xero_file = st.sidebar.file_uploader("Xero Report (.csv)", type=['csv'])

if tracker_file:
    # Load data
    tracker_df = pd.read_excel(tracker_file)
    tracker_df['_TM_NO'] = tracker_df['REPORT TM NO.'].apply(extract_tm_number)
    tracker_df['_TP_NAME'] = tracker_df['REPORT TP/DC NAME (IF APPLICABLE)']
    tracker_df['_FF_DATE'] = tracker_df['FF INSPECTION DATE'].apply(parse_date)
    tracker_df['_MONTH'] = tracker_df['_FF_DATE'].apply(lambda x: x.strftime('%b %Y') if pd.notna(x) else 'No Date')
    
    # Get unique values for filters
    all_months = sorted([m for m in tracker_df['_MONTH'].unique() if m != 'No Date'])
    all_months.append('No Date')
    all_po_types = sorted([str(x) for x in tracker_df['PO TYPE'].dropna().unique()])
    all_statuses = sorted([str(x) for x in tracker_df['STATUS'].dropna().unique()])
    all_clients = sorted([str(x) for x in tracker_df['CLIENT NAME'].dropna().unique()])
    
    # Get all TP names from tracker for exclusion list
    all_tracker_tps = [str(x) for x in tracker_df['_TP_NAME'].dropna().unique()]
    
    st.sidebar.header("🔧 Matching Mode")
    match_mode = st.sidebar.radio(
        "Select comparison",
        ["Tracker vs TM", "TM vs Xero", "3-way Full"],
        index=0
    )
    
    st.sidebar.header("🎛️ Filters")
    
    selected_months = st.sidebar.multiselect("Month (FF Inspection Date)", all_months, default=all_months)
    selected_po_types = st.sidebar.multiselect("PO Type", all_po_types, default=all_po_types)
    selected_statuses = st.sidebar.multiselect("Status", all_statuses, default=all_statuses)
    selected_clients = st.sidebar.multiselect("Client Name", all_clients, default=all_clients)
    
    st.sidebar.header("🚫 Exclude TPs")
    excluded_tps = st.sidebar.multiselect(
        "TPs to exclude from analysis",
        options=all_tracker_tps,
        default=[tp for tp in DEFAULT_EXCLUDED_TPS if tp in all_tracker_tps]
    )
    
    # Apply filters to tracker
    filtered_tracker = tracker_df[
        (tracker_df['_MONTH'].isin(selected_months)) &
        (tracker_df['PO TYPE'].astype(str).isin(selected_po_types)) &
        (tracker_df['STATUS'].astype(str).isin(selected_statuses)) &
        (tracker_df['CLIENT NAME'].astype(str).isin(selected_clients)) &
        (~tracker_df['_TP_NAME'].isin(excluded_tps))
    ].copy()
    
    st.info(f"📊 Filtered tracker: {len(filtered_tracker)} jobs (from {len(tracker_df)} total)")
    
    # Load TM Report if uploaded
    tm_df = None
    if tm_file:
        tm_df = pd.read_excel(tm_file)
        tm_df['_TM_NO'] = tm_df['JobNo'].apply(extract_tm_number)
        tm_df['_TP_NAME'] = tm_df['treeprofessional']
        tm_df['_COST'] = pd.to_numeric(tm_df['TPCost'], errors='coerce').fillna(0)
    
    # Load Xero if uploaded
    xero_df = None
    if xero_file:
        xero_df = pd.read_csv(xero_file)
        xero_df['_TM_NO'] = xero_df['InvoiceNumber'].apply(extract_tm_number)
        xero_df['_TP_NAME'] = xero_df['ContactName']
        xero_df['_COST'] = pd.to_numeric(xero_df['Total'], errors='coerce').fillna(0)
    
    # Run reconciliation
    if st.button("🔍 Run Reconciliation", type="primary"):
        
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
            else:
                for _, row in filtered_tracker.iterrows():
                    tm_no = row['_TM_NO']
                    tracker_tp = row['_TP_NAME']
                    
                    # Find in TM
                    tm_match = tm_df[tm_df['_TM_NO'] == tm_no]
                    
                    if tm_match.empty:
                        results['missing_in_tm'].append({
                            'TM NO': tm_no,
                            'Tracker TP': tracker_tp,
                            'Client': row['CLIENT NAME'],
                            'PO Type': row['PO TYPE'],
                            'Status': row['STATUS'],
                            'FF Date': row['FF INSPECTION DATE']
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
                                'Client': row['CLIENT NAME'],
                                'Status': row['STATUS']
                            })
                        elif not fuzzy_match(tracker_tp, tm_tp):
                            results['tp_mismatch_tm'].append({
                                'TM NO': tm_no,
                                'Tracker TP': tracker_tp,
                                'TM TP': tm_tp,
                                'TM Cost': tm_cost,
                                'Client': row['CLIENT NAME'],
                                'Status': row['STATUS']
                            })
                        else:
                            results['matched'].append({
                                'TM NO': tm_no,
                                'TP': tracker_tp,
                                'TM Cost': tm_cost
                            })
        
        elif match_mode == "TM vs Xero":
            if tm_df is None or xero_df is None:
                st.error("Please upload both TM Report and Xero Report for this comparison")
            else:
                # Filter TM by excluded TPs too
                filtered_tm = tm_df[~tm_df['_TP_NAME'].isin(excluded_tps)].copy()
                
                for _, row in filtered_tm.iterrows():
                    tm_no = row['_TM_NO']
                    tm_tp = row['_TP_NAME']
                    tm_cost = row['_COST']
                    
                    # Find in Xero
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
                            results['cost_mismatch'].append({
                                'TM NO': tm_no,
                                'TP': tm_tp,
                                'TM Cost': tm_cost,
                                'Xero Total': xero_cost,
                                'Difference': abs(tm_cost - xero_cost),
                                'Diff %': f"{abs(tm_cost - xero_cost) / max(tm_cost, xero_cost) * 100:.1f}%"
                            })
                        else:
                            results['matched'].append({
                                'TM NO': tm_no,
                                'TP': tm_tp,
                                'Cost': tm_cost
                            })
        
        elif match_mode == "3-way Full":
            if tm_df is None or xero_df is None:
                st.error("Please upload both TM Report and Xero Report for 3-way comparison")
            else:
                for _, row in filtered_tracker.iterrows():
                    tm_no = row['_TM_NO']
                    tracker_tp = row['_TP_NAME']
                    
                    # Find in TM
                    tm_match = tm_df[tm_df['_TM_NO'] == tm_no]
                    
                    if tm_match.empty:
                        results['missing_in_tm'].append({
                            'TM NO': tm_no,
                            'Tracker TP': tracker_tp,
                            'Client': row['CLIENT NAME'],
                            'Status': row['STATUS']
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
                            'Client': row['CLIENT NAME']
                        })
                        continue
                    
                    if not fuzzy_match(tracker_tp, tm_tp):
                        results['tp_mismatch_tm'].append({
                            'TM NO': tm_no,
                            'Tracker TP': tracker_tp,
                            'TM TP': tm_tp,
                            'TM Cost': tm_cost,
                            'Client': row['CLIENT NAME']
                        })
                        continue
                    
                    # Find in Xero
                    xero_match = xero_df[xero_df['_TM_NO'] == tm_no]
                    
                    if xero_match.empty:
                        results['missing_in_xero'].append({
                            'TM NO': tm_no,
                            'TP': tracker_tp,
                            'TM Cost': tm_cost,
                            'Client': row['CLIENT NAME']
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
                        results['cost_mismatch'].append({
                            'TM NO': tm_no,
                            'TP': tracker_tp,
                            'TM Cost': tm_cost,
                            'Xero Total': xero_cost,
                            'Difference': abs(tm_cost - xero_cost),
                            'Diff %': f"{abs(tm_cost - xero_cost) / max(tm_cost, xero_cost) * 100:.1f}%"
                        })
                    else:
                        results['matched'].append({
                            'TM NO': tm_no,
                            'TP': tracker_tp,
                            'Cost': tm_cost
                        })
        
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
            cols = st.columns(len(mismatch_summary))
            for i, (reason, count) in enumerate(mismatch_summary.items()):
                cols[i].metric(reason, count)
        else:
            st.success("🎉 All jobs matched!")
        
        # Drill-down tables
        st.header("📋 Drill-Down Details")
        
        if results['missing_in_tm']:
            with st.expander(f"❌ Missing in TM ({len(results['missing_in_tm'])})"):
                df = pd.DataFrame(results['missing_in_tm'])
                st.dataframe(df, use_container_width=True)
                st.download_button("📥 Download", to_excel(df), "missing_in_tm.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        if results['missing_in_xero']:
            with st.expander(f"❌ Missing in Xero ({len(results['missing_in_xero'])})"):
                df = pd.DataFrame(results['missing_in_xero'])
                st.dataframe(df, use_container_width=True)
                st.download_button("📥 Download", to_excel(df), "missing_in_xero.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        if results['tp_mismatch_tm']:
            with st.expander(f"⚠️ TP Mismatch - Tracker vs TM ({len(results['tp_mismatch_tm'])})"):
                df = pd.DataFrame(results['tp_mismatch_tm'])
                st.dataframe(df, use_container_width=True)
                st.download_button("📥 Download", to_excel(df), "tp_mismatch_tm.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        if results['tp_mismatch_xero']:
            with st.expander(f"⚠️ TP Mismatch - TM vs Xero ({len(results['tp_mismatch_xero'])})"):
                df = pd.DataFrame(results['tp_mismatch_xero'])
                st.dataframe(df, use_container_width=True)
                st.download_button("📥 Download", to_excel(df), "tp_mismatch_xero.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        if results['no_quote_in_tm']:
            with st.expander(f"💰 No Quote in TM - Cost=0 ({len(results['no_quote_in_tm'])})"):
                df = pd.DataFrame(results['no_quote_in_tm'])
                st.dataframe(df, use_container_width=True)
                st.download_button("📥 Download", to_excel(df), "no_quote_in_tm.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        if results['cost_mismatch']:
            with st.expander(f"💸 Cost Mismatch >1% ({len(results['cost_mismatch'])})"):
                df = pd.DataFrame(results['cost_mismatch'])
                st.dataframe(df, use_container_width=True)
                st.download_button("📥 Download", to_excel(df), "cost_mismatch.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        # Full export
        st.header("📥 Export All Mismatches")
        all_mismatches = []
        for key, items in results.items():
            if key != 'matched' and items:
                for item in items:
                    item['Mismatch Type'] = key.replace('_', ' ').title()
                    all_mismatches.append(item)
        
        if all_mismatches:
            all_df = pd.DataFrame(all_mismatches)
            st.download_button(
                "📥 Download All Mismatches (Excel)",
                to_excel(all_df),
                "all_mismatches.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

else:
    st.info("👈 Please upload the Job Tracker file to begin")
