import streamlit as st
import pandas as pd
import numpy as np
from rapidfuzz import process, fuzz
import io
import xlsxwriter
from typing import Tuple, List, Dict, Optional
import time
from datetime import datetime
from utils import suggest_column_mappings, get_column_info

# Configure Streamlit page
st.set_page_config(
    page_title="Excel Comparison Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class ExcelComparator:
    def __init__(self):
        self.df_a = None
        self.df_b = None
        self.sheet_names_a = []
        self.sheet_names_b = []
        self.results = None
        
    def load_excel_file(self, uploaded_file, file_type: str) -> Tuple[pd.DataFrame, List[str]]:
        """Load Excel file and return DataFrame and sheet names"""
        try:
            # Get all sheet names
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            return None, sheet_names
        except Exception as e:
            st.error(f"Error loading {file_type}: {str(e)}")
            return None, []
    
    def read_sheet(self, uploaded_file, sheet_name: str) -> pd.DataFrame:
        """Read specific sheet from Excel file"""
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            return df
        except Exception as e:
            st.error(f"Error reading sheet '{sheet_name}': {str(e)}")
            return None
    
    def perform_comparison(self, df_a: pd.DataFrame, df_b: pd.DataFrame, 
                          key_col_a: str, key_col_b: str, 
                          cols_to_extract: List[str], threshold: int,
                          ignore_case: bool = True) -> Dict:
        """Perform exact and fuzzy matching between two DataFrames with progress tracking"""
        
        # Initialize progress tracking
        total_rows = len(df_a)
        start_time = time.time()
        
        # Create progress containers
        progress_container = st.container()
        with progress_container:
            st.subheader("🔄 Processing Comparison")
            
            # Main progress bar
            main_progress = st.progress(0, text="Initializing comparison...")
            
            # Status metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                processed_metric = st.metric("Processed", "0", f"of {total_rows:,}")
            with col2:
                matched_metric = st.metric("✅ Matched", "0")
            with col3:
                suggested_metric = st.metric("⚠️ Suggested", "0")
            with col4:
                time_metric = st.metric("⏱️ Time", "0s")
            
            # Live status text
            status_text = st.empty()
            
            # Estimated time remaining
            eta_text = st.empty()
        
        results = {
            'matched': [],
            'suggested': [],
            'unmatched': []
        }
        
        # Phase 1: Building lookup dictionary
        main_progress.progress(0.05, text="📚 Building lookup dictionary...")
        status_text.info("Creating efficient lookup structure for Sheet B...")
        
        b_lookup = {}
        for idx, row in df_b.iterrows():
            key_val = str(row[key_col_b])
            if ignore_case:
                key_val = key_val.lower().strip()
            b_lookup[key_val] = {
                'index': idx,
                'data': row[cols_to_extract].to_dict() if cols_to_extract else {},
                'original_key': row[key_col_b]
            }
        
        # Phase 2: Processing each row with progress updates
        main_progress.progress(0.1, text="🔍 Starting row-by-row comparison...")
        
        for i, (idx_a, row_a) in enumerate(df_a.iterrows()):
            # Calculate progress
            progress = 0.1 + (i / total_rows) * 0.85  # Reserve 10% for setup, 5% for final
            current_time = time.time()
            elapsed_time = current_time - start_time
            
            # Update progress every 5 rows or for small datasets
            if i % max(1, total_rows // 100) == 0 or total_rows < 100:
                # Update main progress bar
                main_progress.progress(
                    progress, 
                    text=f"Processing row {i+1:,} of {total_rows:,} ({((i+1)/total_rows)*100:.1f}%)"
                )
                
                # Update metrics
                processed_metric.metric("Processed", f"{i+1:,}", f"of {total_rows:,}")
                matched_metric.metric("✅ Matched", f"{len(results['matched']):,}")
                suggested_metric.metric("⚠️ Suggested", f"{len(results['suggested']):,}")
                time_metric.metric("⏱️ Time", f"{elapsed_time:.1f}s")
                
                # Calculate and show ETA
                rows_per_second = 0
                if i > 0 and elapsed_time > 0:
                    rows_per_second = i / elapsed_time
                    remaining_rows = total_rows - i
                    eta_seconds = remaining_rows / rows_per_second if rows_per_second > 0 else 0
                    
                    if eta_seconds > 60:
                        eta_display = f"{eta_seconds/60:.1f}m remaining"
                    else:
                        eta_display = f"{eta_seconds:.0f}s remaining"
                    
                    eta_text.text(f"⏳ Estimated time remaining: {eta_display}")
                
                # Show current processing status
                key_val_a = str(row_a[key_col_a])
                if len(key_val_a) > 50:
                    display_key = key_val_a[:47] + "..."
                else:
                    display_key = key_val_a
                
                if rows_per_second > 0:
                    status_text.text(f"🔍 Processing: '{display_key}' | Speed: {rows_per_second:.1f} rows/sec")
                else:
                    status_text.text(f"🔍 Processing: '{display_key}'")
            
            # Original comparison logic
            key_val_a = str(row_a[key_col_a])
            original_key_a = key_val_a
            
            if ignore_case:
                key_val_a = key_val_a.lower().strip()
            
            # Try exact match first
            if key_val_a in b_lookup:
                match_data = b_lookup[key_val_a]
                result_row = row_a.to_dict()
                result_row.update(match_data['data'])
                result_row['match_type'] = 'Exact'
                result_row['similarity_score'] = 100.0
                result_row['matched_key'] = match_data['original_key']
                results['matched'].append(result_row)
                continue
            
            # Try fuzzy matching
            b_keys = list(b_lookup.keys())
            if b_keys:
                match_result = process.extractOne(
                    key_val_a, 
                    b_keys, 
                    scorer=fuzz.ratio
                )
                
                if match_result and match_result[1] >= threshold:
                    matched_key = match_result[0]
                    similarity = match_result[1]
                    match_data = b_lookup[matched_key]
                    
                    result_row = row_a.to_dict()
                    result_row.update(match_data['data'])
                    result_row['match_type'] = 'Fuzzy'
                    result_row['similarity_score'] = similarity
                    result_row['matched_key'] = match_data['original_key']
                    
                    if similarity >= 90:
                        results['matched'].append(result_row)
                    else:
                        results['suggested'].append(result_row)
                else:
                    # No match found
                    result_row = row_a.to_dict()
                    result_row['match_type'] = 'No Match'
                    result_row['similarity_score'] = 0.0
                    result_row['matched_key'] = None
                    results['unmatched'].append(result_row)
            else:
                # No data in Sheet B to match against
                result_row = row_a.to_dict()
                result_row['match_type'] = 'No Match'
                result_row['similarity_score'] = 0.0
                result_row['matched_key'] = None
                results['unmatched'].append(result_row)
        
        # Final progress update
        total_time = time.time() - start_time
        main_progress.progress(1.0, text="✅ Comparison completed successfully!")
        
        # Final metrics update
        processed_metric.metric("Processed", f"{total_rows:,}", "Complete!")
        matched_metric.metric("✅ Matched", f"{len(results['matched']):,}")
        suggested_metric.metric("⚠️ Suggested", f"{len(results['suggested']):,}")
        time_metric.metric("⏱️ Total Time", f"{total_time:.1f}s")
        
        # Success summary
        match_rate = (len(results['matched']) / total_rows) * 100 if total_rows > 0 else 0
        avg_speed = total_rows / total_time if total_time > 0 else 0
        
        status_text.success(
            f"🎉 Processing complete! "
            f"Match rate: {match_rate:.1f}% | "
            f"Average speed: {avg_speed:.1f} rows/sec | "
            f"Total time: {total_time:.1f}s"
        )
        eta_text.empty()
        
        # Brief pause to show completion state
        time.sleep(1.5)
        
        return results
    
    def create_excel_export(self, results: Dict, filename: str = "comparison_results.xlsx") -> io.BytesIO:
        """Create Excel file with comparison results"""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write matched results
            if results['matched']:
                df_matched = pd.DataFrame(results['matched'])
                df_matched.to_excel(writer, sheet_name='Matched', index=False)
            
            # Write suggested matches
            if results['suggested']:
                df_suggested = pd.DataFrame(results['suggested'])
                df_suggested.to_excel(writer, sheet_name='Suggested_Matches', index=False)
            
            # Write unmatched results
            if results['unmatched']:
                df_unmatched = pd.DataFrame(results['unmatched'])
                df_unmatched.to_excel(writer, sheet_name='Unmatched', index=False)
        
        output.seek(0)
        return output

def main():
    st.title("📊 Excel Comparison Tool")
    st.markdown("Upload two Excel files to compare and match data between them")
    
    # Initialize comparator
    if 'comparator' not in st.session_state:
        st.session_state.comparator = ExcelComparator()
    
    comparator = st.session_state.comparator
    
    # Sidebar for file uploads and settings
    with st.sidebar:
        st.header("📁 File Upload")
        
        # File A upload
        uploaded_file_a = st.file_uploader(
            "Choose Sheet A (Excel file)", 
            type=['xlsx', 'xls'],
            key="file_a"
        )
        
        # File B upload
        uploaded_file_b = st.file_uploader(
            "Choose Sheet B (Excel file)", 
            type=['xlsx', 'xls'],
            key="file_b"
        )
        
        st.divider()
        
        # Settings
        st.header("⚙️ Settings")
        threshold = st.slider(
            "Match Threshold (%)", 
            min_value=50, 
            max_value=100, 
            value=80,
            help="Minimum similarity score for fuzzy matching"
        )
        
        ignore_case = st.checkbox(
            "Ignore case and whitespace", 
            value=True,
            help="Normalize text for better matching"
        )
    
    # Main content area
    col1, col2 = st.columns(2)
    
    # Handle Sheet A
    with col1:
        st.subheader("📋 Sheet A")
        if uploaded_file_a:
            # Get sheet names
            _, sheet_names_a = comparator.load_excel_file(uploaded_file_a, "Sheet A")
            
            if sheet_names_a:
                selected_sheet_a = st.selectbox(
                    "Select sheet from File A:",
                    sheet_names_a,
                    key="sheet_a"
                )
                
                # Load selected sheet
                df_a = comparator.read_sheet(uploaded_file_a, selected_sheet_a)
                if df_a is not None:
                    comparator.df_a = df_a
                    st.success(f"Loaded {len(df_a)} rows")
                    st.dataframe(df_a.head(10), width="stretch")
        else:
            st.info("Please upload Sheet A")
    
    # Handle Sheet B
    with col2:
        st.subheader("📋 Sheet B")
        if uploaded_file_b:
            # Get sheet names
            _, sheet_names_b = comparator.load_excel_file(uploaded_file_b, "Sheet B")
            
            if sheet_names_b:
                selected_sheet_b = st.selectbox(
                    "Select sheet from File B:",
                    sheet_names_b,
                    key="sheet_b"
                )
                
                # Load selected sheet
                df_b = comparator.read_sheet(uploaded_file_b, selected_sheet_b)
                if df_b is not None:
                    comparator.df_b = df_b
                    st.success(f"Loaded {len(df_b)} rows")
                    st.dataframe(df_b.head(10), width="stretch")
        else:
            st.info("Please upload Sheet B")
    
    # Column selection and comparison
    if comparator.df_a is not None and comparator.df_b is not None:
        st.divider()
        st.header("🔍 Column Selection & Comparison")
        
        # Smart Column Mapping Feature
        st.subheader("🤖 Smart Column Mapping Suggestions")
        
        with st.expander("💡 AI-Powered Column Suggestions", expanded=True):
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.write("Our AI analyzes column names, data patterns, and content to suggest the best matches:")
            
            with col2:
                if st.button("🔍 Generate Smart Suggestions", type="secondary"):
                    with st.spinner("🧠 Analyzing columns and generating suggestions..."):
                        suggestions = suggest_column_mappings(comparator.df_a, comparator.df_b)
                        st.session_state.column_suggestions = suggestions
            
            # Display suggestions if available
            if hasattr(st.session_state, 'column_suggestions') and st.session_state.column_suggestions:
                st.write("**🎯 Top Column Mapping Suggestions:**")
                
                for i, suggestion in enumerate(st.session_state.column_suggestions[:5]):
                    confidence = suggestion['confidence']
                    
                    # Color code by confidence
                    if confidence >= 80:
                        confidence_color = "🟢"
                        confidence_text = "High Confidence"
                    elif confidence >= 60:
                        confidence_color = "🟡" 
                        confidence_text = "Medium Confidence"
                    else:
                        confidence_color = "🟠"
                        confidence_text = "Low Confidence"
                    
                    with st.container():
                        col_left, col_middle, col_right = st.columns([2, 1, 2])
                        
                        with col_left:
                            st.write(f"**Sheet A:** `{suggestion['column_a']}`")
                        
                        with col_middle:
                            st.write(f"{confidence_color} {confidence:.0f}%")
                            st.caption(confidence_text)
                        
                        with col_right:
                            st.write(f"**Sheet B:** `{suggestion['column_b']}`")
                        
                        # Show reasons
                        reasons_text = " • ".join(suggestion['reasons'][:2])  # Top 2 reasons
                        st.caption(f"💭 {reasons_text}")
                        
                        # Quick apply buttons
                        col_btn1, col_btn2 = st.columns(2)
                        with col_btn1:
                            if st.button(f"✅ Use as Key Columns", key=f"key_{i}"):
                                st.session_state.suggested_key_a = suggestion['column_a']
                                st.session_state.suggested_key_b = suggestion['column_b']
                                st.success(f"Applied key mapping: {suggestion['column_a']} ↔ {suggestion['column_b']}")
                        
                        with col_btn2:
                            if st.button(f"📊 Use for Extraction", key=f"extract_{i}"):
                                if 'suggested_extract' not in st.session_state:
                                    st.session_state.suggested_extract = []
                                if suggestion['column_b'] not in st.session_state.suggested_extract:
                                    st.session_state.suggested_extract.append(suggestion['column_b'])
                                    st.success(f"Added {suggestion['column_b']} to extraction list")
                        
                        st.divider()
        
        # Manual Column Selection (Enhanced with suggestions)
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.subheader("Key Columns")
            
            # Use suggested key columns if available
            key_col_a_index = 0
            key_col_b_index = 0
            
            if hasattr(st.session_state, 'suggested_key_a') and st.session_state.suggested_key_a in comparator.df_a.columns:
                key_col_a_index = list(comparator.df_a.columns).index(st.session_state.suggested_key_a)
            
            if hasattr(st.session_state, 'suggested_key_b') and st.session_state.suggested_key_b in comparator.df_b.columns:
                key_col_b_index = list(comparator.df_b.columns).index(st.session_state.suggested_key_b)
            
            key_col_a = st.selectbox(
                "Key column in Sheet A:",
                comparator.df_a.columns,
                index=key_col_a_index,
                help="Column used for matching (🤖 AI suggestion applied if available)"
            )
            
            key_col_b = st.selectbox(
                "Key column in Sheet B:",
                comparator.df_b.columns,
                index=key_col_b_index,
                help="Column used for matching (🤖 AI suggestion applied if available)"
            )
            
            # Show column analysis for selected key columns
            if st.checkbox("📊 Show Key Column Analysis"):
                with st.expander(f"Analysis: {key_col_a} (Sheet A)"):
                    col_info_a = get_column_info(comparator.df_a)
                    info_a = col_info_a[key_col_a]
                    
                    metric_col1, metric_col2 = st.columns(2)
                    with metric_col1:
                        st.metric("Unique Values", f"{info_a['unique_count']:,}")
                        st.metric("Missing Values", f"{info_a['null_count']:,}")
                    with metric_col2:
                        st.metric("Data Type", info_a['dtype'])
                        st.write("**Sample Values:**")
                        for val in info_a['sample_values']:
                            st.caption(f"• {val}")
                
                with st.expander(f"Analysis: {key_col_b} (Sheet B)"):
                    col_info_b = get_column_info(comparator.df_b)
                    info_b = col_info_b[key_col_b]
                    
                    metric_col1, metric_col2 = st.columns(2)
                    with metric_col1:
                        st.metric("Unique Values", f"{info_b['unique_count']:,}")
                        st.metric("Missing Values", f"{info_b['null_count']:,}")
                    with metric_col2:
                        st.metric("Data Type", info_b['dtype'])
                        st.write("**Sample Values:**")
                        for val in info_b['sample_values']:
                            st.caption(f"• {val}")
        
        with col2:
            st.subheader("Columns to Extract")
            
            # Use suggested extraction columns if available
            default_extract = []
            if hasattr(st.session_state, 'suggested_extract'):
                default_extract = st.session_state.suggested_extract
            
            cols_to_extract = st.multiselect(
                "Select columns from Sheet B to merge:",
                comparator.df_b.columns,
                default=default_extract,
                help="These columns will be added to the results (🤖 AI suggestions applied if available)"
            )
            
            # Quick add all suggested columns
            if hasattr(st.session_state, 'column_suggestions') and st.session_state.column_suggestions:
                if st.button("🚀 Add All AI Suggested Columns"):
                    suggested_cols = [s['column_b'] for s in st.session_state.column_suggestions[:3]]
                    cols_to_extract = list(set(cols_to_extract + suggested_cols))
                    st.rerun()
        
        with col3:
            st.subheader("Actions")
            st.write("")  # Spacing
            
            if st.button("🔍 Start Comparison", type="primary", use_container_width=True):
                # Pre-comparison validation and info
                st.info("🚀 Starting enhanced comparison with real-time progress tracking...")
                
                # Show comparison parameters
                with st.expander("📋 Comparison Parameters", expanded=False):
                    st.write(f"**📊 Data Overview:**")
                    st.write(f"- Sheet A: {len(comparator.df_a):,} rows")
                    st.write(f"- Sheet B: {len(comparator.df_b):,} rows")
                    st.write(f"- Key columns: {key_col_a} ↔ {key_col_b}")
                    st.write(f"- Extracting: {', '.join(cols_to_extract) if cols_to_extract else 'No additional columns'}")
                    st.write(f"- Similarity threshold: {threshold}%")
                    st.write(f"- Case sensitive: {'No' if ignore_case else 'Yes'}")
                
                # Estimate processing time
                estimated_time = len(comparator.df_a) * 0.01  # Rough estimate
                if estimated_time > 60:
                    time_estimate = f"~{estimated_time/60:.1f} minutes"
                else:
                    time_estimate = f"~{estimated_time:.0f} seconds"
                
                st.write(f"⏱️ **Estimated processing time:** {time_estimate}")
                
                # Run comparison with progress tracking
                try:
                    results = comparator.perform_comparison(
                        comparator.df_a, comparator.df_b,
                        key_col_a, key_col_b,
                        cols_to_extract, threshold, ignore_case
                    )
                    comparator.results = results
                    
                    # Show completion celebration
                    st.balloons()
                    st.success("🎉 Comparison completed successfully! Scroll down to view results.")
                    
                except Exception as e:
                    st.error(f"❌ Error during comparison: {str(e)}")
                    st.write("Please check your data and try again.")
        
        # Display results
        if comparator.results:
            st.divider()
            st.header("📊 Results")
            
            # Summary metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("✅ Matched", len(comparator.results['matched']))
            with col2:
                st.metric("⚠️ Suggested", len(comparator.results['suggested']))
            with col3:
                st.metric("❌ Unmatched", len(comparator.results['unmatched']))
            with col4:
                total = len(comparator.results['matched']) + len(comparator.results['suggested']) + len(comparator.results['unmatched'])
                match_rate = (len(comparator.results['matched']) / total * 100) if total > 0 else 0
                st.metric("Match Rate", f"{match_rate:.1f}%")
            
            # Results tabs
            tab1, tab2, tab3 = st.tabs(["✅ Matched", "⚠️ Suggested Matches", "❌ Unmatched"])
            
            with tab1:
                if comparator.results['matched']:
                    df_matched = pd.DataFrame(comparator.results['matched'])
                    st.dataframe(df_matched, width="stretch")
                else:
                    st.info("No exact matches found")
            
            with tab2:
                if comparator.results['suggested']:
                    df_suggested = pd.DataFrame(comparator.results['suggested'])
                    st.dataframe(df_suggested, width="stretch")
                else:
                    st.info("No suggested matches found")
            
            with tab3:
                if comparator.results['unmatched']:
                    df_unmatched = pd.DataFrame(comparator.results['unmatched'])
                    st.dataframe(df_unmatched, width="stretch")
                else:
                    st.info("All records were matched!")
            
            # Export functionality
            st.divider()
            st.header("📥 Export Results")
            
            col1, col2 = st.columns([2, 1])
            with col1:
                st.write("Download the comparison results as an Excel file with separate sheets for each category.")
                st.info("📊 The Excel file will contain separate sheets for Matched, Suggested Matches, and Unmatched records.")
            
            with col2:
                # Enhanced download section with multiple options
                excel_data = comparator.create_excel_export(comparator.results)
                
                # Main download button
                st.download_button(
                    label="� Download Complete Results",
                    data=excel_data,
                    file_name=f"comparison_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                
                # Quick download options
                st.write("**Quick Downloads:**")
                if comparator.results['matched']:
                    matched_df = pd.DataFrame(comparator.results['matched'])
                    matched_csv = matched_df.to_csv(index=False)
                    st.download_button(
                        label="📊 Matched Only (CSV)",
                        data=matched_csv,
                        file_name=f"matched_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                if comparator.results['suggested']:
                    suggested_df = pd.DataFrame(comparator.results['suggested'])
                    suggested_csv = suggested_df.to_csv(index=False)
                    st.download_button(
                        label="⚠️ Suggested Only (CSV)", 
                        data=suggested_csv,
                        file_name=f"suggested_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

if __name__ == "__main__":
    main()