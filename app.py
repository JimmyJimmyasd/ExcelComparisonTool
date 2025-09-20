import streamlit as st
import pandas as pd
import numpy as np
from rapidfuzz import process, fuzz
import io
import xlsxwriter
from typing import Tuple, List, Dict, Optional

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
        """Perform exact and fuzzy matching between two DataFrames"""
        
        results = {
            'matched': [],
            'suggested': [],
            'unmatched': []
        }
        
        # Create lookup dictionary for Sheet B
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
        
        # Process each row in Sheet A
        for idx_a, row_a in df_a.iterrows():
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
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.subheader("Key Columns")
            key_col_a = st.selectbox(
                "Key column in Sheet A:",
                comparator.df_a.columns,
                help="Column used for matching"
            )
            
            key_col_b = st.selectbox(
                "Key column in Sheet B:",
                comparator.df_b.columns,
                help="Column used for matching"
            )
        
        with col2:
            st.subheader("Columns to Extract")
            cols_to_extract = st.multiselect(
                "Select columns from Sheet B to merge:",
                comparator.df_b.columns,
                help="These columns will be added to the results"
            )
        
        with col3:
            st.subheader("Actions")
            st.write("")  # Spacing
            
            if st.button("🔍 Start Comparison", type="primary", use_container_width=True):
                with st.spinner("Comparing data..."):
                    results = comparator.perform_comparison(
                        comparator.df_a, comparator.df_b,
                        key_col_a, key_col_b,
                        cols_to_extract, threshold, ignore_case
                    )
                    comparator.results = results
                    st.success("Comparison completed!")
        
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
            
            with col2:
                if st.button("📥 Download Excel", type="secondary", use_container_width=True):
                    excel_data = comparator.create_excel_export(comparator.results)
                    st.download_button(
                        label="💾 Download Results",
                        data=excel_data,
                        file_name="comparison_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

if __name__ == "__main__":
    main()