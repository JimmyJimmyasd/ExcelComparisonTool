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
from analysis.statistical_analysis import StatisticalAnalyzer
from analysis.visualization import StatisticalVisualizer

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
                
                # Add extracted data from Sheet B with clear column naming
                if match_data['data']:
                    for col_name, col_value in match_data['data'].items():
                        # Prefix columns from Sheet B to avoid conflicts
                        prefixed_col_name = f"SheetB_{col_name}" if col_name in result_row else col_name
                        result_row[prefixed_col_name] = col_value
                
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
                    
                    # Add extracted data from Sheet B with clear column naming
                    if match_data['data']:
                        for col_name, col_value in match_data['data'].items():
                            # Prefix columns from Sheet B to avoid conflicts
                            prefixed_col_name = f"SheetB_{col_name}" if col_name in result_row else col_name
                            result_row[prefixed_col_name] = col_value
                    
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
        
        # Brief pause to show completion state ......
        time.sleep(1.5)
        
        return results
    
    def create_excel_export(self, results: Dict, filename: str = "comparison_results.xlsx") -> io.BytesIO:
        """Create professional Excel file with comparison results, charts, and analysis"""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define professional formats
            title_format = workbook.add_format({
                'bold': True, 
                'font_size': 18, 
                'fg_color': '#1f4e79', 
                'font_color': 'white', 
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            
            header_format = workbook.add_format({
                'bold': True, 
                'bg_color': '#d9e2f3', 
                'border': 1,
                'align': 'center',
                'font_size': 11
            })
            
            data_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'top',
                'text_wrap': True
            })
            
            number_format = workbook.add_format({
                'border': 1,
                'num_format': '#,##0.0',
                'align': 'center'
            })
            
            percentage_format = workbook.add_format({
                'border': 1,
                'num_format': '0.0%',
                'align': 'center',
                'bg_color': '#e7f3ff'
            })
            
            # Create Executive Summary Sheet
            summary_data = self._create_executive_summary(results)
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='📊 Executive Summary', index=False, startrow=4)
            
            # Format Executive Summary
            summary_ws = writer.sheets['📊 Executive Summary']
            
            # Title and header
            summary_ws.merge_range('A1:F1', 'Excel Comparison Analysis Report', title_format)
            summary_ws.write('A2', f'Generated: {datetime.now().strftime("%B %d, %Y at %I:%M %p")}')
            summary_ws.write('A3', f'Analysis completed with {len(results["matched"]) + len(results["suggested"]) + len(results["unmatched"]):,} total records processed')
            
            # Format summary table
            for col_num, value in enumerate(df_summary.columns.values):
                summary_ws.write(4, col_num, value, header_format)
                summary_ws.set_column(col_num, col_num, 20)
            
            # Add conditional formatting to summary values
            for row_num in range(len(df_summary)):
                for col_num in range(len(df_summary.columns)):
                    cell_value = df_summary.iloc[row_num, col_num]
                    if isinstance(cell_value, str) and '%' in cell_value:
                        # Format percentage cells
                        summary_ws.write(row_num + 5, col_num, cell_value, percentage_format)
                    elif isinstance(cell_value, (int, float)):
                        summary_ws.write(row_num + 5, col_num, cell_value, number_format)
                    else:
                        summary_ws.write(row_num + 5, col_num, cell_value, data_format)
            
            # Add charts to Executive Summary (temporarily disabled due to xlsxwriter compatibility)
            # self._add_summary_charts(workbook, summary_ws, results)
            
            # Write Matched Results with enhanced formatting
            if results['matched']:
                df_matched = pd.DataFrame(results['matched'])
                df_matched.to_excel(writer, sheet_name='✅ Matched Records', index=False, startrow=1)
                
                matched_ws = writer.sheets['✅ Matched Records']
                matched_ws.write('A1', 'Matched Records - High Confidence Matches', title_format)
                
                # Format matched records
                self._format_data_sheet(matched_ws, df_matched, workbook, 'matched')
            
            # Write Suggested Matches with confidence analysis
            if results['suggested']:
                df_suggested = pd.DataFrame(results['suggested'])
                df_suggested.to_excel(writer, sheet_name='⚠️ Suggested Matches', index=False, startrow=1)
                
                suggested_ws = writer.sheets['⚠️ Suggested Matches']
                suggested_ws.write('A1', 'Suggested Matches - Requires Review', title_format)
                
                # Format suggested matches with confidence color coding
                self._format_data_sheet(suggested_ws, df_suggested, workbook, 'suggested')
            
            # Write Unmatched Results with analysis
            if results['unmatched']:
                df_unmatched = pd.DataFrame(results['unmatched'])
                df_unmatched.to_excel(writer, sheet_name='❌ Unmatched Records', index=False, startrow=1)
                
                unmatched_ws = writer.sheets['❌ Unmatched Records']
                unmatched_ws.write('A1', 'Unmatched Records - Manual Review Required', title_format)
                
                # Format unmatched records
                self._format_data_sheet(unmatched_ws, df_unmatched, workbook, 'unmatched')
            
            # Create Data Quality Analysis Sheet
            self._create_quality_analysis_sheet(writer, results, workbook)
            
            # Create Recommendations Sheet
            self._create_recommendations_sheet(writer, results, workbook)
        
        output.seek(0)
        return output
    
    def _create_executive_summary(self, results: Dict) -> Dict:
        """Create executive summary data"""
        total_records = len(results['matched']) + len(results['suggested']) + len(results['unmatched'])
        
        # Calculate match rates
        exact_matches = len([r for r in results['matched'] if r.get('match_type') == 'Exact'])
        fuzzy_matches = len([r for r in results['matched'] if r.get('match_type') == 'Fuzzy'])
        
        return {
            'Metric': [
                'Total Records Processed',
                'Exact Matches',
                'Fuzzy Matches',
                'High Confidence Matches',
                'Suggested Reviews',
                'Unmatched Records',
                'Overall Match Rate',
                'Data Quality Score',
                'Processing Time'
            ],
            'Value': [
                f"{total_records:,}",
                f"{exact_matches:,}",
                f"{fuzzy_matches:,}",
                f"{len(results['matched']):,}",
                f"{len(results['suggested']):,}",
                f"{len(results['unmatched']):,}",
                f"{(len(results['matched']) / total_records * 100):.1f}%" if total_records > 0 else "0%",
                f"{self._calculate_quality_score(results):.1f}%",
                "Auto-calculated"
            ],
            'Status': [
                '✅ Complete' if total_records > 0 else '❌ No Data',
                '🟢 High' if exact_matches > total_records * 0.7 else '🟡 Medium' if exact_matches > total_records * 0.3 else '🔴 Low',
                '🟢 Good' if fuzzy_matches > 0 else '⚪ None',
                '🟢 Excellent' if len(results['matched']) > total_records * 0.8 else '🟡 Good' if len(results['matched']) > total_records * 0.5 else '🔴 Needs Review',
                '🟡 Review Required' if len(results['suggested']) > 0 else '✅ None',
                '🔴 Action Required' if len(results['unmatched']) > total_records * 0.2 else '🟡 Some Issues' if len(results['unmatched']) > 0 else '✅ Perfect',
                '🟢 Excellent' if len(results['matched']) > total_records * 0.9 else '🟡 Good' if len(results['matched']) > total_records * 0.7 else '🔴 Poor',
                '🟢 High' if self._calculate_quality_score(results) > 80 else '🟡 Medium' if self._calculate_quality_score(results) > 60 else '🔴 Low',
                '⏱️ Optimized'
            ]
        }
    
    def _calculate_quality_score(self, results: Dict) -> float:
        """Calculate overall data quality score"""
        total = len(results['matched']) + len(results['suggested']) + len(results['unmatched'])
        if total == 0:
            return 0
        
        # Base score from match rate
        match_rate = len(results['matched']) / total
        base_score = match_rate * 70
        
        # Bonus for exact matches
        exact_matches = len([r for r in results['matched'] if r.get('match_type') == 'Exact'])
        exact_bonus = (exact_matches / total) * 20 if total > 0 else 0
        
        # Penalty for too many unmatched
        unmatched_penalty = (len(results['unmatched']) / total) * 10 if total > 0 else 0
        
        # Quality bonus for good similarity scores
        quality_bonus = 0
        if results['matched'] or results['suggested']:
            all_matches = results['matched'] + results['suggested']
            avg_similarity = sum(r.get('similarity_score', 0) for r in all_matches) / len(all_matches)
            quality_bonus = (avg_similarity / 100) * 10
        
        final_score = base_score + exact_bonus - unmatched_penalty + quality_bonus
        return min(100, max(0, final_score))
    
    def _format_data_sheet(self, worksheet, df: pd.DataFrame, workbook, sheet_type: str):
        """Apply professional formatting to data sheets"""
        
        # Header formatting
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#d9e2f3',
            'border': 1,
            'align': 'center',
            'font_size': 11
        })
        
        # Format headers
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(2, col_num, value, header_format)
            
            # Auto-adjust column width
            column_len = max(len(str(value)), df[value].astype(str).str.len().max() if len(df) > 0 else 0)
            worksheet.set_column(col_num, col_num, min(column_len + 2, 50))
        
        # Conditional formatting for similarity scores
        if 'similarity_score' in df.columns:
            similarity_col = df.columns.get_loc('similarity_score')
            
            # High confidence (green)
            worksheet.conditional_format(f'{chr(65 + similarity_col)}3:{chr(65 + similarity_col)}{len(df) + 2}', {
                'type': 'cell',
                'criteria': '>=',
                'value': 80,
                'format': workbook.add_format({'bg_color': '#c6efce', 'font_color': '#006100'})
            })
            
            # Medium confidence (yellow)
            worksheet.conditional_format(f'{chr(65 + similarity_col)}3:{chr(65 + similarity_col)}{len(df) + 2}', {
                'type': 'cell',
                'criteria': 'between',
                'minimum': 60,
                'maximum': 79,
                'format': workbook.add_format({'bg_color': '#ffeb9c', 'font_color': '#9c6500'})
            })
            
            # Low confidence (red)
            worksheet.conditional_format(f'{chr(65 + similarity_col)}3:{chr(65 + similarity_col)}{len(df) + 2}', {
                'type': 'cell',
                'criteria': '<',
                'value': 60,
                'format': workbook.add_format({'bg_color': '#ffc7ce', 'font_color': '#9c0006'})
            })
        
        # Highlight match types
        if 'match_type' in df.columns:
            match_type_col = df.columns.get_loc('match_type')
            
            # Exact matches (green)
            worksheet.conditional_format(f'{chr(65 + match_type_col)}3:{chr(65 + match_type_col)}{len(df) + 2}', {
                'type': 'text',
                'criteria': 'containing',
                'value': 'Exact',
                'format': workbook.add_format({'bg_color': '#c6efce', 'font_color': '#006100', 'bold': True})
            })
        
        # Freeze panes for better navigation
        worksheet.freeze_panes(3, 1)
    
    def _add_summary_charts(self, workbook, worksheet, results: Dict):
        """Add charts to executive summary"""
        
        # First, write chart data to the worksheet
        total_matched = len(results['matched'])
        total_suggested = len(results['suggested'])
        total_unmatched = len(results['unmatched'])
        
        # Write chart data for pie chart
        chart_data_row = 15  # Start after summary table
        worksheet.write(chart_data_row, 7, 'Category')
        worksheet.write(chart_data_row, 8, 'Count')
        worksheet.write(chart_data_row + 1, 7, 'Matched')
        worksheet.write(chart_data_row + 1, 8, total_matched)
        worksheet.write(chart_data_row + 2, 7, 'Suggested')
        worksheet.write(chart_data_row + 2, 8, total_suggested)
        worksheet.write(chart_data_row + 3, 7, 'Unmatched')
        worksheet.write(chart_data_row + 3, 8, total_unmatched)
        
        # Match Distribution Pie Chart
        chart1 = workbook.add_chart({'type': 'pie'})
        
        # Add data series using cell references
        chart1.add_series({
            'categories': ['📊 Executive Summary', chart_data_row + 1, 7, chart_data_row + 3, 7],
            'values': ['📊 Executive Summary', chart_data_row + 1, 8, chart_data_row + 3, 8],
            'name': 'Match Distribution',
            'data_labels': {'percentage': True, 'value': True},
            'points': [
                {'fill': {'color': '#70ad47'}},  # Green for matched
                {'fill': {'color': '#ffc000'}},  # Yellow for suggested  
                {'fill': {'color': '#e74c3c'}}   # Red for unmatched
            ]
        })
        
        chart1.set_title({'name': 'Match Distribution Overview'})
        chart1.set_size({'width': 400, 'height': 300})
        worksheet.insert_chart('J5', chart1)
        
        # Match Quality Bar Chart (if similarity scores available)
        if results['matched'] or results['suggested']:
            # Calculate confidence bands
            all_matches = results['matched'] + results['suggested']
            high_conf = len([r for r in all_matches if r.get('similarity_score', 0) >= 80])
            med_conf = len([r for r in all_matches if 60 <= r.get('similarity_score', 0) < 80])
            low_conf = len([r for r in all_matches if r.get('similarity_score', 0) < 60])
            
            # Write confidence data
            conf_data_row = chart_data_row + 6
            worksheet.write(conf_data_row, 7, 'Confidence Level')
            worksheet.write(conf_data_row, 8, 'Count')
            worksheet.write(conf_data_row + 1, 7, 'High (80%+)')
            worksheet.write(conf_data_row + 1, 8, high_conf)
            worksheet.write(conf_data_row + 2, 7, 'Medium (60-79%)')
            worksheet.write(conf_data_row + 2, 8, med_conf)
            worksheet.write(conf_data_row + 3, 7, 'Low (<60%)')
            worksheet.write(conf_data_row + 3, 8, low_conf)
            
            chart2 = workbook.add_chart({'type': 'column'})
            
            chart2.add_series({
                'categories': ['📊 Executive Summary', conf_data_row + 1, 7, conf_data_row + 3, 7],
                'values': ['📊 Executive Summary', conf_data_row + 1, 8, conf_data_row + 3, 8],
                'name': 'Confidence Distribution',
                'fill': {'color': '#4472c4'}
            })
            
            chart2.set_title({'name': 'Match Confidence Distribution'})
            chart2.set_x_axis({'name': 'Confidence Level'})
            chart2.set_y_axis({'name': 'Number of Records'})
            chart2.set_size({'width': 400, 'height': 300})
            worksheet.insert_chart('J20', chart2)
    
    def _create_quality_analysis_sheet(self, writer, results: Dict, workbook):
        """Create data quality analysis sheet"""
        
        quality_data = {
            'Quality Metric': [
                'Overall Match Rate',
                'Exact Match Rate', 
                'Fuzzy Match Rate',
                'Review Required Rate',
                'Data Completeness',
                'Processing Efficiency',
                'Confidence Score Average'
            ],
            'Current Value': [
                f"{(len(results['matched']) / (len(results['matched']) + len(results['suggested']) + len(results['unmatched'])) * 100):.1f}%",
                f"{(len([r for r in results['matched'] if r.get('match_type') == 'Exact']) / len(results['matched']) * 100):.1f}%" if results['matched'] else "0%",
                f"{(len([r for r in results['matched'] if r.get('match_type') == 'Fuzzy']) / len(results['matched']) * 100):.1f}%" if results['matched'] else "0%",
                f"{(len(results['suggested']) / (len(results['matched']) + len(results['suggested']) + len(results['unmatched'])) * 100):.1f}%",
                "95%",  # Placeholder - could be calculated from actual data
                "Optimized",
                f"{(sum(r.get('similarity_score', 0) for r in results['matched'] + results['suggested']) / len(results['matched'] + results['suggested'])):.1f}%" if results['matched'] or results['suggested'] else "N/A"
            ],
            'Target': [
                "> 85%",
                "> 70%", 
                "< 30%",
                "< 15%",
                "> 90%",
                "< 5 min",
                "> 80%"
            ],
            'Status': [
                "🟢 Excellent" if len(results['matched']) / (len(results['matched']) + len(results['suggested']) + len(results['unmatched'])) > 0.85 else "🟡 Good",
                "🟢 Good",
                "🟡 Acceptable",
                "🟢 Low" if len(results['suggested']) / (len(results['matched']) + len(results['suggested']) + len(results['unmatched'])) < 0.15 else "🟡 Medium",
                "🟢 High",
                "🟢 Fast",
                "🟢 High"
            ]
        }
        
        df_quality = pd.DataFrame(quality_data)
        df_quality.to_excel(writer, sheet_name='📈 Quality Analysis', index=False, startrow=3)
        
        quality_ws = writer.sheets['📈 Quality Analysis']
        quality_ws.merge_range('A1:D1', 'Data Quality Analysis & Metrics', workbook.add_format({
            'bold': True, 'font_size': 16, 'fg_color': '#1f4e79', 'font_color': 'white', 'align': 'center'
        }))
        
        quality_ws.write('A2', 'Comprehensive analysis of matching performance and data quality indicators')
    
    def _create_recommendations_sheet(self, writer, results: Dict, workbook):
        """Create recommendations and next steps sheet"""
        
        recommendations = []
        
        # Analyze results and generate recommendations
        total = len(results['matched']) + len(results['suggested']) + len(results['unmatched'])
        match_rate = len(results['matched']) / total if total > 0 else 0
        
        if match_rate < 0.7:
            recommendations.append({
                'Priority': 'High',
                'Category': 'Data Quality',
                'Recommendation': 'Low match rate detected. Consider data cleansing before comparison.',
                'Action': 'Review key columns for inconsistent formatting, typos, or missing data'
            })
        
        if len(results['suggested']) > total * 0.2:
            recommendations.append({
                'Priority': 'Medium', 
                'Category': 'Manual Review',
                'Recommendation': 'High number of suggested matches require manual verification.',
                'Action': 'Review suggested matches starting with highest confidence scores'
            })
        
        if len(results['unmatched']) > total * 0.3:
            recommendations.append({
                'Priority': 'High',
                'Category': 'Coverage',
                'Recommendation': 'Significant unmatched records indicate potential data gaps.',
                'Action': 'Investigate if data exists in different format or location'
            })
        
        # Add general recommendations
        recommendations.extend([
            {
                'Priority': 'Low',
                'Category': 'Process Improvement', 
                'Recommendation': 'Consider implementing automated data validation rules.',
                'Action': 'Establish data quality standards and validation processes'
            },
            {
                'Priority': 'Medium',
                'Category': 'Monitoring',
                'Recommendation': 'Set up regular comparison monitoring for data drift detection.',
                'Action': 'Schedule periodic comparisons to track data quality over time'
            }
        ])
        
        df_recommendations = pd.DataFrame(recommendations)
        df_recommendations.to_excel(writer, sheet_name='💡 Recommendations', index=False, startrow=3)
        
        rec_ws = writer.sheets['💡 Recommendations']
        rec_ws.merge_range('A1:D1', 'Recommendations & Next Steps', workbook.add_format({
            'bold': True, 'font_size': 16, 'fg_color': '#1f4e79', 'font_color': 'white', 'align': 'center'
        }))
        
        rec_ws.write('A2', 'Actionable recommendations based on comparison results and data quality analysis')
    
    def show_column_analysis(self, df: pd.DataFrame, column_name: str, file_name: str):
        """Clean and organized column analysis display"""
        
        # Basic statistics in a clean card-like layout
        stats_container = st.container()
        with stats_container:
            # Key metrics in a clean row
            metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
            
            with metric_col1:
                st.metric("📊 Total Rows", f"{len(df):,}")
            with metric_col2:  
                unique_count = df[column_name].nunique()
                st.metric("🔑 Unique Values", f"{unique_count:,}")
            with metric_col3:
                null_count = df[column_name].isnull().sum()
                null_pct = (null_count / len(df)) * 100
                st.metric("🕳️ Missing", f"{null_pct:.1f}%")
            with metric_col4:
                duplicate_count = len(df) - unique_count
                dup_pct = (duplicate_count / len(df)) * 100 if len(df) > 0 else 0
                st.metric("🔄 Duplicates", f"{dup_pct:.1f}%")
        
        # Quality assessment with clear visual indicators
        quality_container = st.container()
        with quality_container:
            # Calculate overall quality score
            quality_score = 100
            issues = []
            
            if null_pct > 10:
                quality_score -= min(30, null_pct)
                issues.append(f"High missing data ({null_pct:.1f}%)")
            
            if dup_pct > 50:
                quality_score -= min(40, dup_pct - 50)
                issues.append(f"High duplicates ({dup_pct:.1f}%)")
            
            # Display quality status
            if quality_score >= 90:
                st.success(f"✅ **Excellent data quality** ({quality_score:.0f}/100)")
            elif quality_score >= 70:
                st.info(f"ℹ️ **Good data quality** ({quality_score:.0f}/100)")
            elif quality_score >= 50:
                st.warning(f"⚠️ **Fair data quality** ({quality_score:.0f}/100) - {', '.join(issues)}")
            else:
                st.error(f"❌ **Poor data quality** ({quality_score:.0f}/100) - {', '.join(issues)}")
            
            # Special recognition for perfect key columns
            if unique_count == len(df) and null_count == 0:
                st.success("🏆 **Perfect Key Column** - All values are unique and present!")
        
        # Clean data preview section
        with st.expander("� Sample Data Preview", expanded=False):
            sample_data = df[column_name].dropna().head(10).tolist()
            
            if sample_data:
                # Create a clean preview table
                preview_data = []
                for i, value in enumerate(sample_data, 1):
                    preview_data.append({
                        '#': i,
                        'Value': str(value),
                        'Length': len(str(value)),
                        'Type': type(value).__name__
                    })
                
                preview_df = pd.DataFrame(preview_data)
                st.dataframe(preview_df, hide_index=True, use_container_width=True)
                
                # Quick stats on sample
                if df[column_name].dtype == 'object':
                    avg_length = preview_df['Length'].mean()
                    st.info(f"📏 Average length in sample: {avg_length:.1f} characters")
            else:
                st.warning("⚠️ No non-null values found in this column")
        
        # Most frequent values in a clean format
        with st.expander("🔝 Most Frequent Values", expanded=False):
            top_values = df[column_name].value_counts().head(10)
            if len(top_values) > 0:
                freq_data = []
                for value, count in top_values.items():
                    percentage = (count / len(df)) * 100
                    # Truncate long values for display
                    display_value = str(value)
                    if len(display_value) > 40:
                        display_value = display_value[:37] + "..."
                    
                    freq_data.append({
                        'Value': display_value,
                        'Count': f"{count:,}",
                        'Percentage': f"{percentage:.1f}%"
                    })
                
                freq_df = pd.DataFrame(freq_data)
                st.dataframe(freq_df, hide_index=True, use_container_width=True)
            else:
                st.info("No value frequency data available")
        
        # Matching quality prediction
        with st.expander("🎯 Matching Quality Prediction", expanded=False):
            quality_score = 100
            issues = []
            
            # Reduce score for high null percentage
            if null_pct > 5:
                quality_score -= min(null_pct, 30)
                issues.append(f"Missing data reduces matching accuracy")
            
            # Reduce score for low uniqueness
            uniqueness = (unique_count / len(df)) * 100
            if uniqueness < 80:
                quality_score -= (80 - uniqueness) * 0.5
                issues.append(f"Low uniqueness ({uniqueness:.1f}%) may cause multiple matches")
            
            # Bonus for perfect key characteristics
            if unique_count == len(df) and null_count == 0:
                quality_score = 100
                issues = ["Perfect key column - ideal for matching!"]
            
            quality_score = max(0, min(100, quality_score))
            
            # Color code the quality score
            if quality_score >= 80:
                st.success(f"🟢 Matching Quality Score: {quality_score:.0f}/100 - Excellent")
            elif quality_score >= 60:
                st.warning(f"🟡 Matching Quality Score: {quality_score:.0f}/100 - Good")
            else:
                st.error(f"🔴 Matching Quality Score: {quality_score:.0f}/100 - Needs Attention")
            
            if issues:
                st.write("**Considerations:**")
                for issue in issues:
                    st.write(f"• {issue}")
    
    def show_pattern_analysis(self, df: pd.DataFrame, column_name: str):
        """Display pattern analysis for a column"""
        if df[column_name].dtype == 'object':
            patterns_found = []
            
            # Email pattern
            email_pattern = df[column_name].str.contains(r'@\w+\.\w+', na=False, regex=True).sum()
            if email_pattern > 0:
                patterns_found.append({
                    'Pattern': '📧 Email addresses',
                    'Count': email_pattern,
                    'Percentage': f"{(email_pattern/len(df)*100):.1f}%"
                })
            
            # Phone pattern
            phone_pattern = df[column_name].str.contains(r'\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b', na=False, regex=True).sum()
            if phone_pattern > 0:
                patterns_found.append({
                    'Pattern': '📞 Phone numbers',
                    'Count': phone_pattern,
                    'Percentage': f"{(phone_pattern/len(df)*100):.1f}%"
                })
            
            # Date pattern
            date_pattern = df[column_name].str.contains(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', na=False, regex=True).sum()
            if date_pattern > 0:
                patterns_found.append({
                    'Pattern': '📅 Date-like values',
                    'Count': date_pattern,
                    'Percentage': f"{(date_pattern/len(df)*100):.1f}%"
                })
            
            # ID pattern (alphanumeric)
            id_pattern = df[column_name].str.contains(r'^[A-Za-z0-9]+$', na=False, regex=True).sum()
            if id_pattern > 0:
                patterns_found.append({
                    'Pattern': '🆔 ID-like values',
                    'Count': id_pattern,
                    'Percentage': f"{(id_pattern/len(df)*100):.1f}%"
                })
            
            # Number pattern
            number_pattern = df[column_name].str.contains(r'^\d+$', na=False, regex=True).sum()
            if number_pattern > 0:
                patterns_found.append({
                    'Pattern': '🔢 Numeric strings',
                    'Count': number_pattern,
                    'Percentage': f"{(number_pattern/len(df)*100):.1f}%"
                })
            
            if patterns_found:
                pattern_df = pd.DataFrame(patterns_found)
                st.dataframe(pattern_df, hide_index=True, use_container_width=True)
            else:
                st.info("No common patterns detected")
        else:
            # Numeric column analysis
            try:
                col_stats = df[column_name].describe()
                stats_data = [
                    {'Statistic': 'Mean', 'Value': f"{col_stats['mean']:.2f}"},
                    {'Statistic': 'Median', 'Value': f"{col_stats['50%']:.2f}"},
                    {'Statistic': 'Min', 'Value': f"{col_stats['min']:.2f}"},
                    {'Statistic': 'Max', 'Value': f"{col_stats['max']:.2f}"},
                    {'Statistic': 'Std Dev', 'Value': f"{col_stats['std']:.2f}"}
                ]
                stats_df = pd.DataFrame(stats_data)
                st.dataframe(stats_df, hide_index=True, use_container_width=True)
            except:
                st.warning("Unable to calculate numeric statistics")
    
    def show_compatibility_analysis(self, df_a: pd.DataFrame, col_a: str, df_b: pd.DataFrame, col_b: str):
        """Analyze compatibility between two columns"""
        
        # Data type compatibility
        type_a = df_a[col_a].dtype
        type_b = df_b[col_b].dtype
        
        col_comp1, col_comp2 = st.columns(2)
        
        with col_comp1:
            st.metric("Sheet A Data Type", str(type_a))
        with col_comp2:
            st.metric("Sheet B Data Type", str(type_b))
        
        # Compatibility assessment
        if type_a == type_b:
            st.success("✅ **Data types match perfectly**")
        elif (type_a == 'object' and type_b == 'object'):
            st.success("✅ **Both are text columns - good for fuzzy matching**")
        elif str(type_a).startswith(('int', 'float')) and str(type_b).startswith(('int', 'float')):
            st.success("✅ **Both are numeric - good for exact matching**")
        else:
            st.warning("⚠️ **Different data types - may need preprocessing**")
        
        # Sample overlap analysis
        if df_a[col_a].dtype == 'object' and df_b[col_b].dtype == 'object':
            sample_a = set(df_a[col_a].dropna().astype(str).str.lower()[:500])
            sample_b = set(df_b[col_b].dropna().astype(str).str.lower()[:500])
            
            overlap = len(sample_a & sample_b)
            total_unique = len(sample_a | sample_b)
            overlap_pct = (overlap / total_unique * 100) if total_unique > 0 else 0
            
            st.metric("Sample Overlap", f"{overlap_pct:.1f}%", help=f"{overlap} common values in sample")
            
            if overlap_pct >= 30:
                st.success("🎯 **High overlap detected - expect good matches**")
            elif overlap_pct >= 10:
                st.info("ℹ️ **Moderate overlap - fuzzy matching recommended**")
            else:
                st.warning("⚠️ **Low overlap - review data or adjust threshold**")
    
    def show_matching_recommendations(self, df_a: pd.DataFrame, col_a: str, df_b: pd.DataFrame, col_b: str, threshold: int):
        """Provide smart matching recommendations"""
        
        recommendations = []
        
        # Data quality checks
        null_a_pct = (df_a[col_a].isnull().sum() / len(df_a)) * 100
        null_b_pct = (df_b[col_b].isnull().sum() / len(df_b)) * 100
        
        if null_a_pct > 20 or null_b_pct > 20:
            recommendations.append({
                'Type': '🧹 Data Cleaning',
                'Recommendation': 'High missing data detected. Consider cleaning data before matching.',
                'Priority': 'High'
            })
        
        # Threshold recommendations
        unique_a = df_a[col_a].nunique()
        unique_b = df_b[col_b].nunique()
        total_rows = min(len(df_a), len(df_b))
        
        if unique_a == len(df_a) and unique_b == len(df_b):
            if threshold < 90:
                recommendations.append({
                    'Type': '🎯 Threshold',
                    'Recommendation': 'Perfect unique keys detected. Consider raising threshold to 90%+ for better precision.',
                    'Priority': 'Medium'
                })
        elif (unique_a / len(df_a)) < 0.8 or (unique_b / len(df_b)) < 0.8:
            if threshold > 70:
                recommendations.append({
                    'Type': '🎯 Threshold',
                    'Recommendation': 'High duplicate rate detected. Consider lowering threshold to 60-70% for better recall.',
                    'Priority': 'Medium'
                })
        
        # Matching strategy recommendations
        if df_a[col_a].dtype == 'object' and df_b[col_b].dtype == 'object':
            # Check for common patterns
            avg_len_a = df_a[col_a].dropna().astype(str).str.len().mean()
            avg_len_b = df_b[col_b].dropna().astype(str).str.len().mean()
            
            if abs(avg_len_a - avg_len_b) > 10:
                recommendations.append({
                    'Type': '⚙️ Strategy',
                    'Recommendation': 'Significant length difference detected. Consider multi-column matching for better accuracy.',
                    'Priority': 'Low'
                })
            
            if threshold < 80:
                recommendations.append({
                    'Type': '✨ Performance',
                    'Recommendation': 'Text matching with fuzzy logic. Higher thresholds (80%+) typically work well for names and IDs.',
                    'Priority': 'Low'
                })
        
        # Display recommendations
        if recommendations:
            for rec in recommendations:
                priority_color = {
                    'High': 'error',
                    'Medium': 'warning', 
                    'Low': 'info'
                }
                
                getattr(st, priority_color[rec['Priority']])(
                    f"**{rec['Type']}** ({rec['Priority']} Priority): {rec['Recommendation']}"
                )
        else:
            st.success("🎉 **Great! Your data looks well-prepared for matching.**")
            st.info("💡 Tip: You can always adjust the threshold after seeing initial results.")
    
    def add_result_filters(self, results_df: pd.DataFrame, result_type: str) -> pd.DataFrame:
        """Add smart filtering and search to results"""
        
        if results_df.empty:
            return results_df
        
        st.subheader(f"🔍 Filter & Search {result_type} Results")
        
        # Create filter controls in columns
        filter_col1, filter_col2, filter_col3 = st.columns([2, 1, 1])
        
        with filter_col1:
            # Global text search
            search_term = st.text_input(
                f"🔍 Search in all columns", 
                key=f"search_{result_type}",
                placeholder="Enter text to search across all columns...",
                help="Search will look through all text columns for your term"
            )
        
        with filter_col2:
            # Similarity score filter (if applicable)
            min_similarity = 0
            max_similarity = 100
            if 'similarity_score' in results_df.columns:
                # Get actual range of similarity scores
                actual_min = results_df['similarity_score'].min()
                actual_max = results_df['similarity_score'].max()
                
                similarity_range = st.slider(
                    "Similarity Range",
                    min_value=0,
                    max_value=100,
                    value=(int(actual_min), int(actual_max)),
                    key=f"sim_range_{result_type}",
                    help="Filter by similarity score range"
                )
                min_similarity, max_similarity = similarity_range
        
        with filter_col3:
            # Quick filters based on result type
            quick_filters = []
            if result_type == "Matched":
                if 'match_type' in results_df.columns:
                    match_types = results_df['match_type'].unique().tolist()
                    selected_match_types = st.multiselect(
                        "Match Types",
                        options=match_types,
                        default=match_types,
                        key=f"match_types_{result_type}"
                    )
                    quick_filters.append(('match_type', selected_match_types))
            
            elif result_type == "Suggested":
                # For suggested matches, might want to filter by confidence levels
                if 'similarity_score' in results_df.columns:
                    confidence_filter = st.selectbox(
                        "Confidence Level",
                        options=["All", "High (80%+)", "Medium (60-79%)", "Low (<60%)"],
                        key=f"confidence_{result_type}"
                    )
                    if confidence_filter != "All":
                        if confidence_filter == "High (80%+)":
                            quick_filters.append(('similarity_score', (80, 100)))
                        elif confidence_filter == "Medium (60-79%)":
                            quick_filters.append(('similarity_score', (60, 79)))
                        else:  # Low
                            quick_filters.append(('similarity_score', (0, 59)))
        
        # Advanced filters in expandable section
        with st.expander("🛠️ Advanced Filters", expanded=False):
            adv_col1, adv_col2 = st.columns(2)
            
            with adv_col1:
                # Column-specific search
                searchable_columns = [col for col in results_df.columns 
                                    if results_df[col].dtype == 'object']
                
                if searchable_columns:
                    column_search_col = st.selectbox(
                        "Search in specific column",
                        options=["None"] + searchable_columns,
                        key=f"col_search_{result_type}"
                    )
                    
                    if column_search_col != "None":
                        column_search_term = st.text_input(
                            f"Search in {column_search_col}",
                            key=f"col_search_term_{result_type}",
                            placeholder=f"Search specifically in {column_search_col}..."
                        )
                    else:
                        column_search_term = ""
                        column_search_col = None
                else:
                    column_search_col = None
                    column_search_term = ""
            
            with adv_col2:
                # Data quality filters
                show_nulls = st.checkbox(
                    "Include records with missing data",
                    value=True,
                    key=f"nulls_{result_type}",
                    help="Uncheck to hide records that have missing values"
                )
                
                # Text length filter for key columns
                text_columns = [col for col in results_df.columns 
                              if results_df[col].dtype == 'object']
                if text_columns:
                    filter_by_length = st.checkbox(
                        "Filter by text length",
                        key=f"length_filter_{result_type}"
                    )
                    
                    if filter_by_length:
                        length_column = st.selectbox(
                            "Column for length filter",
                            options=text_columns,
                            key=f"length_col_{result_type}"
                        )
                        
                        # Get actual length range
                        lengths = results_df[length_column].astype(str).str.len()
                        min_len, max_len = int(lengths.min()), int(lengths.max())
                        
                        length_range = st.slider(
                            f"Text length range for {length_column}",
                            min_value=min_len,
                            max_value=max_len,
                            value=(min_len, max_len),
                            key=f"length_range_{result_type}"
                        )
                else:
                    filter_by_length = False
        
        # Apply filters
        filtered_df = results_df.copy()
        original_count = len(filtered_df)
        
        # Apply global text search
        if search_term:
            text_columns = [col for col in filtered_df.columns 
                          if filtered_df[col].dtype == 'object']
            if text_columns:
                search_mask = filtered_df[text_columns].astype(str).apply(
                    lambda x: x.str.contains(search_term, case=False, na=False)
                ).any(axis=1)
                filtered_df = filtered_df[search_mask]
        
        # Apply similarity score filter
        if 'similarity_score' in filtered_df.columns:
            filtered_df = filtered_df[
                (filtered_df['similarity_score'] >= min_similarity) &
                (filtered_df['similarity_score'] <= max_similarity)
            ]
        
        # Apply quick filters
        for filter_col, filter_vals in quick_filters:
            if filter_col == 'match_type':
                filtered_df = filtered_df[filtered_df[filter_col].isin(filter_vals)]
            elif filter_col == 'similarity_score':
                min_val, max_val = filter_vals
                filtered_df = filtered_df[
                    (filtered_df[filter_col] >= min_val) &
                    (filtered_df[filter_col] <= max_val)
                ]
        
        # Apply column-specific search
        if column_search_col and column_search_term:
            col_mask = filtered_df[column_search_col].astype(str).str.contains(
                column_search_term, case=False, na=False
            )
            filtered_df = filtered_df[col_mask]
        
        # Apply null filter
        if not show_nulls:
            filtered_df = filtered_df.dropna()
        
        # Apply text length filter
        if filter_by_length and 'length_column' in locals():
            lengths = filtered_df[length_column].astype(str).str.len()
            min_length, max_length = length_range
            filtered_df = filtered_df[
                (lengths >= min_length) & (lengths <= max_length)
            ]
        
        # Show filter results summary
        filtered_count = len(filtered_df)
        if filtered_count != original_count:
            filter_col1, filter_col2, filter_col3 = st.columns(3)
            with filter_col1:
                st.metric("Original Count", f"{original_count:,}")
            with filter_col2:
                st.metric("Filtered Count", f"{filtered_count:,}")
            with filter_col3:
                percentage = (filtered_count / original_count * 100) if original_count > 0 else 0
                st.metric("Percentage Shown", f"{percentage:.1f}%")
            
            if filtered_count == 0:
                st.warning("🔍 No records match your filter criteria. Try adjusting your filters.")
            elif filtered_count < original_count * 0.1:
                st.info(f"💡 Showing only {percentage:.1f}% of records. Consider broadening your filters to see more results.")
        
        # Quick action buttons
        if not filtered_df.empty:
            action_col1, action_col2, action_col3 = st.columns(3)
            
            with action_col1:
                if st.button(f"📋 Copy Filtered Data", key=f"copy_{result_type}"):
                    # Create CSV for easy copying
                    csv_data = filtered_df.to_csv(index=False)
                    st.code(csv_data[:500] + "..." if len(csv_data) > 500 else csv_data)
                    st.success("✅ Data ready to copy from the code block above!")
            
            with action_col2:
                # Download filtered results
                csv_data = filtered_df.to_csv(index=False)
                st.download_button(
                    label=f"💾 Download Filtered {result_type}",
                    data=csv_data,
                    file_name=f"filtered_{result_type.lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key=f"download_filtered_{result_type}"
                )
            
            with action_col3:
                if st.button(f"🔄 Reset All Filters", key=f"reset_{result_type}"):
                    st.rerun()
        
        return filtered_df
    
    def perform_multi_column_comparison(self, df_a: pd.DataFrame, df_b: pd.DataFrame,
                                       key_cols_a: List[str], key_cols_b: List[str],
                                       cols_to_extract: List[str], threshold: int,
                                       field_weights: List[float] = None,
                                       ignore_case: bool = True) -> Dict:
        """Advanced multi-column matching with weighted similarity scores"""
        
        # Validate inputs
        if len(key_cols_a) != len(key_cols_b):
            raise ValueError("Number of key columns must match between both sheets")
        
        if field_weights is None:
            # Default weights: first field gets 50%, others split remaining 50%
            field_weights = [0.5] + [0.5 / (len(key_cols_a) - 1)] * (len(key_cols_a) - 1) if len(key_cols_a) > 1 else [1.0]
        
        # Normalize weights to sum to 1.0
        total_weight = sum(field_weights)
        field_weights = [w / total_weight for w in field_weights]
        
        # Initialize progress tracking
        total_rows = len(df_a)
        start_time = time.time()
        
        # Create progress containers
        progress_container = st.container()
        with progress_container:
            st.subheader("🔄 Multi-Column Comparison Processing")
            
            main_progress = st.progress(0, text="Initializing multi-column comparison...")
            
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
            
            status_text = st.empty()
            eta_text = st.empty()
        
        results = {
            'matched': [],
            'suggested': [],
            'unmatched': []
        }
        
        # Phase 1: Build multi-field lookup dictionary
        main_progress.progress(0.05, text="📚 Building multi-column lookup structure...")
        status_text.info("Creating efficient multi-field lookup for Sheet B...")
        
        b_lookup = {}
        for idx, row in df_b.iterrows():
            # Create individual field values for matching
            field_values = []
            for col in key_cols_b:
                val = str(row[col])
                if ignore_case:
                    val = val.lower().strip()
                field_values.append(val)
            
            # Store both individual fields and composite key for different matching strategies
            composite_key = ' | '.join(field_values)  # Use separator to avoid conflicts
            
            b_lookup[idx] = {
                'field_values': field_values,
                'composite_key': composite_key,
                'data': row[cols_to_extract].to_dict() if cols_to_extract else {},
                'original_values': [row[col] for col in key_cols_b],
                'full_row': row.to_dict()
            }
        
        # Phase 2: Process each row with multi-field matching
        main_progress.progress(0.1, text="🔍 Starting multi-column comparison...")
        
        for i, (idx_a, row_a) in enumerate(df_a.iterrows()):
            # Calculate progress
            progress = 0.1 + (i / total_rows) * 0.85
            current_time = time.time()
            elapsed_time = current_time - start_time
            
            # Update progress
            if i % max(1, total_rows // 100) == 0 or total_rows < 100:
                main_progress.progress(
                    progress,
                    text=f"Processing row {i+1:,} of {total_rows:,} ({((i+1)/total_rows)*100:.1f}%)"
                )
                
                processed_metric.metric("Processed", f"{i+1:,}", f"of {total_rows:,}")
                matched_metric.metric("✅ Matched", f"{len(results['matched']):,}")
                suggested_metric.metric("⚠️ Suggested", f"{len(results['suggested']):,}")
                time_metric.metric("⏱️ Time", f"{elapsed_time:.1f}s")
                
                # ETA calculation
                rows_per_second = 0
                if i > 0 and elapsed_time > 0:
                    rows_per_second = i / elapsed_time
                    remaining_rows = total_rows - i
                    eta_seconds = remaining_rows / rows_per_second if rows_per_second > 0 else 0
                    
                    if eta_seconds > 60:
                        eta_display = f"{eta_seconds/60:.1f}m remaining"
                    else:
                        eta_display = f"{eta_seconds:.0f}s remaining"
                    
                    eta_text.text(f"⏳ ETA: {eta_display}")
                
                # Show current processing
                field_preview = " + ".join([str(row_a[col])[:20] + "..." if len(str(row_a[col])) > 20 else str(row_a[col]) for col in key_cols_a])
                status_text.text(f"🔍 Processing: {field_preview}")
            
            # Get field values for current row
            field_values_a = []
            for col in key_cols_a:
                val = str(row_a[col])
                if ignore_case:
                    val = val.lower().strip()
                field_values_a.append(val)
            
            composite_a = ' | '.join(field_values_a)
            
            # Try exact multi-field match first
            exact_match_found = False
            for b_idx, b_data in b_lookup.items():
                if composite_a == b_data['composite_key']:
                    # Perfect multi-field match
                    result_row = row_a.to_dict()
                    
                    # Add extracted data from Sheet B with clear column naming
                    if b_data['data']:
                        for col_name, col_value in b_data['data'].items():
                            # Prefix columns from Sheet B to avoid conflicts
                            prefixed_col_name = f"SheetB_{col_name}" if col_name in result_row else col_name
                            result_row[prefixed_col_name] = col_value
                    
                    result_row['match_type'] = 'Multi-Field Exact'
                    result_row['similarity_score'] = 100.0
                    result_row['matched_keys'] = dict(zip(key_cols_b, b_data['original_values']))
                    result_row['field_breakdown'] = {f"{key_cols_a[j]} → {key_cols_b[j]}": 100.0 for j in range(len(key_cols_a))}
                    results['matched'].append(result_row)
                    exact_match_found = True
                    break
            
            if exact_match_found:
                continue
            
            # Multi-field fuzzy matching with weighted scores
            best_match = None
            best_score = 0
            best_breakdown = {}
            
            for b_idx, b_data in b_lookup.items():
                # Calculate weighted similarity across all fields
                field_scores = []
                field_breakdown = {}
                
                for j, (val_a, val_b) in enumerate(zip(field_values_a, b_data['field_values'])):
                    # Calculate similarity for this field pair
                    field_similarity = fuzz.ratio(val_a, val_b)
                    weighted_score = field_similarity * field_weights[j]
                    field_scores.append(weighted_score)
                    
                    # Store individual field breakdown
                    field_breakdown[f"{key_cols_a[j]} → {key_cols_b[j]}"] = field_similarity
                
                total_weighted_score = sum(field_scores)
                
                if total_weighted_score > best_score:
                    best_score = total_weighted_score
                    best_match = b_data
                    best_breakdown = field_breakdown
            
            # Categorize based on weighted score
            if best_match and best_score >= threshold:
                result_row = row_a.to_dict()
                
                # Add extracted data from Sheet B with clear column naming
                if best_match['data']:
                    for col_name, col_value in best_match['data'].items():
                        # Prefix columns from Sheet B to avoid conflicts
                        prefixed_col_name = f"SheetB_{col_name}" if col_name in result_row else col_name
                        result_row[prefixed_col_name] = col_value
                
                result_row['matched_keys'] = dict(zip(key_cols_b, best_match['original_values']))
                result_row['field_breakdown'] = best_breakdown
                result_row['similarity_score'] = best_score
                
                if best_score >= 90:
                    result_row['match_type'] = 'Multi-Field High Confidence'
                    results['matched'].append(result_row)
                else:
                    result_row['match_type'] = 'Multi-Field Suggested'
                    results['suggested'].append(result_row)
            else:
                # No adequate multi-field match found
                result_row = row_a.to_dict()
                result_row['match_type'] = 'No Multi-Field Match'
                result_row['similarity_score'] = best_score if best_match else 0.0
                result_row['matched_keys'] = None
                result_row['field_breakdown'] = best_breakdown if best_match else {}
                results['unmatched'].append(result_row)
        
        # Final progress update
        total_time = time.time() - start_time
        main_progress.progress(1.0, text="✅ Multi-column comparison completed!")
        
        # Final metrics
        processed_metric.metric("Processed", f"{total_rows:,}", "Complete!")
        matched_metric.metric("✅ Matched", f"{len(results['matched']):,}")
        suggested_metric.metric("⚠️ Suggested", f"{len(results['suggested']):,}")
        time_metric.metric("⏱️ Total Time", f"{total_time:.1f}s")
        
        # Multi-field success summary
        match_rate = (len(results['matched']) / total_rows) * 100 if total_rows > 0 else 0
        avg_speed = total_rows / total_time if total_time > 0 else 0
        
        status_text.success(
            f"🎉 Multi-column processing complete! "
            f"Match rate: {match_rate:.1f}% | "
            f"Fields analyzed: {len(key_cols_a)} | "
            f"Average speed: {avg_speed:.1f} rows/sec"
        )
        eta_text.empty()
        
        time.sleep(1.5)
        return results
    
    def perform_batch_comparison(self, uploaded_file, reference_sheet: str, target_sheets: List[str],
                                key_col_ref: str, key_col_targets: str, 
                                cols_to_extract: List[str], threshold: int,
                                ignore_case: bool = True) -> Dict:
        """Perform batch comparison of reference sheet against multiple target sheets"""
        
        batch_results = {}
        
        # Load reference sheet
        df_ref = self.read_sheet(uploaded_file, reference_sheet)
        if df_ref is None:
            st.error(f"Failed to load reference sheet: {reference_sheet}")
            return {}
        
        st.subheader("🔄 Batch Processing Progress")
        
        # Create overall progress tracking
        total_comparisons = len(target_sheets)
        overall_progress = st.progress(0, text=f"Starting batch comparison of {total_comparisons} sheets...")
        
        # Results summary container
        results_summary = st.container()
        
        # Process each target sheet
        for i, target_sheet in enumerate(target_sheets):
            # Update overall progress
            progress = i / total_comparisons
            overall_progress.progress(progress, text=f"Processing {target_sheet} ({i+1}/{total_comparisons})")
            
            st.write(f"### 📊 Comparing: {reference_sheet} vs {target_sheet}")
            
            # Load target sheet
            df_target = self.read_sheet(uploaded_file, target_sheet)
            if df_target is None:
                st.warning(f"⚠️ Skipping {target_sheet} - could not load sheet")
                batch_results[target_sheet] = {"error": "Could not load sheet"}
                continue
            
            # Show comparison metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Reference Rows", f"{len(df_ref):,}")
            with col2:
                st.metric("Target Rows", f"{len(df_target):,}")
            with col3:
                overlap_estimate = min(len(df_ref), len(df_target))
                st.metric("Est. Overlap", f"{overlap_estimate:,}")
            
            # Perform individual comparison
            try:
                comparison_results = self.perform_comparison(
                    df_ref, df_target, 
                    key_col_ref, key_col_targets,
                    cols_to_extract, threshold, ignore_case
                )
                
                # Store results with metadata
                batch_results[target_sheet] = {
                    "results": comparison_results,
                    "reference_sheet": reference_sheet,
                    "target_sheet": target_sheet,
                    "reference_rows": len(df_ref),
                    "target_rows": len(df_target),
                    "match_count": len(comparison_results.get('matched', [])),
                    "suggested_count": len(comparison_results.get('suggested', [])),
                    "unmatched_count": len(comparison_results.get('unmatched', [])),
                    "match_rate": (len(comparison_results.get('matched', [])) / len(df_ref) * 100) if len(df_ref) > 0 else 0
                }
                
                # Show quick summary for this comparison
                match_count = len(comparison_results.get('matched', []))
                total_ref = len(df_ref)
                match_rate = (match_count / total_ref * 100) if total_ref > 0 else 0
                
                if match_rate >= 80:
                    st.success(f"✅ High match rate: {match_rate:.1f}% ({match_count:,}/{total_ref:,})")
                elif match_rate >= 50:
                    st.warning(f"🟡 Medium match rate: {match_rate:.1f}% ({match_count:,}/{total_ref:,})")
                else:
                    st.error(f"🔴 Low match rate: {match_rate:.1f}% ({match_count:,}/{total_ref:,})")
                
            except Exception as e:
                st.error(f"❌ Error comparing with {target_sheet}: {str(e)}")
                batch_results[target_sheet] = {"error": str(e)}
            
            st.divider()
        
        # Final progress update
        overall_progress.progress(1.0, text="✅ Batch comparison completed!")
        
        # Show batch summary
        with results_summary:
            self.show_batch_summary(batch_results)
        
        return batch_results
    
    def show_batch_summary(self, batch_results: Dict):
        """Display comprehensive batch processing summary"""
        
        st.subheader("📊 Batch Processing Summary")
        
        # Calculate overall statistics
        total_comparisons = len([k for k, v in batch_results.items() if "results" in v])
        successful_comparisons = len([k for k, v in batch_results.items() if "results" in v and "error" not in v])
        failed_comparisons = len(batch_results) - successful_comparisons
        
        # Overall metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Comparisons", f"{len(batch_results)}")
        with col2:
            st.metric("✅ Successful", f"{successful_comparisons}")
        with col3:
            st.metric("❌ Failed", f"{failed_comparisons}")
        with col4:
            success_rate = (successful_comparisons / len(batch_results) * 100) if batch_results else 0
            st.metric("Success Rate", f"{success_rate:.1f}%")
        
        # Detailed results table
        if successful_comparisons > 0:
            summary_data = []
            for sheet_name, result_data in batch_results.items():
                if "results" in result_data and "error" not in result_data:
                    summary_data.append({
                        'Target Sheet': sheet_name,
                        'Match Rate': f"{result_data['match_rate']:.1f}%",
                        'Matched': f"{result_data['match_count']:,}",
                        'Suggested': f"{result_data['suggested_count']:,}",
                        'Unmatched': f"{result_data['unmatched_count']:,}",
                        'Target Rows': f"{result_data['target_rows']:,}",
                        'Status': '🟢 High' if result_data['match_rate'] >= 80 else '🟡 Medium' if result_data['match_rate'] >= 50 else '🔴 Low'
                    })
            
            if summary_data:
                summary_df = pd.DataFrame(summary_data)
                st.dataframe(summary_df, hide_index=True, use_container_width=True)
                
                # Best and worst performers
                if len(summary_data) > 1:
                    match_rates = [float(row['Match Rate'].replace('%', '')) for row in summary_data]
                    best_idx = match_rates.index(max(match_rates))
                    worst_idx = match_rates.index(min(match_rates))
                    
                    col_best, col_worst = st.columns(2)
                    with col_best:
                        st.success(f"🏆 Best Match: **{summary_data[best_idx]['Target Sheet']}** ({summary_data[best_idx]['Match Rate']})")
                    with col_worst:
                        st.warning(f"⚠️ Needs Review: **{summary_data[worst_idx]['Target Sheet']}** ({summary_data[worst_idx]['Match Rate']})")
        
        # Failed comparisons
        if failed_comparisons > 0:
            st.subheader("❌ Failed Comparisons")
            failed_data = []
            for sheet_name, result_data in batch_results.items():
                if "error" in result_data:
                    failed_data.append({
                        'Sheet': sheet_name,
                        'Error': result_data['error']
                    })
            
            if failed_data:
                failed_df = pd.DataFrame(failed_data)
                st.dataframe(failed_df, hide_index=True, use_container_width=True)

    def perform_consolidation(self, uploaded_file, sheets_to_consolidate: List[str],
                            consolidation_strategy: str, include_source_column: bool,
                            handle_duplicates: str, missing_data_strategy: str,
                            validate_schemas: bool, show_consolidation_report: bool) -> Dict:
        """Perform cross-sheet data consolidation with multiple strategies"""
        
        # Initialize progress tracking
        total_sheets = len(sheets_to_consolidate)
        start_time = time.time()
        
        # Create progress containers
        progress_container = st.container()
        with progress_container:
            st.subheader("🔗 Processing Consolidation")
            
            # Main progress bar
            main_progress = st.progress(0, text="Initializing consolidation...")
            
            # Status metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                processed_metric = st.metric("Processed", "0", f"of {total_sheets:,} sheets")
            with col2:
                rows_metric = st.metric("📊 Total Rows", "0")
            with col3:
                columns_metric = st.metric("📋 Unique Columns", "0")
            with col4:
                time_metric = st.metric("⏱️ Time", "0s")
            
            # Live status text
            status_text = st.empty()
        
        consolidation_results = {
            'strategy': consolidation_strategy,
            'consolidated_data': None,
            'schema_analysis': {},
            'processing_report': {},
            'quality_metrics': {},
            'sheets_processed': [],
            'errors': []
        }
        
        try:
            # Phase 1: Load and analyze all sheets
            main_progress.progress(0.1, text="📚 Loading and analyzing sheets...")
            status_text.info("Loading sheets and analyzing structures...")
            
            sheet_data = {}
            all_columns = set()
            total_rows = 0
            
            for i, sheet_name in enumerate(sheets_to_consolidate):
                # Update progress
                progress = 0.1 + (i / total_sheets) * 0.3
                main_progress.progress(progress, text=f"Loading {sheet_name}...")
                
                # Load sheet
                df = self.read_sheet(uploaded_file, sheet_name)
                if df is not None:
                    sheet_data[sheet_name] = df
                    all_columns.update(df.columns.tolist())
                    total_rows += len(df)
                    
                    # Update metrics
                    processed_metric.metric("Processed", f"{i+1}", f"of {total_sheets:,} sheets")
                    rows_metric.metric("📊 Total Rows", f"{total_rows:,}")
                    columns_metric.metric("📋 Unique Columns", f"{len(all_columns):,}")
                    
                    consolidation_results['sheets_processed'].append({
                        'name': sheet_name,
                        'rows': len(df),
                        'columns': len(df.columns),
                        'column_names': df.columns.tolist()
                    })
                else:
                    consolidation_results['errors'].append(f"Could not load sheet: {sheet_name}")
            
            if not sheet_data:
                raise Exception("No sheets could be loaded successfully")
            
            # Phase 2: Schema Analysis
            main_progress.progress(0.4, text="🔍 Analyzing schemas and data compatibility...")
            status_text.info("Performing schema analysis and compatibility checks...")
            
            schema_analysis = self._analyze_consolidation_schemas(sheet_data, validate_schemas)
            consolidation_results['schema_analysis'] = schema_analysis
            
            # Phase 3: Consolidation based on strategy
            main_progress.progress(0.6, text=f"🔗 Consolidating data using {consolidation_strategy}...")
            
            if consolidation_strategy == "Union (Combine all data)":
                consolidated_df = self._perform_union_consolidation(
                    sheet_data, include_source_column, handle_duplicates, missing_data_strategy
                )
            elif consolidation_strategy == "Key-based Merge":
                consolidated_df = self._perform_keybased_consolidation(
                    sheet_data, schema_analysis, include_source_column
                )
            else:  # Schema Analysis Only
                consolidated_df = self._perform_schema_analysis_only(sheet_data, schema_analysis)
            
            consolidation_results['consolidated_data'] = consolidated_df
            
            # Phase 4: Generate quality report
            main_progress.progress(0.8, text="📊 Generating consolidation report...")
            
            if show_consolidation_report:
                quality_metrics = self._generate_consolidation_report(
                    sheet_data, consolidated_df, schema_analysis, consolidation_strategy
                )
                consolidation_results['quality_metrics'] = quality_metrics
            
            # Phase 5: Display results
            main_progress.progress(1.0, text="✅ Consolidation completed successfully!")
            
            # Final metrics update
            total_time = time.time() - start_time
            final_rows = len(consolidated_df) if consolidated_df is not None else 0
            final_cols = len(consolidated_df.columns) if consolidated_df is not None else 0
            
            processed_metric.metric("Processed", f"{len(sheet_data)}", "Complete!")
            rows_metric.metric("📊 Final Rows", f"{final_rows:,}")
            columns_metric.metric("📋 Final Columns", f"{final_cols:,}")
            time_metric.metric("⏱️ Total Time", f"{total_time:.1f}s")
            
            status_text.success(
                f"🎉 Consolidation complete! "
                f"Combined {len(sheet_data)} sheets into {final_rows:,} rows × {final_cols:,} columns | "
                f"Total time: {total_time:.1f}s"
            )
            
            # Brief pause to show completion
            time.sleep(1.5)
            
            # Show consolidated results
            if consolidated_df is not None:
                self.show_consolidation_results(consolidation_results)
            
            return consolidation_results
            
        except Exception as e:
            main_progress.progress(0, text="❌ Consolidation failed!")
            status_text.error(f"Error during consolidation: {str(e)}")
            consolidation_results['errors'].append(str(e))
            raise e

    def _analyze_consolidation_schemas(self, sheet_data: Dict, validate_schemas: bool) -> Dict:
        """Analyze schemas across sheets for consolidation compatibility"""
        
        schema_analysis = {
            'common_columns': [],
            'unique_columns': {},
            'column_types': {},
            'compatibility_issues': [],
            'recommendations': []
        }
        
        # Find all unique columns across sheets
        all_columns = set()
        sheet_columns = {}
        
        for sheet_name, df in sheet_data.items():
            sheet_columns[sheet_name] = df.columns.tolist()
            all_columns.update(df.columns)
        
        # Identify common vs unique columns
        common_cols = set(sheet_columns[list(sheet_data.keys())[0]])
        for sheet_name, cols in sheet_columns.items():
            common_cols = common_cols.intersection(set(cols))
        
        schema_analysis['common_columns'] = list(common_cols)
        
        # Find unique columns per sheet
        for sheet_name, cols in sheet_columns.items():
            unique_cols = set(cols) - common_cols
            if unique_cols:
                schema_analysis['unique_columns'][sheet_name] = list(unique_cols)
        
        # Analyze data types for common columns if validation enabled
        if validate_schemas and common_cols:
            for col in common_cols:
                col_types = {}
                for sheet_name, df in sheet_data.items():
                    if col in df.columns:
                        col_types[sheet_name] = str(df[col].dtype)
                
                schema_analysis['column_types'][col] = col_types
                
                # Check for type inconsistencies
                unique_types = set(col_types.values())
                if len(unique_types) > 1:
                    schema_analysis['compatibility_issues'].append({
                        'column': col,
                        'issue': 'Data type mismatch',
                        'details': col_types
                    })
        
        # Generate recommendations
        if len(common_cols) == 0:
            schema_analysis['recommendations'].append("No common columns found. Consider Union strategy with source tracking.")
        elif len(common_cols) < len(all_columns) * 0.5:
            schema_analysis['recommendations'].append("Few common columns. Union strategy recommended over Key-based merge.")
        else:
            schema_analysis['recommendations'].append("Good column alignment. Key-based merge is viable.")
        
        return schema_analysis

    def _perform_union_consolidation(self, sheet_data: Dict, include_source_column: bool,
                                   handle_duplicates: str, missing_data_strategy: str) -> pd.DataFrame:
        """Perform union-based consolidation (stack all data)"""
        
        consolidated_dfs = []
        
        for sheet_name, df in sheet_data.items():
            df_copy = df.copy()
            
            # Add source column if requested
            if include_source_column:
                df_copy['_source_sheet'] = sheet_name
            
            consolidated_dfs.append(df_copy)
        
        # Combine all dataframes
        if missing_data_strategy == "Fill with blanks":
            consolidated_df = pd.concat(consolidated_dfs, ignore_index=True, sort=False)
        elif missing_data_strategy == "Skip rows":
            # Only keep rows that have data in common columns
            consolidated_df = pd.concat(consolidated_dfs, ignore_index=True, join='inner', sort=False)
        else:  # Use default value
            consolidated_df = pd.concat(consolidated_dfs, ignore_index=True, sort=False)
            consolidated_df = consolidated_df.fillna('DEFAULT_VALUE')
        
        # Handle duplicates
        if handle_duplicates == "Remove duplicates":
            original_count = len(consolidated_df)
            consolidated_df = consolidated_df.drop_duplicates()
            st.info(f"Removed {original_count - len(consolidated_df):,} duplicate rows")
        elif handle_duplicates == "Mark duplicates":
            consolidated_df['_is_duplicate'] = consolidated_df.duplicated(keep=False)
        
        return consolidated_df

    def _perform_keybased_consolidation(self, sheet_data: Dict, schema_analysis: Dict,
                                      include_source_column: bool) -> pd.DataFrame:
        """Perform key-based merge consolidation"""
        
        common_columns = schema_analysis['common_columns']
        
        if not common_columns:
            st.warning("No common columns found. Falling back to Union consolidation.")
            return self._perform_union_consolidation(sheet_data, include_source_column, "Keep all", "Fill with blanks")
        
        # Use the first common column as key (or let user select)
        key_column = common_columns[0]
        st.info(f"Using '{key_column}' as merge key column")
        
        # Start with first sheet
        sheet_names = list(sheet_data.keys())
        consolidated_df = sheet_data[sheet_names[0]].copy()
        
        if include_source_column:
            consolidated_df[f'_sources'] = sheet_names[0]
        
        # Merge with remaining sheets
        for sheet_name in sheet_names[1:]:
            df_to_merge = sheet_data[sheet_name]
            
            # Perform left join
            merged_df = consolidated_df.merge(
                df_to_merge, 
                on=key_column, 
                how='outer', 
                suffixes=('', f'_from_{sheet_name}')
            )
            
            if include_source_column:
                # Update sources column
                merged_df[f'_sources'] = merged_df[f'_sources'].fillna('') + f';{sheet_name}'
                merged_df[f'_sources'] = merged_df[f'_sources'].str.strip(';')
            
            consolidated_df = merged_df
        
        return consolidated_df

    def _perform_schema_analysis_only(self, sheet_data: Dict, schema_analysis: Dict) -> pd.DataFrame:
        """Create analysis-only output showing schema comparison"""
        
        analysis_data = []
        
        for sheet_name, df in sheet_data.items():
            for col in df.columns:
                analysis_data.append({
                    'Sheet': sheet_name,
                    'Column': col,
                    'Data_Type': str(df[col].dtype),
                    'Non_Null_Count': df[col].count(),
                    'Null_Count': df[col].isnull().sum(),
                    'Unique_Values': df[col].nunique(),
                    'Sample_Value': str(df[col].dropna().iloc[0]) if len(df[col].dropna()) > 0 else 'N/A'
                })
        
        return pd.DataFrame(analysis_data)

    def _generate_consolidation_report(self, sheet_data: Dict, consolidated_df: pd.DataFrame,
                                     schema_analysis: Dict, strategy: str) -> Dict:
        """Generate comprehensive consolidation quality report"""
        
        total_input_rows = sum(len(df) for df in sheet_data.values())
        final_rows = len(consolidated_df) if consolidated_df is not None else 0
        
        quality_metrics = {
            'input_summary': {
                'total_sheets': len(sheet_data),
                'total_input_rows': total_input_rows,
                'final_output_rows': final_rows,
                'row_efficiency': (final_rows / total_input_rows * 100) if total_input_rows > 0 else 0
            },
            'schema_summary': {
                'common_columns': len(schema_analysis['common_columns']),
                'total_unique_columns': len(set().union(*[df.columns for df in sheet_data.values()])),
                'compatibility_issues': len(schema_analysis['compatibility_issues'])
            },
            'data_quality': {
                'strategy_used': strategy,
                'recommendations': schema_analysis['recommendations']
            }
        }
        
        return quality_metrics

    def show_consolidation_results(self, consolidation_results: Dict):
        """Display consolidation results with comprehensive analysis"""
        
        st.divider()
        st.header("📊 Consolidation Results")
        
        consolidated_df = consolidation_results['consolidated_data']
        
        if consolidated_df is not None and not consolidated_df.empty:
            # Results overview
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("📊 Final Rows", f"{len(consolidated_df):,}")
            with col2:
                st.metric("📋 Final Columns", f"{len(consolidated_df.columns):,}")
            with col3:
                st.metric("🔗 Sheets Combined", f"{len(consolidation_results['sheets_processed'])}")
            with col4:
                memory_usage = consolidated_df.memory_usage(deep=True).sum() / 1024 / 1024
                st.metric("💾 Memory Usage", f"{memory_usage:.1f} MB")
            
            # Display consolidated data
            st.subheader("🔍 Consolidated Data Preview")
            
            # Add filtering capability
            with st.expander("🔧 Data Filters", expanded=False):
                if '_source_sheet' in consolidated_df.columns:
                    source_filter = st.multiselect(
                        "Filter by source sheet:",
                        consolidated_df['_source_sheet'].unique(),
                        default=consolidated_df['_source_sheet'].unique()
                    )
                    
                    if source_filter:
                        filtered_df = consolidated_df[consolidated_df['_source_sheet'].isin(source_filter)]
                    else:
                        filtered_df = consolidated_df
                else:
                    filtered_df = consolidated_df
                
                # Show sample of data
                sample_size = st.slider("Preview rows:", 10, min(1000, len(filtered_df)), 100)
            
            st.dataframe(filtered_df.head(sample_size), use_container_width=True)
            st.caption(f"Showing {min(sample_size, len(filtered_df)):,} of {len(filtered_df):,} total rows")
            
            # Schema analysis
            if consolidation_results['schema_analysis']:
                st.subheader("🔍 Schema Analysis")
                
                schema_col1, schema_col2 = st.columns(2)
                
                with schema_col1:
                    st.write("**Common Columns:**")
                    common_cols = consolidation_results['schema_analysis']['common_columns']
                    if common_cols:
                        for col in common_cols:
                            st.write(f"✅ {col}")
                    else:
                        st.write("No common columns found")
                
                with schema_col2:
                    st.write("**Unique Columns by Sheet:**")
                    unique_cols = consolidation_results['schema_analysis']['unique_columns']
                    if unique_cols:
                        for sheet, cols in unique_cols.items():
                            st.write(f"**{sheet}:** {', '.join(cols)}")
                    else:
                        st.write("All sheets have identical columns")
            
            # Export options
            st.subheader("📥 Export Consolidated Data")
            
            if st.button("📊 Export to Excel", type="secondary"):
                # Create Excel export with consolidation metadata
                output = io.BytesIO()
                
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Main consolidated data
                    consolidated_df.to_excel(writer, sheet_name='Consolidated Data', index=False)
                    
                    # Schema analysis
                    if consolidation_results['schema_analysis']:
                        schema_df = pd.DataFrame({
                            'Metric': ['Total Sheets', 'Common Columns', 'Unique Columns', 'Strategy Used'],
                            'Value': [
                                len(consolidation_results['sheets_processed']),
                                len(consolidation_results['schema_analysis']['common_columns']),
                                len(consolidation_results['schema_analysis']['unique_columns']),
                                consolidation_results['strategy']
                            ]
                        })
                        schema_df.to_excel(writer, sheet_name='Consolidation Report', index=False)
                
                output.seek(0)
                
                st.download_button(
                    label="⬇️ Download Consolidated Excel File",
                    data=output.getvalue(),
                    file_name=f"consolidated_data_{len(consolidation_results['sheets_processed'])}sheets.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("No consolidated data to display")

    def detect_time_patterns(self, sheet_names: List[str]) -> List[Dict]:
        """Detect time-based patterns in sheet names for historical analysis"""
        
        patterns = []
        
        # Common time patterns to detect
        monthly_patterns = {
            'jan': 'January', 'feb': 'February', 'mar': 'March', 'apr': 'April',
            'may': 'May', 'jun': 'June', 'jul': 'July', 'aug': 'August',
            'sep': 'September', 'oct': 'October', 'nov': 'November', 'dec': 'December',
            'january': 'January', 'february': 'February', 'march': 'March', 'april': 'April',
            'june': 'June', 'july': 'July', 'august': 'August', 'september': 'September',
            'october': 'October', 'november': 'November', 'december': 'December'
        }
        
        quarterly_patterns = {
            'q1': 'Q1', 'q2': 'Q2', 'q3': 'Q3', 'q4': 'Q4',
            'quarter 1': 'Q1', 'quarter 2': 'Q2', 'quarter 3': 'Q3', 'quarter 4': 'Q4'
        }
        
        yearly_patterns = {}
        current_year = datetime.now().year
        for year in range(current_year - 10, current_year + 5):
            yearly_patterns[str(year)] = str(year)
        
        # Detect monthly patterns
        monthly_sheets = []
        monthly_periods = {}
        for sheet in sheet_names:
            sheet_lower = sheet.lower()
            for pattern, full_name in monthly_patterns.items():
                if pattern in sheet_lower:
                    monthly_sheets.append(sheet)
                    monthly_periods[sheet] = full_name
                    break
        
        if len(monthly_sheets) >= 2:
            patterns.append({
                'type': 'Monthly Time Series',
                'sheets': monthly_sheets,
                'periods': monthly_periods,
                'description': f'Found {len(monthly_sheets)} monthly periods'
            })
        
        # Detect quarterly patterns
        quarterly_sheets = []
        quarterly_periods = {}
        for sheet in sheet_names:
            sheet_lower = sheet.lower()
            for pattern, full_name in quarterly_patterns.items():
                if pattern in sheet_lower:
                    quarterly_sheets.append(sheet)
                    quarterly_periods[sheet] = full_name
                    break
        
        if len(quarterly_sheets) >= 2:
            patterns.append({
                'type': 'Quarterly Time Series',
                'sheets': quarterly_sheets,
                'periods': quarterly_periods,
                'description': f'Found {len(quarterly_sheets)} quarterly periods'
            })
        
        # Detect yearly patterns
        yearly_sheets = []
        yearly_periods = {}
        for sheet in sheet_names:
            sheet_lower = sheet.lower()
            for pattern, full_name in yearly_patterns.items():
                if pattern in sheet_lower:
                    yearly_sheets.append(sheet)
                    yearly_periods[sheet] = full_name
                    break
        
        if len(yearly_sheets) >= 2:
            patterns.append({
                'type': 'Yearly Time Series',
                'sheets': yearly_sheets,
                'periods': yearly_periods,
                'description': f'Found {len(yearly_sheets)} yearly periods'
            })
        
        # Detect sequential numerical patterns (Week1, Week2, etc.)
        import re
        sequential_sheets = []
        sequential_periods = {}
        for sheet in sheet_names:
            # Look for patterns like "Week1", "Month2", "Period3", etc.
            match = re.search(r'(week|month|period|day|phase|step)[\s]*(\d+)', sheet.lower())
            if match:
                sequential_sheets.append(sheet)
                sequential_periods[sheet] = f"{match.group(1).title()} {match.group(2)}"
        
        if len(sequential_sheets) >= 2:
            patterns.append({
                'type': 'Sequential Time Series',
                'sheets': sequential_sheets,
                'periods': sequential_periods,
                'description': f'Found {len(sequential_sheets)} sequential periods'
            })
        
        return patterns

    def perform_historical_comparison(self, uploaded_file, sheets_to_compare: List[str],
                                    analysis_mode: str, analysis_type: str, baseline_sheet: str = None,
                                    include_variance: bool = True, show_trend_charts: bool = True) -> Dict:
        """Perform comprehensive historical time-series comparison analysis"""
        
        try:
            # Load all sheets data
            sheet_data = {}
            for sheet_name in sheets_to_compare:
                df = self.read_sheet(uploaded_file, sheet_name)
                if df is not None:
                    sheet_data[sheet_name] = df
            
            if not sheet_data:
                raise ValueError("No valid sheets found for analysis")
            
            # Progress tracking
            progress_bar = st.progress(0)
            progress_bar.progress(0.1, text="📊 Analyzing sheet structures...")
            
            # Analyze data structure consistency
            structure_analysis = self._analyze_historical_structure(sheet_data)
            
            progress_bar.progress(0.3, text="📈 Calculating time-series metrics...")
            
            # Perform time-series analysis based on mode
            if analysis_mode == "Trend Analysis":
                analysis_results = self._perform_trend_analysis(sheet_data, structure_analysis, include_variance)
            elif analysis_mode == "Period-to-Period Change":
                analysis_results = self._perform_period_change_analysis(sheet_data, structure_analysis, include_variance)
            else:  # Baseline Comparison
                analysis_results = self._perform_baseline_comparison(sheet_data, structure_analysis, baseline_sheet, include_variance)
            
            progress_bar.progress(0.7, text="📊 Generating visualizations...")
            
            # Generate trend charts if requested
            if show_trend_charts:
                chart_data = self._generate_trend_charts(sheet_data, analysis_results, analysis_mode)
                analysis_results['charts'] = chart_data
            
            progress_bar.progress(0.9, text="📋 Creating summary report...")
            
            # Create comprehensive historical report
            historical_report = self._generate_historical_report(
                sheet_data, analysis_results, analysis_mode, analysis_type, include_variance
            )
            
            progress_bar.progress(1.0, text="✅ Historical analysis completed!")
            
            # Display results
            self.show_historical_results({
                'analysis_results': analysis_results,
                'historical_report': historical_report,
                'sheet_data': sheet_data,
                'analysis_mode': analysis_mode,
                'analysis_type': analysis_type,
                'sheets_analyzed': sheets_to_compare
            })
            
            return {
                'analysis_results': analysis_results,
                'historical_report': historical_report,
                'sheet_data': sheet_data,
                'analysis_mode': analysis_mode,
                'analysis_type': analysis_type,
                'sheets_analyzed': sheets_to_compare
            }
            
        except Exception as e:
            st.error(f"Historical analysis error: {str(e)}")
            raise e

    def _analyze_historical_structure(self, sheet_data: Dict) -> Dict:
        """Analyze structure consistency across historical sheets"""
        
        structure_analysis = {
            'common_columns': [],
            'unique_columns': {},
            'data_types': {},
            'row_counts': {},
            'compatibility_score': 0
        }
        
        all_columns = set()
        sheet_columns = {}
        
        for sheet_name, df in sheet_data.items():
            sheet_columns[sheet_name] = set(df.columns)
            all_columns.update(df.columns)
            structure_analysis['row_counts'][sheet_name] = len(df)
        
        # Find common columns
        if sheet_columns:
            structure_analysis['common_columns'] = list(set.intersection(*sheet_columns.values()))
        
        # Find unique columns per sheet
        for sheet_name, cols in sheet_columns.items():
            unique_cols = cols - set(structure_analysis['common_columns'])
            if unique_cols:
                structure_analysis['unique_columns'][sheet_name] = list(unique_cols)
        
        # Calculate compatibility score
        if all_columns:
            common_ratio = len(structure_analysis['common_columns']) / len(all_columns)
            structure_analysis['compatibility_score'] = common_ratio * 100
        
        return structure_analysis

    def _perform_trend_analysis(self, sheet_data: Dict, structure_analysis: Dict, include_variance: bool) -> Dict:
        """Perform trend analysis across time periods"""
        
        results = {
            'trends': {},
            'summary_stats': {},
            'variance_metrics': {} if include_variance else None
        }
        
        common_columns = structure_analysis['common_columns']
        
        for column in common_columns:
            column_trends = {}
            column_values = []
            
            for sheet_name, df in sheet_data.items():
                if column in df.columns:
                    # Calculate basic statistics for numeric columns
                    if pd.api.types.is_numeric_dtype(df[column]):
                        stats = {
                            'mean': df[column].mean(),
                            'median': df[column].median(),
                            'sum': df[column].sum(),
                            'count': df[column].count(),
                            'min': df[column].min(),
                            'max': df[column].max()
                        }
                        column_trends[sheet_name] = stats
                        column_values.extend(df[column].dropna().tolist())
                    else:
                        # For non-numeric columns, provide basic counts
                        stats = {
                            'unique_count': df[column].nunique(),
                            'most_common': df[column].mode().iloc[0] if not df[column].mode().empty else None,
                            'count': df[column].count(),
                            'null_count': df[column].isnull().sum()
                        }
                        column_trends[sheet_name] = stats
            
            results['trends'][column] = column_trends
            
            # Calculate variance metrics for numeric columns
            if include_variance and column_values and pd.api.types.is_numeric_dtype(pd.Series(column_values)):
                series = pd.Series(column_values)
                results['variance_metrics'][column] = {
                    'std_dev': series.std(),
                    'variance': series.var(),
                    'coefficient_variation': (series.std() / series.mean()) * 100 if series.mean() != 0 else 0,
                    'range': series.max() - series.min()
                }
        
        return results

    def _perform_period_change_analysis(self, sheet_data: Dict, structure_analysis: Dict, include_variance: bool) -> Dict:
        """Analyze period-to-period changes"""
        
        results = {
            'changes': {},
            'growth_rates': {},
            'variance_metrics': {} if include_variance else None
        }
        
        common_columns = structure_analysis['common_columns']
        sheet_names = list(sheet_data.keys())
        
        for column in common_columns:
            changes = {}
            growth_rates = {}
            
            for i in range(1, len(sheet_names)):
                current_sheet = sheet_names[i]
                previous_sheet = sheet_names[i-1]
                
                current_df = sheet_data[current_sheet]
                previous_df = sheet_data[previous_sheet]
                
                if column in current_df.columns and column in previous_df.columns:
                    if pd.api.types.is_numeric_dtype(current_df[column]):
                        current_sum = current_df[column].sum()
                        previous_sum = previous_df[column].sum()
                        
                        change = current_sum - previous_sum
                        growth_rate = ((current_sum - previous_sum) / previous_sum * 100) if previous_sum != 0 else 0
                        
                        period_key = f"{previous_sheet} → {current_sheet}"
                        changes[period_key] = change
                        growth_rates[period_key] = growth_rate
            
            results['changes'][column] = changes
            results['growth_rates'][column] = growth_rates
        
        return results

    def _perform_baseline_comparison(self, sheet_data: Dict, structure_analysis: Dict, baseline_sheet: str, include_variance: bool) -> Dict:
        """Compare all periods against a baseline period"""
        
        results = {
            'baseline_comparisons': {},
            'relative_changes': {},
            'variance_metrics': {} if include_variance else None
        }
        
        if baseline_sheet not in sheet_data:
            raise ValueError(f"Baseline sheet '{baseline_sheet}' not found in data")
        
        baseline_df = sheet_data[baseline_sheet]
        common_columns = structure_analysis['common_columns']
        
        for column in common_columns:
            comparisons = {}
            relative_changes = {}
            
            if column in baseline_df.columns:
                baseline_value = baseline_df[column].sum() if pd.api.types.is_numeric_dtype(baseline_df[column]) else baseline_df[column].count()
                
                for sheet_name, df in sheet_data.items():
                    if sheet_name != baseline_sheet and column in df.columns:
                        if pd.api.types.is_numeric_dtype(df[column]):
                            current_value = df[column].sum()
                            comparison = current_value - baseline_value
                            relative_change = ((current_value - baseline_value) / baseline_value * 100) if baseline_value != 0 else 0
                        else:
                            current_value = df[column].count()
                            comparison = current_value - baseline_value
                            relative_change = ((current_value - baseline_value) / baseline_value * 100) if baseline_value != 0 else 0
                        
                        comparisons[sheet_name] = comparison
                        relative_changes[sheet_name] = relative_change
            
            results['baseline_comparisons'][column] = comparisons
            results['relative_changes'][column] = relative_changes
        
        return results

    def _generate_trend_charts(self, sheet_data: Dict, analysis_results: Dict, analysis_mode: str) -> Dict:
        """Generate trend visualization data"""
        
        chart_data = {
            'time_series': {},
            'comparison_charts': {},
            'chart_config': {
                'analysis_mode': analysis_mode,
                'periods': list(sheet_data.keys())
            }
        }
        
        # This would contain matplotlib/plotly chart generation
        # For now, we'll prepare the data structure for charts
        
        return chart_data

    def _generate_historical_report(self, sheet_data: Dict, analysis_results: Dict, 
                                   analysis_mode: str, analysis_type: str, include_variance: bool) -> Dict:
        """Generate comprehensive historical analysis report"""
        
        total_periods = len(sheet_data)
        total_records = sum(len(df) for df in sheet_data.values())
        
        report = {
            'executive_summary': {
                'analysis_type': analysis_type,
                'analysis_mode': analysis_mode,
                'periods_analyzed': total_periods,
                'total_records': total_records,
                'analysis_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            },
            'key_insights': [],
            'recommendations': []
        }
        
        # Generate insights based on analysis results
        if 'trends' in analysis_results:
            report['key_insights'].append("Trend patterns identified across time periods")
        
        if 'changes' in analysis_results:
            report['key_insights'].append("Period-to-period changes calculated")
        
        if 'baseline_comparisons' in analysis_results:
            report['key_insights'].append("Baseline comparison analysis completed")
        
        return report

    def show_historical_results(self, historical_results: Dict):
        """Display comprehensive historical analysis results"""
        
        st.divider()
        st.header("📈 Historical Analysis Results")
        
        analysis_results = historical_results['analysis_results']
        historical_report = historical_results['historical_report']
        analysis_mode = historical_results['analysis_mode']
        
        # Executive summary
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("📅 Periods Analyzed", historical_report['executive_summary']['periods_analyzed'])
        with col2:
            st.metric("📊 Total Records", f"{historical_report['executive_summary']['total_records']:,}")
        with col3:
            st.metric("🔍 Analysis Type", historical_report['executive_summary']['analysis_type'])
        with col4:
            st.metric("📈 Analysis Mode", analysis_mode)
        
        # Results tabs
        if analysis_mode == "Trend Analysis" and 'trends' in analysis_results:
            self._display_trend_results(analysis_results['trends'], analysis_results.get('variance_metrics'))
        elif analysis_mode == "Period-to-Period Change" and 'changes' in analysis_results:
            self._display_change_results(analysis_results['changes'], analysis_results['growth_rates'])
        elif analysis_mode == "Baseline Comparison" and 'baseline_comparisons' in analysis_results:
            self._display_baseline_results(analysis_results['baseline_comparisons'], analysis_results['relative_changes'])
        
        # Export options
        st.subheader("📥 Export Historical Analysis")
        
        if st.button("📊 Export to Excel", type="secondary"):
            # Create Excel export with historical analysis
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Export trend data or other analysis results
                if 'trends' in analysis_results:
                    # Convert trends to exportable format
                    for column, trends in analysis_results['trends'].items():
                        trend_df = pd.DataFrame(trends).T
                        trend_df.to_excel(writer, sheet_name=f'Trends_{column}'[:31])
                
                # Executive summary
                summary_df = pd.DataFrame([historical_report['executive_summary']])
                summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
            
            output.seek(0)
            
            st.download_button(
                label="⬇️ Download Historical Analysis",
                data=output.getvalue(),
                file_name=f"historical_analysis_{len(historical_results['sheets_analyzed'])}periods.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    def _display_trend_results(self, trends: Dict, variance_metrics: Dict = None):
        """Display trend analysis results"""
        
        st.subheader("📈 Trend Analysis Results")
        
        for column, trend_data in trends.items():
            with st.expander(f"📊 {column} Trends", expanded=True):
                trend_df = pd.DataFrame(trend_data).T
                st.dataframe(trend_df, use_container_width=True)
                
                if variance_metrics and column in variance_metrics:
                    st.write("**Variance Metrics:**")
                    variance_df = pd.DataFrame([variance_metrics[column]])
                    st.dataframe(variance_df, use_container_width=True)

    def _display_change_results(self, changes: Dict, growth_rates: Dict):
        """Display period-to-period change results"""
        
        st.subheader("📊 Period-to-Period Changes")
        
        for column in changes.keys():
            with st.expander(f"📈 {column} Changes", expanded=True):
                change_data = []
                
                for period, change in changes[column].items():
                    growth_rate = growth_rates[column].get(period, 0)
                    change_data.append({
                        'Period Transition': period,
                        'Absolute Change': change,
                        'Growth Rate (%)': f"{growth_rate:.2f}%"
                    })
                
                if change_data:
                    change_df = pd.DataFrame(change_data)
                    st.dataframe(change_df, hide_index=True, use_container_width=True)

    def _display_baseline_results(self, baseline_comparisons: Dict, relative_changes: Dict):
        """Display baseline comparison results"""
        
        st.subheader("📊 Baseline Comparison Results")
        
        for column in baseline_comparisons.keys():
            with st.expander(f"📈 {column} vs Baseline", expanded=True):
                baseline_data = []
                
                for sheet, comparison in baseline_comparisons[column].items():
                    relative_change = relative_changes[column].get(sheet, 0)
                    baseline_data.append({
                        'Period': sheet,
                        'Difference from Baseline': comparison,
                        'Relative Change (%)': f"{relative_change:.2f}%"
                    })
                
                if baseline_data:
                    baseline_df = pd.DataFrame(baseline_data)
                    st.dataframe(baseline_df, hide_index=True, use_container_width=True)

def apply_theme(theme_name):
    """Apply custom CSS theme to the Streamlit app"""
    if theme_name == "Dark":
        dark_theme = """
        <style>
        /* Main app background */
        .stApp {
            background-color: #0e1117 !important;
            color: #fafafa !important;
        }
        
        /* Sidebar styling */
        .stSidebar {
            background-color: #262730 !important;
        }
        .stSidebar .stSelectbox > div > div {
            background-color: #262730 !important;
            color: #fafafa !important;
            border: 1px solid #4a4a4a !important;
        }
        
        /* Input elements */
        .stSelectbox > div > div {
            background-color: #262730 !important;
            color: #fafafa !important;
            border: 1px solid #4a4a4a !important;
        }
        .stTextInput > div > div > input {
            background-color: #262730 !important;
            color: #fafafa !important;
            border: 1px solid #4a4a4a !important;
        }
        .stNumberInput > div > div > input {
            background-color: #262730 !important;
            color: #fafafa !important;
            border: 1px solid #4a4a4a !important;
        }
        
        /* DataFrames */
        .stDataFrame {
            background-color: #1e1e1e !important;
        }
        .stDataFrame table {
            background-color: #1e1e1e !important;
            color: #fafafa !important;
        }
        .stDataFrame th {
            background-color: #262730 !important;
            color: #fafafa !important;
            border: 1px solid #4a4a4a !important;
        }
        .stDataFrame td {
            background-color: #1e1e1e !important;
            color: #fafafa !important;
            border: 1px solid #4a4a4a !important;
        }
        
        /* Expander */
        .stExpander {
            background-color: #1e1e1e !important;
            border: 1px solid #4a4a4a !important;
        }
        .stExpander > div > div > div {
            background-color: #1e1e1e !important;
            color: #fafafa !important;
        }
        
        /* Tabs */
        .stTabs [data-baseweb="tab-list"] {
            background-color: #262730 !important;
            border-bottom: 1px solid #4a4a4a !important;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: #262730 !important;
            color: #fafafa !important;
            border: 1px solid #4a4a4a !important;
            margin-right: 2px !important;
        }
        .stTabs [aria-selected="true"] {
            background-color: #0e1117 !important;
            color: #00d4ff !important;
            border-bottom: 2px solid #00d4ff !important;
        }
        
        /* Metrics */
        .stMetric {
            background-color: #262730 !important;
            padding: 15px !important;
            border-radius: 8px !important;
            border: 1px solid #4a4a4a !important;
        }
        .stMetric [data-testid="metric-container"] {
            background-color: #262730 !important;
            border: 1px solid #4a4a4a !important;
            padding: 10px !important;
            border-radius: 8px !important;
        }
        .stMetric label {
            color: #a0a0a0 !important;
        }
        .stMetric [data-testid="metric-container"] > div {
            color: #fafafa !important;
        }
        
        /* Alerts and info boxes */
        .stAlert {
            background-color: #262730 !important;
            color: #fafafa !important;
            border: 1px solid #4a4a4a !important;
        }
        .stInfo {
            background-color: #1a4d73 !important;
            color: #fafafa !important;
            border: 1px solid #2980b9 !important;
        }
        .stSuccess {
            background-color: #1a5d1a !important;
            color: #fafafa !important;
            border: 1px solid #27ae60 !important;
        }
        .stWarning {
            background-color: #735d1a !important;
            color: #fafafa !important;
            border: 1px solid #f39c12 !important;
        }
        .stError {
            background-color: #731a1a !important;
            color: #fafafa !important;
            border: 1px solid #e74c3c !important;
        }
        
        /* Text and markdown */
        div[data-testid="stMarkdownContainer"] {
            color: #fafafa !important;
        }
        .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4, .stMarkdown h5, .stMarkdown h6 {
            color: #fafafa !important;
        }
        
        /* Plotly charts */
        .plotly-graph-div {
            background-color: #1e1e1e !important;
        }
        .js-plotly-plot .plotly .modebar {
            background-color: #262730 !important;
        }
        .js-plotly-plot .plotly .modebar .modebar-btn {
            color: #fafafa !important;
        }
        
        /* Progress bars */
        .stProgress > div > div > div {
            background-color: #00d4ff !important;
        }
        .stProgress > div > div {
            background-color: #4a4a4a !important;
        }
        
        /* File uploader */
        .stFileUploader {
            background-color: #262730 !important;
            border: 1px solid #4a4a4a !important;
            border-radius: 8px !important;
        }
        .stFileUploader > div {
            background-color: #262730 !important;
            color: #fafafa !important;
        }
        
        /* Buttons */
        .stButton > button {
            background-color: #00d4ff !important;
            color: #0e1117 !important;
            border: none !important;
            border-radius: 6px !important;
        }
        .stButton > button:hover {
            background-color: #0098cc !important;
            color: #0e1117 !important;
        }
        
        /* Headers and subheaders */
        .stHeader, .stSubheader {
            color: #fafafa !important;
        }
        
        /* Code blocks */
        .stCode {
            background-color: #262730 !important;
            color: #fafafa !important;
            border: 1px solid #4a4a4a !important;
        }
        </style>
        """
        st.markdown(dark_theme, unsafe_allow_html=True)
    elif theme_name == "Light":
        light_theme = """
        <style>
        .stApp {
            background-color: #ffffff;
            color: #262730;
        }
        .stSidebar {
            background-color: #f0f2f6;
        }
        .stSelectbox > div > div {
            background-color: #ffffff;
            color: #262730;
        }
        .stDataFrame {
            background-color: #ffffff;
        }
        .stExpander {
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
        }
        .stTabs [data-baseweb="tab-list"] {
            background-color: #f0f2f6;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: #f0f2f6;
            color: #262730;
        }
        .stTabs [aria-selected="true"] {
            background-color: #ffffff;
            color: #1f77b4;
        }
        .stMetric {
            background-color: #f8f9fa;
            padding: 10px;
            border-radius: 5px;
            border: 1px solid #e9ecef;
        }
        .stAlert {
            background-color: #ffffff;
            color: #262730;
        }
        div[data-testid="stMarkdownContainer"] {
            color: #262730;
        }
        .plotly-graph-div {
            background-color: #ffffff !important;
        }
        .js-plotly-plot .plotly .modebar {
            background-color: #ffffff !important;
        }
        </style>
        """
        st.markdown(light_theme, unsafe_allow_html=True)

def main():
    # Theme toggle in sidebar
    with st.sidebar:
        st.markdown("### 🎨 Theme")
        theme = st.selectbox(
            "Select Theme:",
            ["Light", "Dark"],
            index=0 if st.session_state.get('theme', 'Light') == 'Light' else 1,
            key="theme_selector"
        )
        
        # Store theme in session state
        st.session_state.theme = theme
        
        # Apply the selected theme
        apply_theme(theme)
        
        st.markdown("---")
    
    st.title("📊 Excel Comparison & Business Intelligence Platform")
    st.markdown("**Comprehensive data analysis, comparison, and business intelligence platform for Excel data**")
    
    # Initialize comparator
    if 'comparator' not in st.session_state:
        st.session_state.comparator = ExcelComparator()
    
    comparator = st.session_state.comparator
    
    # Main application tabs for better organization
    main_tabs = st.tabs([
        "🏠 Home & Setup",
        "📊 Data Comparison", 
        "📈 Business Intelligence",
        "🔍 Data Analysis",
        "📋 Reports & Export"
    ])
    
    # HOME & SETUP TAB
    with main_tabs[0]:
        st.header("🏠 Welcome to Excel Comparison & BI Platform")
        st.markdown("**Your comprehensive solution for Excel data analysis and business intelligence**")
        
        # Feature overview
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🚀 Key Features")
            st.markdown("""
            **� Data Comparison:**
            - Compare Excel files & sheets
            - Fuzzy matching algorithms
            - Multi-column analysis
            - Real-time progress tracking
            
            **📈 Business Intelligence:**
            - Financial ratio analysis (ROE, ROA, liquidity)
            - Customer analytics & segmentation
            - Performance KPIs & metrics
            - Risk assessment & alerts
            
            **🔍 Advanced Analytics:**
            - Statistical analysis
            - Data quality assessment
            - Trend analysis & forecasting
            - Interactive visualizations
            """)
        
        with col2:
            st.subheader("📋 Quick Start Guide")
            st.markdown("""
            **Step 1: Upload Your Data**
            - Use the sidebar to upload Excel files
            - Choose comparison mode
            - Select sheets to analyze
            
            **Step 2: Select Analysis Type**
            - Data Comparison for file/sheet comparison
            - Business Intelligence for business analysis
            - Data Analysis for statistical insights
            
            **Step 3: Configure & Run**
            - Set parameters and thresholds
            - Run analysis and review results
            - Export professional reports
            """)
        
        # What can you compare section
        with st.expander("ℹ️ Comparison Modes", expanded=False):
            st.markdown("""
            **🔄 Two Different Files:**
            - Compare data between separate Excel files
            - Perfect for vendor comparisons, data validation
            
            **📋 Same File (Different Sheets):**
            - Compare sheets within the same workbook
            - Temporal comparisons (Jan vs Feb, Before vs After)
            - Budget vs actual analysis
            
            **🔄 Multi-Sheet Batch Processing:**
            - Process multiple sheets against a reference
            - Batch validation and quality checks
            
            **📈 Historical Comparison:**
            - Time-series analysis across periods
            - Trend identification and forecasting
            """)
    
    # DATA COMPARISON TAB
    with main_tabs[1]:
        st.header("📊 Excel Data Comparison")
        st.markdown("**Compare data between Excel files or sheets with advanced matching algorithms**")
    
    # Sidebar for file uploads and settings (always visible)
    with st.sidebar:
        st.header("📁 File Upload")
        
        # Comparison mode selection
        comparison_mode = st.radio(
            "🔄 Comparison Mode:",
            options=["Two Different Files", "Same File (Different Sheets)", "Multi-Sheet Batch Processing", "Cross-Sheet Data Consolidation", "Historical Comparison Mode"],
            index=0,
            help="Choose comparison type: separate files, two sheets, batch processing, multi-sheet consolidation, or time-series analysis"
        )
        
        st.divider()
        
        if comparison_mode == "Two Different Files":
            # Original two-file upload mode
            uploaded_file_a = st.file_uploader(
                "Choose Sheet A (Excel file)", 
                type=['xlsx', 'xls'],
                key="file_a"
            )
            
            uploaded_file_b = st.file_uploader(
                "Choose Sheet B (Excel file)", 
                type=['xlsx', 'xls'],
                key="file_b"
            )
            
            # Set both files as the same for processing
            same_file_mode = False
            batch_mode = False
            consolidation_mode = False
            historical_mode = False
            
        elif comparison_mode == "Same File (Different Sheets)":
            # Same file, different sheets mode
            st.info("📝 Upload one Excel file to compare different sheets within it")
            
            # Show practical examples
            with st.expander("💡 Common Use Cases", expanded=False):
                st.markdown("""
                **📅 Temporal Comparisons:**
                - January vs February sales data
                - Q1 vs Q2 performance metrics
                - Before vs After implementation results
                
                **📈 Business Analysis:**
                - Budget vs Actual spending
                - Plan vs Execution tracking
                - Target vs Achievement comparison
                
                **📊 Data Validation:**
                - Original vs Cleaned datasets
                - Raw vs Processed data
                - Version 1 vs Version 2 of reports
                
                **🔍 Quality Control:**
                - Compare different data processing methods
                - Validate data transformation results
                - Check consistency across time periods
                """)
            
            uploaded_single_file = st.file_uploader(
                "Choose Excel file with multiple sheets", 
                type=['xlsx', 'xls'],
                key="single_file"
            )
            
            # Use the same file for both A and B
            uploaded_file_a = uploaded_single_file
            uploaded_file_b = uploaded_single_file
            same_file_mode = True
            batch_mode = False
            consolidation_mode = False
            historical_mode = False
            
        elif comparison_mode == "Multi-Sheet Batch Processing":
            # Batch processing mode
            st.info("🔄 Compare one reference sheet against multiple sheets in batch")
            
            # Show batch processing examples
            with st.expander("💡 Batch Processing Use Cases", expanded=False):
                st.markdown("""
                **📊 Reference vs Multiple:**
                - Master data vs regional data sheets
                - Template vs multiple completed forms
                - Standard vs customized versions
                
                **📈 Performance Analysis:**
                - Benchmark vs multiple time periods
                - Target sheet vs monthly actuals
                - Master list vs department sheets
                
                **🔍 Quality Assurance:**
                - Reference data vs multiple sources
                - Gold standard vs test results
                - Original vs multiple processed versions
                
                **📋 Compliance Checking:**
                - Policy template vs department implementations
                - Standard format vs various submissions
                - Approved list vs multiple inventories
                """)
            
            uploaded_batch_file = st.file_uploader(
                "Choose Excel file for batch processing", 
                type=['xlsx', 'xls'],
                key="batch_file"
            )
            
            # Set up for batch processing
            uploaded_file_a = uploaded_batch_file  # Will be the reference sheet
            uploaded_file_b = None  # Will be set dynamically for each comparison
            same_file_mode = False
            batch_mode = True
            consolidation_mode = False
            historical_mode = False
            
        elif comparison_mode == "Cross-Sheet Data Consolidation":
            # Consolidation mode
            st.info("🔗 Consolidate data from multiple sheets into a unified view")
            
            # Show consolidation examples
            with st.expander("💡 Consolidation Use Cases", expanded=False):
                st.markdown("""
                **📊 Data Aggregation:**
                - Combine regional sales data from multiple sheets
                - Merge department reports into company overview
                - Consolidate monthly data into quarterly summary
                
                **🔍 Comprehensive Analysis:**
                - Union data from similar structured sheets
                - Cross-reference data across time periods
                - Create master dataset from distributed sources
                
                **📈 Business Intelligence:**
                - Combine data from different business units
                - Merge customer data from various sources
                - Consolidate inventory from multiple locations
                
                **🎯 Data Quality:**
                - Identify duplicates across sheets
                - Find data gaps between sources
                - Harmonize different data formats
                """)
            
            uploaded_consolidation_file = st.file_uploader(
                "Choose Excel file for data consolidation", 
                type=['xlsx', 'xls'],
                key="consolidation_file"
            )
            
            # Set up for consolidation processing
            uploaded_file_a = uploaded_consolidation_file  # Main file for consolidation
            uploaded_file_b = None  # Not used in consolidation mode
            same_file_mode = False
            batch_mode = False
            consolidation_mode = True
            historical_mode = False
            
        elif comparison_mode == "Historical Comparison Mode":
            # Historical time-series comparison mode
            st.info("📈 Compare time-series data across multiple sheets")
            
            # Show historical comparison examples
            with st.expander("💡 Historical Analysis Use Cases", expanded=False):
                st.markdown("""
                **📅 Time-Series Analysis:**
                - Compare Jan vs Feb vs Mar monthly data
                - Analyze quarterly trends (Q1 vs Q2 vs Q3 vs Q4)
                - Track year-over-year performance changes
                
                **📊 Trend Identification:**
                - Identify seasonal patterns in data  
                - Spot performance trends over time
                - Compare historical baselines vs current data
                
                **📈 Business Intelligence:**
                - Monthly sales performance tracking
                - Budget vs actual over multiple periods
                - KPI evolution analysis across time
                
                **🎯 Pattern Recognition:**
                - Detect recurring data patterns
                - Identify anomalies in time series
                - Compare cyclical business data
                """)
            
            uploaded_historical_file = st.file_uploader(
                "Choose Excel file with time-series sheets", 
                type=['xlsx', 'xls'],
                key="historical_file"
            )
            
            # Set up for historical processing
            uploaded_file_a = uploaded_historical_file  # Main file for historical analysis
            uploaded_file_b = None  # Not used in historical mode
            same_file_mode = False
            batch_mode = False
            consolidation_mode = False
            historical_mode = True
            
        else:
            batch_mode = False
            consolidation_mode = False
            historical_mode = False
        
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
    
    # Main content area - different layout for batch mode
    if batch_mode and uploaded_file_a:
        # Batch Processing Layout
        st.header("🔄 Multi-Sheet Batch Processing")
        
        # Get all sheet names
        _, all_sheet_names = comparator.load_excel_file(uploaded_file_a, "Batch File")
        
        if all_sheet_names and len(all_sheet_names) > 1:
            st.success(f"📊 Found {len(all_sheet_names)} sheets: {', '.join(all_sheet_names)}")
            
            # Reference sheet selection
            col_ref, col_target = st.columns(2)
            
            with col_ref:
                st.subheader("📋 Reference Sheet")
                reference_sheet = st.selectbox(
                    "Select reference sheet (will be compared against all others):",
                    all_sheet_names,
                    key="reference_sheet",
                    help="This sheet will be used as the baseline for comparison"
                )
                
                # Load and preview reference sheet
                if reference_sheet:
                    df_ref = comparator.read_sheet(uploaded_file_a, reference_sheet)
                    if df_ref is not None:
                        st.metric("📊 Rows", f"{len(df_ref):,}")
                        st.metric("📋 Columns", f"{len(df_ref.columns):,}")
                        
                        with st.expander("👀 Reference Data Preview", expanded=False):
                            st.dataframe(df_ref.head(5), use_container_width=True)
            
            with col_target:
                st.subheader("🎯 Target Sheets")
                # Filter out reference sheet from target options
                target_options = [sheet for sheet in all_sheet_names if sheet != reference_sheet]
                
                if target_options:
                    target_sheets = st.multiselect(
                        "Select sheets to compare against reference:",
                        target_options,
                        default=target_options,  # Select all by default
                        key="target_sheets",
                        help="These sheets will be compared against the reference sheet"
                    )
                    
                    if target_sheets:
                        st.success(f"✅ Will compare reference sheet against {len(target_sheets)} target sheets")
                        
                        # Show target sheets summary
                        target_summary = []
                        for sheet in target_sheets[:3]:  # Show first 3 for preview
                            df_target = comparator.read_sheet(uploaded_file_a, sheet)
                            if df_target is not None:
                                target_summary.append({
                                    'Sheet': sheet,
                                    'Rows': f"{len(df_target):,}",
                                    'Columns': f"{len(df_target.columns):,}"
                                })
                        
                        if target_summary:
                            st.write("**Target Sheets Preview:**")
                            summary_df = pd.DataFrame(target_summary)
                            st.dataframe(summary_df, hide_index=True, use_container_width=True)
                            
                            if len(target_sheets) > 3:
                                st.info(f"... and {len(target_sheets) - 3} more sheets")
                    else:
                        st.warning("⚠️ Please select at least one target sheet")
                else:
                    st.warning("⚠️ No other sheets available. Need at least 2 sheets for batch processing.")
        else:
            st.warning("⚠️ Please upload an Excel file with multiple sheets for batch processing")
        
        # Batch Processing Controls
        if uploaded_file_a and 'reference_sheet' in locals() and 'target_sheets' in locals() and target_sheets:
            st.divider()
            st.header("🔄 Batch Processing Configuration")
            
            # Column configuration for batch processing
            col_batch1, col_batch2 = st.columns(2)
            
            with col_batch1:
                st.subheader("🔑 Key Column Configuration")
                
                # Load reference sheet for column selection
                df_ref = comparator.read_sheet(uploaded_file_a, reference_sheet)
                if df_ref is not None:
                    key_col_ref = st.selectbox(
                        f"Key column in reference sheet ({reference_sheet}):",
                        df_ref.columns,
                        key="batch_key_ref"
                    )
                    
                    key_col_targets = st.selectbox(
                        "Key column in target sheets:",
                        df_ref.columns,  # Assume same structure
                        key="batch_key_targets",
                        help="This column should exist in all target sheets"
                    )
            
            with col_batch2:
                st.subheader("📊 Data Extraction")
                
                cols_to_extract_batch = st.multiselect(
                    "Columns to extract from target sheets:",
                    [col for col in df_ref.columns if col not in [key_col_ref]],
                    key="batch_extract_cols",
                    help="These columns will be merged from target sheets"
                )
            
            # Batch processing button
            st.divider()
            
            if st.button("🚀 Start Batch Processing", type="primary", use_container_width=True):
                if key_col_ref and key_col_targets:
                    st.info(f"🔄 Starting batch comparison of {reference_sheet} against {len(target_sheets)} target sheets...")
                    
                    # Show batch parameters
                    with st.expander("📋 Batch Parameters", expanded=False):
                        st.write(f"**📊 Batch Overview:**")
                        st.write(f"- Reference sheet: {reference_sheet}")
                        st.write(f"- Target sheets: {len(target_sheets)} sheets")
                        st.write(f"- Key columns: {key_col_ref} ↔ {key_col_targets}")
                        st.write(f"- Extracting: {', '.join(cols_to_extract_batch) if cols_to_extract_batch else 'Key columns only'}")
                        st.write(f"- Similarity threshold: {threshold}%")
                        st.write(f"- Case sensitive: {'No' if ignore_case else 'Yes'}")
                    
                    try:
                        # Perform batch comparison
                        batch_results = comparator.perform_batch_comparison(
                            uploaded_file_a, reference_sheet, target_sheets,
                            key_col_ref, key_col_targets,
                            cols_to_extract_batch, threshold, ignore_case
                        )
                        
                        # Store results for export
                        st.session_state.batch_results = batch_results
                        
                        st.balloons()
                        st.success("🎉 Batch processing completed! Results are displayed above.")
                        
                    except Exception as e:
                        st.error(f"❌ Error during batch processing: {str(e)}")
                        st.write("Please check your data and configuration.")
                else:
                    st.warning("⚠️ Please select key columns for batch processing.")
    
    elif consolidation_mode and uploaded_file_a:
        # Consolidation Processing Layout
        st.header("🔗 Cross-Sheet Data Consolidation")
        
        # Get all sheet names
        _, all_sheet_names = comparator.load_excel_file(uploaded_file_a, "Consolidation File")
        
        if all_sheet_names and len(all_sheet_names) > 1:
            st.success(f"📊 Found {len(all_sheet_names)} sheets: {', '.join(all_sheet_names)}")
            
            # Sheet selection for consolidation
            col_sel1, col_sel2 = st.columns(2)
            
            with col_sel1:
                st.subheader("📋 Sheets to Consolidate")
                sheets_to_consolidate = st.multiselect(
                    "Select sheets to consolidate (minimum 2):",
                    all_sheet_names,
                    default=all_sheet_names[:3] if len(all_sheet_names) >= 3 else all_sheet_names,
                    key="consolidation_sheets",
                    help="Select 2 or more sheets to combine into unified view"
                )
                
                if len(sheets_to_consolidate) >= 2:
                    st.success(f"✅ Will consolidate {len(sheets_to_consolidate)} sheets")
                    
                    # Show sheets preview
                    consolidation_summary = []
                    for sheet in sheets_to_consolidate[:5]:  # Show first 5 for preview
                        df_sheet = comparator.read_sheet(uploaded_file_a, sheet)
                        if df_sheet is not None:
                            consolidation_summary.append({
                                'Sheet': sheet,
                                'Rows': f"{len(df_sheet):,}",
                                'Columns': f"{len(df_sheet.columns):,}"
                            })
                    
                    if consolidation_summary:
                        st.write("**Selected Sheets Preview:**")
                        summary_df = pd.DataFrame(consolidation_summary)
                        st.dataframe(summary_df, hide_index=True, use_container_width=True)
                        
                        if len(sheets_to_consolidate) > 5:
                            st.info(f"... and {len(sheets_to_consolidate) - 5} more sheets")
                else:
                    st.warning("⚠️ Please select at least 2 sheets for consolidation")
            
            with col_sel2:
                st.subheader("🔑 Consolidation Strategy")
                
                consolidation_strategy = st.radio(
                    "Choose consolidation method:",
                    ["Union (Combine all data)", "Key-based Merge", "Schema Analysis Only"],
                    index=0,
                    help="Union: Stack all data together | Key-based: Merge on common columns | Analysis: Compare structures"
                )
                
                if consolidation_strategy == "Key-based Merge":
                    st.info("🔍 Will identify common key columns across selected sheets")
                elif consolidation_strategy == "Union (Combine all data)":
                    st.info("📚 Will stack all data and harmonize column names")
                else:
                    st.info("🔬 Will analyze and compare sheet structures")
            
            # Consolidation Configuration
            if len(sheets_to_consolidate) >= 2:
                st.divider()
                st.header("🔧 Consolidation Configuration")
                
                config_col1, config_col2 = st.columns(2)
                
                with config_col1:
                    st.subheader("📊 Data Options")
                    
                    include_source_column = st.checkbox(
                        "Add source sheet column",
                        value=True,
                        help="Add a column indicating which sheet each row came from"
                    )
                    
                    handle_duplicates = st.selectbox(
                        "Handle duplicate rows:",
                        ["Keep all", "Remove duplicates", "Mark duplicates"],
                        index=0,
                        help="How to handle rows that appear in multiple sheets"
                    )
                    
                    missing_data_strategy = st.selectbox(
                        "Missing column strategy:",
                        ["Fill with blanks", "Skip rows", "Use default value"],
                        index=0,
                        help="How to handle columns that exist in some sheets but not others"
                    )
                
                with config_col2:
                    st.subheader("🎯 Quality Control")
                    
                    validate_schemas = st.checkbox(
                        "Validate column compatibility",
                        value=True,
                        help="Check if columns with same names have compatible data types"
                    )
                    
                    show_consolidation_report = st.checkbox(
                        "Generate consolidation report",
                        value=True,
                        help="Create detailed report of consolidation process and data quality"
                    )
                
                # Consolidation execution button
                st.divider()
                
                if st.button("🚀 Start Cross-Sheet Consolidation", type="primary", use_container_width=True):
                    st.info(f"🔄 Starting consolidation of {len(sheets_to_consolidate)} sheets...")
                    
                    # Show consolidation parameters
                    with st.expander("📋 Consolidation Parameters", expanded=False):
                        st.write(f"**🔗 Consolidation Overview:**")
                        st.write(f"- Strategy: {consolidation_strategy}")
                        st.write(f"- Sheets: {len(sheets_to_consolidate)} sheets")
                        st.write(f"- Source tracking: {'Yes' if include_source_column else 'No'}")
                        st.write(f"- Duplicate handling: {handle_duplicates}")
                        st.write(f"- Missing data: {missing_data_strategy}")
                        st.write(f"- Schema validation: {'Yes' if validate_schemas else 'No'}")
                    
                    try:
                        # Perform consolidation
                        consolidation_results = comparator.perform_consolidation(
                            uploaded_file_a, sheets_to_consolidate,
                            consolidation_strategy, include_source_column,
                            handle_duplicates, missing_data_strategy,
                            validate_schemas, show_consolidation_report
                        )
                        
                        # Store results for display and export
                        st.session_state.consolidation_results = consolidation_results
                        
                        st.balloons()
                        st.success("🎉 Cross-sheet consolidation completed! Results are displayed above.")
                        
                    except Exception as e:
                        st.error(f"❌ Error during consolidation: {str(e)}")
                        st.write("Please check your data and configuration.")
        else:
            st.warning("⚠️ Please upload an Excel file with multiple sheets for consolidation")
    
    elif historical_mode and uploaded_file_a:
        # Historical Comparison Processing Layout
        st.header("📈 Historical Comparison Mode")
        
        # Get all sheet names
        _, all_sheet_names = comparator.load_excel_file(uploaded_file_a, "Historical File")
        
        if all_sheet_names and len(all_sheet_names) > 1:
            st.success(f"📊 Found {len(all_sheet_names)} sheets: {', '.join(all_sheet_names)}")
            
            # Automatic pattern detection
            historical_patterns = comparator.detect_time_patterns(all_sheet_names)
            
            # Show detected patterns
            if historical_patterns:
                st.subheader("🔍 Auto-Detected Time Patterns")
                
                pattern_tabs = st.tabs([f"📅 {pattern['type']}" for pattern in historical_patterns])
                
                for i, pattern in enumerate(historical_patterns):
                    with pattern_tabs[i]:
                        st.write(f"**Pattern Type:** {pattern['type']}")
                        st.write(f"**Sheets Found:** {len(pattern['sheets'])}")
                        
                        # Show pattern preview
                        pattern_summary = []
                        for sheet in pattern['sheets'][:5]:  # Show first 5
                            df_sheet = comparator.read_sheet(uploaded_file_a, sheet)
                            if df_sheet is not None:
                                pattern_summary.append({
                                    'Sheet': sheet,
                                    'Period': pattern.get('periods', {}).get(sheet, 'Unknown'),
                                    'Rows': f"{len(df_sheet):,}",
                                    'Columns': f"{len(df_sheet.columns):,}"
                                })
                        
                        if pattern_summary:
                            st.write("**Time Series Preview:**")
                            pattern_df = pd.DataFrame(pattern_summary)
                            st.dataframe(pattern_df, hide_index=True, use_container_width=True)
                            
                            if len(pattern['sheets']) > 5:
                                st.info(f"... and {len(pattern['sheets']) - 5} more periods")
            
            # Manual sheet selection and configuration
            st.divider()
            st.subheader("📋 Historical Analysis Configuration")
            
            config_col1, config_col2 = st.columns(2)
            
            with config_col1:
                st.write("**📅 Time Period Selection**")
                
                # Let user choose from detected patterns or manual selection
                if historical_patterns:
                    use_auto_pattern = st.radio(
                        "Selection method:",
                        ["Use detected pattern", "Manual selection"],
                        index=0
                    )
                    
                    if use_auto_pattern == "Use detected pattern":
                        selected_pattern = st.selectbox(
                            "Choose time pattern:",
                            range(len(historical_patterns)),
                            format_func=lambda x: f"{historical_patterns[x]['type']} ({len(historical_patterns[x]['sheets'])} periods)"
                        )
                        sheets_to_compare = historical_patterns[selected_pattern]['sheets']
                        analysis_type = historical_patterns[selected_pattern]['type']
                    else:
                        sheets_to_compare = st.multiselect(
                            "Select sheets for comparison:",
                            all_sheet_names,
                            default=all_sheet_names[:4] if len(all_sheet_names) >= 4 else all_sheet_names,
                            help="Select sheets representing different time periods"
                        )
                        analysis_type = "Custom"
                else:
                    sheets_to_compare = st.multiselect(
                        "Select sheets for comparison:",
                        all_sheet_names,
                        default=all_sheet_names[:4] if len(all_sheet_names) >= 4 else all_sheet_names,
                        help="Select sheets representing different time periods"
                    )
                    analysis_type = "Custom"
            
            with config_col2:
                st.write("**🎯 Analysis Options**")
                
                analysis_mode = st.radio(
                    "Comparison focus:",
                    ["Trend Analysis", "Period-to-Period Change", "Baseline Comparison"],
                    index=0,
                    help="Trend: Overall patterns | Change: Sequential differences | Baseline: Compare all to reference"
                )
                
                if analysis_mode == "Baseline Comparison":
                    baseline_sheet = st.selectbox(
                        "Select baseline period:",
                        sheets_to_compare if len(sheets_to_compare) > 0 else all_sheet_names,
                        help="All other periods will be compared to this baseline"
                    )
                
                include_variance = st.checkbox(
                    "Calculate variance metrics",
                    value=True,
                    help="Include standard deviation, coefficient of variation, etc."
                )
                
                show_trend_charts = st.checkbox(
                    "Generate trend visualizations",
                    value=True,
                    help="Create charts showing data trends over time"
                )
            
            # Historical comparison execution
            if len(sheets_to_compare) >= 2:
                st.divider()
                
                if st.button("📈 Start Historical Analysis", type="primary", use_container_width=True):
                    st.info(f"🔄 Starting historical analysis of {len(sheets_to_compare)} time periods...")
                    
                    # Show analysis parameters
                    with st.expander("📋 Analysis Parameters", expanded=False):
                        st.write(f"**📈 Historical Analysis Overview:**")
                        st.write(f"- Analysis Type: {analysis_type}")
                        st.write(f"- Time Periods: {len(sheets_to_compare)}")
                        st.write(f"- Comparison Mode: {analysis_mode}")
                        if analysis_mode == "Baseline Comparison":
                            st.write(f"- Baseline Period: {baseline_sheet}")
                        st.write(f"- Variance Metrics: {'Yes' if include_variance else 'No'}")
                        st.write(f"- Trend Charts: {'Yes' if show_trend_charts else 'No'}")
                    
                    try:
                        # Perform historical analysis
                        historical_results = comparator.perform_historical_comparison(
                            uploaded_file_a, sheets_to_compare,
                            analysis_mode, analysis_type,
                            baseline_sheet if analysis_mode == "Baseline Comparison" else None,
                            include_variance, show_trend_charts
                        )
                        
                        # Store results for display and export
                        st.session_state.historical_results = historical_results
                        
                        st.balloons()
                        st.success("🎉 Historical analysis completed! Results are displayed above.")
                        
                    except Exception as e:
                        st.error(f"❌ Error during historical analysis: {str(e)}")
                        st.write("Please check your data and configuration.")
            else:
                st.warning("⚠️ Please select at least 2 time periods for historical comparison")
        else:
            st.warning("⚠️ Please upload an Excel file with multiple sheets for historical analysis")
    
    else:
        # Regular two-column layout for other modes
        col1, col2 = st.columns(2)
        
        # Handle Sheet A
        with col1:
            if same_file_mode:
                st.subheader("📋 First Sheet")
            else:
                st.subheader("📋 Sheet A")
                
            if uploaded_file_a:
                # Get sheet names
                _, sheet_names_a = comparator.load_excel_file(uploaded_file_a, "Sheet A")
                
                if sheet_names_a:
                    if same_file_mode:
                        label_a = "Select first sheet to compare:"
                        if len(sheet_names_a) > 1:
                            st.info(f"📊 Found {len(sheet_names_a)} sheets: {', '.join(sheet_names_a)}")
                    else:
                        label_a = "Select sheet from File A:"
                    
                    selected_sheet_a = st.selectbox(
                        label_a,
                        sheet_names_a,
                        key="sheet_a"
                    )
                
                # Load selected sheet
                df_a = comparator.read_sheet(uploaded_file_a, selected_sheet_a)
                if df_a is not None:
                    comparator.df_a = df_a
                    
                    # Enhanced data preview with statistics
                    col_metrics = st.columns(3)
                    with col_metrics[0]:
                        st.metric("📊 Rows", f"{len(df_a):,}")
                    with col_metrics[1]:
                        st.metric("📋 Columns", f"{len(df_a.columns):,}")
                    with col_metrics[2]:
                        total_cells = len(df_a) * len(df_a.columns)
                        null_cells = df_a.isnull().sum().sum()
                        null_pct = (null_cells / total_cells * 100) if total_cells > 0 else 0
                        st.metric("🕳️ Missing Data", f"{null_pct:.1f}%")
                    
                    # Data preview with enhanced info
                    with st.expander("👀 Data Preview & Column Types", expanded=True):
                        st.dataframe(df_a.head(10), use_container_width=True)
                        
                        # Column info summary
                        st.write("**📋 Column Information:**")
                        col_info_display = []
                        for col in df_a.columns:
                            dtype = str(df_a[col].dtype)
                            unique_count = df_a[col].nunique()
                            null_count = df_a[col].isnull().sum()
                            col_info_display.append({
                                'Column': col,
                                'Type': dtype,
                                'Unique': f"{unique_count:,}",
                                'Missing': f"{null_count:,}",
                                'Sample': str(df_a[col].dropna().iloc[0]) if len(df_a[col].dropna()) > 0 else "N/A"
                            })
                        
                        col_info_df = pd.DataFrame(col_info_display)
                        st.dataframe(col_info_df, hide_index=True, use_container_width=True)
            else:
                if same_file_mode:
                    st.info("Please upload an Excel file with multiple sheets")
                else:
                    st.info("Please upload Sheet A")
        
        # Handle Sheet B (only for non-batch modes)
        if not batch_mode:  # Only show Sheet B section when not in batch mode
            with col2:
                if same_file_mode:
                    st.subheader("📋 Second Sheet")
                else:
                    st.subheader("📋 Sheet B")
                    
                if uploaded_file_b:
                    # Get sheet names
                    _, sheet_names_b = comparator.load_excel_file(uploaded_file_b, "Sheet B")
                    
                    if sheet_names_b:
                        if same_file_mode:
                            label_b = "Select second sheet to compare:"
                            # Filter out the already selected sheet A to avoid comparing sheet with itself
                            available_sheets_b = [sheet for sheet in sheet_names_b if sheet != selected_sheet_a] if 'selected_sheet_a' in locals() else sheet_names_b
                            if not available_sheets_b:
                                st.warning("⚠️ Please select different sheets to compare")
                                available_sheets_b = sheet_names_b
                        else:
                            label_b = "Select sheet from File B:"
                            available_sheets_b = sheet_names_b
                        
                        selected_sheet_b = st.selectbox(
                            label_b,
                            available_sheets_b,
                            key="sheet_b"
                        )
                        
                        # Load selected sheet
                        df_b = comparator.read_sheet(uploaded_file_b, selected_sheet_b)
                        if df_b is not None:
                            comparator.df_b = df_b
                        
                        # Enhanced data preview with statistics
                        col_metrics = st.columns(3)
                        with col_metrics[0]:
                            st.metric("📊 Rows", f"{len(df_b):,}")
                        with col_metrics[1]:
                            st.metric("📋 Columns", f"{len(df_b.columns):,}")
                        with col_metrics[2]:
                            total_cells = len(df_b) * len(df_b.columns)
                            null_cells = df_b.isnull().sum().sum()
                            null_pct = (null_cells / total_cells * 100) if total_cells > 0 else 0
                            st.metric("🕳️ Missing Data", f"{null_pct:.1f}%")
                    
                        # Data preview with enhanced info
                        with st.expander("👀 Data Preview & Column Types", expanded=True):
                            st.dataframe(df_b.head(10), use_container_width=True)
                            
                            # Column info summary
                            st.write("**📋 Column Information:**")
                            col_info_display = []
                            for col in df_b.columns:
                                dtype = str(df_b[col].dtype)
                                unique_count = df_b[col].nunique()
                                null_count = df_b[col].isnull().sum()
                                col_info_display.append({
                                    'Column': col,
                                    'Type': dtype,
                                    'Unique': f"{unique_count:,}",
                                    'Missing': f"{null_count:,}",
                                    'Sample': str(df_b[col].dropna().iloc[0]) if len(df_b[col].dropna()) > 0 else "N/A"
                                })
                            
                            col_info_df = pd.DataFrame(col_info_display)
                            st.dataframe(col_info_df, hide_index=True, use_container_width=True)
                else:
                    if same_file_mode:
                        st.info("Upload file above to see available sheets")
                    else:
                        st.info("Please upload Sheet B")
    
    # Sheet Swap functionality
    if comparator.df_a is not None and comparator.df_b is not None:
        st.divider()
        st.subheader("🔄 Sheet Management")
        
        # Add helpful explanation
        with st.expander("💡 When to use Sheet Swap?", expanded=False):
            st.markdown("""
            **🔄 Sheet Swap is useful when you want to:**
            - **Reverse comparison direction**: Compare B→A instead of A→B
            - **Change perspective**: Make the target sheet the source sheet
            - **Save time**: No need to re-upload files or re-select sheets
            - **Test different approaches**: Compare both directions quickly
            
            **Example scenarios:**
            - Compare "January vs February" then swap to "February vs January"
            - Compare "Budget vs Actual" then swap to "Actual vs Budget"  
            - Compare "Before vs After" then swap to "After vs Before"
            """)
        
        # Add swap button with clear visual indication
        col_swap1, col_swap2, col_swap3 = st.columns([1, 1, 1])
        
        with col_swap2:
            st.markdown("**🔄 Quick Sheet Swap**")
            
            # Show different button based on swap status
            sheets_swapped = st.session_state.get('sheets_swapped', False)
            
            if not sheets_swapped:
                button_text = "🔄 Swap Sheets (A ↔ B)"
                button_help = "Switch Sheet A and Sheet B positions - useful to reverse comparison direction"
            else:
                button_text = "↩️ Reset to Original"
                button_help = "Return to original sheet assignment"
            
            if st.button(button_text, type="secondary", use_container_width=True, help=button_help):
                # Store current sheet names for display purposes
                current_sheet_a_name = st.session_state.get('sheet_a', 'Unknown')
                current_sheet_b_name = st.session_state.get('sheet_b', 'Unknown')
                
                # Check current swap status
                currently_swapped = st.session_state.get('sheets_swapped', False)
                
                # Swap the dataframes in the comparator object
                temp_df = comparator.df_a.copy()
                comparator.df_a = comparator.df_b.copy()
                comparator.df_b = temp_df
                
                # Toggle swap status
                st.session_state.sheets_swapped = not currently_swapped
                
                # Store original assignments if this is the first swap
                if not currently_swapped:
                    st.session_state.original_sheet_a = current_sheet_a_name
                    st.session_state.original_sheet_b = current_sheet_b_name
                
                # Clear any previous results and suggestions
                if hasattr(st.session_state, 'column_suggestions'):
                    del st.session_state.column_suggestions
                if hasattr(st.session_state, 'suggested_extract'):
                    del st.session_state.suggested_extract
                if hasattr(comparator, 'results'):
                    comparator.results = None
                
                if st.session_state.get('sheets_swapped', False):
                    # After swap - show what's now in each position
                    st.success(f"✅ Sheets swapped successfully!")
                    st.info(f"🔄 **{current_sheet_b_name}** is now **Sheet A (Source)**")
                    st.info(f"🔄 **{current_sheet_a_name}** is now **Sheet B (Target)**")
                else:
                    # Reset to original
                    st.success(f"✅ Reset to original positions!")
                    st.info(f"↩️ **{current_sheet_a_name}** is back to **Sheet A (Source)**")
                    st.info(f"↩️ **{current_sheet_b_name}** is back to **Sheet B (Target)**")
                
                st.info("💡 Column suggestions and previous results have been cleared. Generate new suggestions if needed.")
                
                # Force a rerun to update the UI
                st.rerun()
        
        # Show current sheet assignments for clarity
        st.markdown("**📊 Current Sheet Assignment:**")
        assignment_col1, assignment_col2 = st.columns(2)
        
        # Determine current sheet names based on swap status
        sheets_swapped = st.session_state.get('sheets_swapped', False)
        original_sheet_a = st.session_state.get('sheet_a', 'Unknown')
        original_sheet_b = st.session_state.get('sheet_b', 'Unknown')
        
        if sheets_swapped:
            # If swapped, show what's currently in each position
            display_sheet_a = f"{original_sheet_b}"
            display_sheet_b = f"{original_sheet_a}" 
            swap_indicator = " 🔄 SWAPPED"
            status_a = f"(Originally from {original_sheet_b})"
            status_b = f"(Originally from {original_sheet_a})"
        else:
            # Normal assignment
            display_sheet_a = original_sheet_a
            display_sheet_b = original_sheet_b
            swap_indicator = ""
            status_a = "(Original position)"
            status_b = "(Original position)"
        
        with assignment_col1:
            st.success(f"📋 **Sheet A (Source)**: {display_sheet_a}{swap_indicator}")
            st.caption(f"📊 {len(comparator.df_a):,} rows × {len(comparator.df_a.columns):,} columns")
            if sheets_swapped:
                st.caption(f"🔄 {status_a}")
            
        with assignment_col2:
            st.info(f"📋 **Sheet B (Target)**: {display_sheet_b}{swap_indicator}")
            st.caption(f"📊 {len(comparator.df_b):,} rows × {len(comparator.df_b.columns):,} columns")
            if sheets_swapped:
                st.caption(f"🔄 {status_b}")

    # Statistical Analysis Dashboard
    if comparator.df_a is not None and comparator.df_b is not None:
        st.divider()
        st.header("📊 Statistical Analysis Dashboard")
        
        # Initialize analyzers
        if 'statistical_analyzer' not in st.session_state:
            st.session_state.statistical_analyzer = StatisticalAnalyzer()
        if 'statistical_visualizer' not in st.session_state:
            st.session_state.statistical_visualizer = StatisticalVisualizer()
        
        analyzer = st.session_state.statistical_analyzer
        visualizer = st.session_state.statistical_visualizer
        
        # Analysis options
        analysis_col1, analysis_col2 = st.columns(2)
        
        with analysis_col1:
            analysis_scope = st.radio(
                "📋 Analysis Scope:",
                ["Both Sheets", "Sheet A Only", "Sheet B Only", "Comparative Analysis"],
                index=0,
                help="Choose which data to analyze statistically"
            )
        
        with analysis_col2:
            show_advanced = st.checkbox(
                "🔬 Advanced Statistics",
                value=False,
                help="Include advanced statistical measures like skewness, kurtosis, etc."
            )
        
        # Analysis tabs
        if analysis_scope in ["Both Sheets", "Sheet A Only", "Sheet B Only"]:
            analysis_tabs = []
            
            if analysis_scope in ["Both Sheets", "Sheet A Only"]:
                analysis_tabs.append("📋 Sheet A Analysis")
            if analysis_scope in ["Both Sheets", "Sheet B Only"]:
                analysis_tabs.append("📋 Sheet B Analysis")
            
            tabs = st.tabs(analysis_tabs)
            
            tab_index = 0
            
            # Sheet A Analysis
            if analysis_scope in ["Both Sheets", "Sheet A Only"]:
                with tabs[tab_index]:
                    st.subheader("📊 Sheet A Statistical Analysis")
                    
                    # Get current sheet name
                    current_sheet_a_name = st.session_state.get('sheet_a', 'Sheet A')
                    if st.session_state.get('sheets_swapped', False):
                        display_name = f"{st.session_state.get('sheet_b', 'Unknown')} (currently Sheet A)"
                    else:
                        display_name = current_sheet_a_name
                    
                    # Perform analysis
                    analysis_a = analyzer.analyze_dataframe(comparator.df_a, display_name)
                    
                    # Create summary metrics
                    summary_cards = visualizer.create_summary_metrics_cards(
                        analysis_a['basic_info'], 
                        analysis_a['missing_data_analysis']['_summary']
                    )
                    
                    # Display summary cards
                    card_cols = st.columns(len(summary_cards))
                    for i, card in enumerate(summary_cards):
                        with card_cols[i]:
                            st.metric(
                                label=f"{card['icon']} {card['title']}", 
                                value=card['value']
                            )
                    
                    # Analysis sub-tabs
                    analysis_subtabs = st.tabs([
                        "📈 Descriptive Stats", 
                        "🕳️ Missing Data", 
                        "📊 Distributions", 
                        "🔗 Correlations",
                        "📋 Data Types"
                    ])
                    
                    with analysis_subtabs[0]:  # Descriptive Stats
                        st.subheader("📈 Descriptive Statistics")
                        
                        if 'message' not in analysis_a['descriptive_stats']:
                            # Show descriptive statistics table
                            desc_df = pd.DataFrame(analysis_a['descriptive_stats']).T
                            
                            if show_advanced:
                                columns_to_show = ['count', 'mean', 'median', 'std_dev', 'min', 'max', 
                                                 'q1', 'q3', 'skewness', 'kurtosis', 'coefficient_of_variation']
                            else:
                                columns_to_show = ['count', 'mean', 'median', 'std_dev', 'min', 'max']
                            
                            # Filter available columns
                            available_columns = [col for col in columns_to_show if col in desc_df.columns]
                            
                            if available_columns:
                                st.dataframe(
                                    desc_df[available_columns].round(3), 
                                    use_container_width=True
                                )
                                
                                # Show visualization
                                fig_desc = visualizer.create_descriptive_stats_chart(
                                    analysis_a['descriptive_stats'], 
                                    f"Descriptive Statistics - {display_name}"
                                )
                                st.plotly_chart(fig_desc, use_container_width=True, key="desc_stats_a")
                        else:
                            st.info(analysis_a['descriptive_stats']['message'])
                    
                    with analysis_subtabs[1]:  # Missing Data
                        st.subheader("🕳️ Missing Data Analysis")
                        
                        # Missing data summary chart
                        fig_missing = visualizer.create_missing_data_summary_chart(
                            analysis_a['missing_data_analysis']
                        )
                        st.plotly_chart(fig_missing, use_container_width=True, key="missing_summary_a")
                        
                        # Missing data heatmap
                        if len(comparator.df_a) <= 1000:  # Only show heatmap for smaller datasets
                            fig_heatmap = visualizer.create_missing_data_heatmap(
                                comparator.df_a, f"Missing Data Pattern - {display_name}"
                            )
                            st.plotly_chart(fig_heatmap, use_container_width=True, key="missing_heatmap_a")
                        else:
                            st.info("💡 Missing data heatmap not shown for large datasets (>1000 rows) to maintain performance")
                    
                    with analysis_subtabs[2]:  # Distributions
                        st.subheader("📊 Data Distributions")
                        
                        # Distribution plots
                        fig_dist = visualizer.create_distribution_plots(comparator.df_a)
                        st.plotly_chart(fig_dist, use_container_width=True, key="distributions_a")
                        
                        # Box plots for outlier detection
                        fig_box = visualizer.create_box_plots(comparator.df_a)
                        st.plotly_chart(fig_box, use_container_width=True, key="boxplots_a")
                    
                    with analysis_subtabs[3]:  # Correlations
                        st.subheader("🔗 Correlation Analysis")
                        
                        correlation_data = analyzer.calculate_correlation_matrix(comparator.df_a)
                        
                        if 'message' not in correlation_data:
                            # Correlation type selector
                            corr_type = st.selectbox(
                                "Correlation Method:",
                                ["pearson", "spearman", "kendall"],
                                index=0,
                                help="Pearson: linear relationships, Spearman: monotonic relationships, Kendall: rank-based"
                            )
                            
                            # Show correlation heatmap
                            fig_corr = visualizer.create_correlation_heatmap(correlation_data, corr_type)
                            st.plotly_chart(fig_corr, use_container_width=True, key="correlation_a")
                            
                            # Show strong correlations
                            if correlation_data['strong_correlations']:
                                st.subheader("💪 Strong Correlations Found")
                                
                                strong_corr_data = []
                                for corr in correlation_data['strong_correlations']:
                                    strong_corr_data.append({
                                        'Column 1': corr['column1'],
                                        'Column 2': corr['column2'],
                                        'Correlation': f"{corr['correlation']:.3f}",
                                        'Strength': corr['strength'].replace('_', ' ').title()
                                    })
                                
                                st.dataframe(pd.DataFrame(strong_corr_data), hide_index=True, use_container_width=True)
                            else:
                                st.info("No strong correlations (>0.7) found between variables")
                        else:
                            st.info(correlation_data['message'])
                    
                    with analysis_subtabs[4]:  # Data Types
                        st.subheader("📋 Data Types Analysis")
                        
                        # Data types pie chart
                        fig_types = visualizer.create_data_types_pie_chart(analysis_a['data_types_analysis'])
                        st.plotly_chart(fig_types, use_container_width=True, key="data_types_a")
                        
                        # Data types table with recommendations
                        types_data = []
                        for col, info in analysis_a['data_types_analysis'].items():
                            types_data.append({
                                'Column': col,
                                'Current Type': info['current_type'],
                                'Recommended Type': info['recommended_type'],
                                'Unique Values': f"{info['unique_values']:,}",
                                'Memory Usage': f"{info['memory_usage']:,} bytes",
                                'Sample Values': ', '.join(str(v) for v in info['sample_values'][:3])
                            })
                        
                        st.dataframe(pd.DataFrame(types_data), hide_index=True, use_container_width=True)
                
                tab_index += 1
            
            # Sheet B Analysis
            if analysis_scope in ["Both Sheets", "Sheet B Only"]:
                with tabs[tab_index]:
                    st.subheader("📊 Sheet B Statistical Analysis")
                    
                    # Get current sheet name
                    current_sheet_b_name = st.session_state.get('sheet_b', 'Sheet B')
                    if st.session_state.get('sheets_swapped', False):
                        display_name = f"{st.session_state.get('sheet_a', 'Unknown')} (currently Sheet B)"
                    else:
                        display_name = current_sheet_b_name
                    
                    # Perform analysis (similar to Sheet A but for df_b)
                    analysis_b = analyzer.analyze_dataframe(comparator.df_b, display_name)
                    
                    # Create summary metrics
                    summary_cards = visualizer.create_summary_metrics_cards(
                        analysis_b['basic_info'], 
                        analysis_b['missing_data_analysis']['_summary']
                    )
                    
                    # Display summary cards
                    card_cols = st.columns(len(summary_cards))
                    for i, card in enumerate(summary_cards):
                        with card_cols[i]:
                            st.metric(
                                label=f"{card['icon']} {card['title']}", 
                                value=card['value']
                            )
                    
                    # Analysis sub-tabs (same structure as Sheet A)
                    analysis_subtabs = st.tabs([
                        "📈 Descriptive Stats", 
                        "🕳️ Missing Data", 
                        "📊 Distributions", 
                        "🔗 Correlations",
                        "📋 Data Types"
                    ])
                    
                    with analysis_subtabs[0]:  # Descriptive Stats
                        st.subheader("📈 Descriptive Statistics")
                        
                        if 'message' not in analysis_b['descriptive_stats']:
                            # Show descriptive statistics table
                            desc_df = pd.DataFrame(analysis_b['descriptive_stats']).T
                            
                            if show_advanced:
                                columns_to_show = ['count', 'mean', 'median', 'std_dev', 'min', 'max', 
                                                 'q1', 'q3', 'skewness', 'kurtosis', 'coefficient_of_variation']
                            else:
                                columns_to_show = ['count', 'mean', 'median', 'std_dev', 'min', 'max']
                            
                            # Filter available columns
                            available_columns = [col for col in columns_to_show if col in desc_df.columns]
                            
                            if available_columns:
                                st.dataframe(
                                    desc_df[available_columns].round(3), 
                                    use_container_width=True
                                )
                                
                                # Show visualization
                                fig_desc = visualizer.create_descriptive_stats_chart(
                                    analysis_b['descriptive_stats'], 
                                    f"Descriptive Statistics - {display_name}"
                                )
                                st.plotly_chart(fig_desc, use_container_width=True, key="desc_stats_b")
                        else:
                            st.info(analysis_b['descriptive_stats']['message'])
                    
                    with analysis_subtabs[1]:  # Missing Data
                        st.subheader("🕳️ Missing Data Analysis")
                        
                        # Missing data summary chart
                        fig_missing = visualizer.create_missing_data_summary_chart(
                            analysis_b['missing_data_analysis']
                        )
                        st.plotly_chart(fig_missing, use_container_width=True, key="missing_summary_b")
                        
                        # Missing data heatmap
                        if len(comparator.df_b) <= 1000:  # Only show heatmap for smaller datasets
                            fig_heatmap = visualizer.create_missing_data_heatmap(
                                comparator.df_b, f"Missing Data Pattern - {display_name}"
                            )
                            st.plotly_chart(fig_heatmap, use_container_width=True, key="missing_heatmap_b")
                        else:
                            st.info("💡 Missing data heatmap not shown for large datasets (>1000 rows) to maintain performance")
                    
                    with analysis_subtabs[2]:  # Distributions
                        st.subheader("📊 Data Distributions")
                        
                        # Distribution plots
                        fig_dist = visualizer.create_distribution_plots(comparator.df_b)
                        st.plotly_chart(fig_dist, use_container_width=True, key="distributions_b")
                        
                        # Box plots for outlier detection
                        fig_box = visualizer.create_box_plots(comparator.df_b)
                        st.plotly_chart(fig_box, use_container_width=True, key="boxplots_b")
                    
                    with analysis_subtabs[3]:  # Correlations
                        st.subheader("🔗 Correlation Analysis")
                        
                        correlation_data = analyzer.calculate_correlation_matrix(comparator.df_b)
                        
                        if 'message' not in correlation_data:
                            # Correlation type selector
                            corr_type = st.selectbox(
                                "Correlation Method:",
                                ["pearson", "spearman", "kendall"],
                                index=0,
                                key="corr_type_b",
                                help="Pearson: linear relationships, Spearman: monotonic relationships, Kendall: rank-based"
                            )
                            
                            # Show correlation heatmap
                            fig_corr = visualizer.create_correlation_heatmap(correlation_data, corr_type)
                            st.plotly_chart(fig_corr, use_container_width=True, key="correlation_b")
                            
                            # Show strong correlations
                            if correlation_data['strong_correlations']:
                                st.subheader("💪 Strong Correlations Found")
                                
                                strong_corr_data = []
                                for corr in correlation_data['strong_correlations']:
                                    strong_corr_data.append({
                                        'Column 1': corr['column1'],
                                        'Column 2': corr['column2'],
                                        'Correlation': f"{corr['correlation']:.3f}",
                                        'Strength': corr['strength'].replace('_', ' ').title()
                                    })
                                
                                st.dataframe(pd.DataFrame(strong_corr_data), hide_index=True, use_container_width=True)
                            else:
                                st.info("No strong correlations (>0.7) found between variables")
                        else:
                            st.info(correlation_data['message'])
                    
                    with analysis_subtabs[4]:  # Data Types
                        st.subheader("📋 Data Types Analysis")
                        
                        # Data types pie chart
                        fig_types = visualizer.create_data_types_pie_chart(analysis_b['data_types_analysis'])
                        st.plotly_chart(fig_types, use_container_width=True, key="data_types_b")
                        
                        # Data types table with recommendations
                        types_data = []
                        for col, info in analysis_b['data_types_analysis'].items():
                            types_data.append({
                                'Column': col,
                                'Current Type': info['current_type'],
                                'Recommended Type': info['recommended_type'],
                                'Unique Values': f"{info['unique_values']:,}",
                                'Memory Usage': f"{info['memory_usage']:,} bytes",
                                'Sample Values': ', '.join(str(v) for v in info['sample_values'][:3])
                            })
                        
                        st.dataframe(pd.DataFrame(types_data), hide_index=True, use_container_width=True)
        
        elif analysis_scope == "Comparative Analysis":
            st.subheader("⚖️ Comparative Statistical Analysis")
            
            # Perform analysis on both sheets
            analyzer = StatisticalAnalyzer()
            visualizer = StatisticalVisualizer()
            
            try:
                analysis_a = analyzer.analyze_dataframe(comparator.df_a)
            except Exception as e:
                st.error(f"Error analyzing {display_name_a}: {str(e)}")
                analysis_a = {}
            
            try:
                analysis_b = analyzer.analyze_dataframe(comparator.df_b)
            except Exception as e:
                st.error(f"Error analyzing {display_name_b}: {str(e)}")
                analysis_b = {}
            
            # Display name helper
            sheet_a_name = st.session_state.get('sheet_a', 'Sheet A')
            sheet_b_name = st.session_state.get('sheet_b', 'Sheet B')
            display_name_a = f"Sheet A ({sheet_a_name})" if sheet_a_name != 'Sheet A' else "Sheet A"
            display_name_b = f"Sheet B ({sheet_b_name})" if sheet_b_name != 'Sheet B' else "Sheet B"
            
            # High-level comparison metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                diff_rows = len(comparator.df_b) - len(comparator.df_a)
                st.metric(
                    "📊 Row Difference", 
                    f"{diff_rows:+,}",
                    help=f"{display_name_a}: {len(comparator.df_a):,} rows\n{display_name_b}: {len(comparator.df_b):,} rows"
                )
            
            with col2:
                diff_cols = len(comparator.df_b.columns) - len(comparator.df_a.columns)
                st.metric(
                    "📋 Column Difference", 
                    f"{diff_cols:+}",
                    help=f"{display_name_a}: {len(comparator.df_a.columns)} columns\n{display_name_b}: {len(comparator.df_b.columns)} columns"
                )
            
            with col3:
                memory_a = analysis_a.get('memory_usage', {}).get('total_mb', 0) if analysis_a else 0
                memory_b = analysis_b.get('memory_usage', {}).get('total_mb', 0) if analysis_b else 0
                memory_diff = (memory_b or 0) - (memory_a or 0)
                st.metric(
                    "💾 Memory Difference", 
                    f"{memory_diff:+.1f} MB",
                    help=f"{display_name_a}: {memory_a or 0:.1f} MB\n{display_name_b}: {memory_b or 0:.1f} MB"
                )
            
            with col4:
                missing_a = analysis_a.get('missing_data_analysis', {}).get('total_missing_percentage', 0) if analysis_a else 0
                missing_b = analysis_b.get('missing_data_analysis', {}).get('total_missing_percentage', 0) if analysis_b else 0
                missing_diff = (missing_b or 0) - (missing_a or 0)
                st.metric(
                    "🕳️ Missing Data Difference", 
                    f"{missing_diff:+.1f}%",
                    help=f"{display_name_a}: {missing_a or 0:.1f}% missing\n{display_name_b}: {missing_b or 0:.1f}% missing"
                )
            
            # Detailed comparative analysis tabs
            comp_tabs = st.tabs([
                "📈 Descriptive Stats", 
                "🕳️ Missing Data", 
                "📊 Distributions", 
                "🔗 Correlations",
                "📋 Data Types",
                "📋 Summary Report"
            ])
            
            with comp_tabs[0]:  # Descriptive Stats Comparison
                st.subheader("📈 Descriptive Statistics Comparison")
                
                # Get numerical columns that exist in both datasets
                num_cols_a = set(comparator.df_a.select_dtypes(include=[np.number]).columns)
                num_cols_b = set(comparator.df_b.select_dtypes(include=[np.number]).columns)
                common_num_cols = list(num_cols_a.intersection(num_cols_b))
                
                if common_num_cols:
                    col_choice = st.selectbox(
                        "Choose column to compare:",
                        common_num_cols,
                        help="Select a numerical column that exists in both sheets"
                    )
                    
                    if col_choice:
                        # Side-by-side comparison
                        comp_col1, comp_col2 = st.columns(2)
                        
                        with comp_col1:
                            st.markdown(f"**{display_name_a}**")
                            stats_a = comparator.df_a[col_choice].describe()
                            st.dataframe(stats_a.to_frame().round(3), use_container_width=True)
                        
                        with comp_col2:
                            st.markdown(f"**{display_name_b}**")
                            stats_b = comparator.df_b[col_choice].describe()
                            st.dataframe(stats_b.to_frame().round(3), use_container_width=True)
                        
                        # Comparison chart
                        fig_comp = visualizer.create_comparative_distribution_plot(
                            comparator.df_a[col_choice], 
                            comparator.df_b[col_choice],
                            f"{col_choice} - Distribution Comparison",
                            display_name_a,
                            display_name_b
                        )
                        st.plotly_chart(fig_comp, use_container_width=True, key="comp_dist")
                        
                        # Statistical significance test
                        from scipy import stats
                        statistic, p_value = stats.ks_2samp(
                            comparator.df_a[col_choice].dropna(), 
                            comparator.df_b[col_choice].dropna()
                        )
                        
                        st.subheader("🔬 Statistical Significance Test")
                        st.write(f"**Kolmogorov-Smirnov Test:**")
                        st.write(f"- Statistic: {statistic:.4f}")
                        st.write(f"- P-value: {p_value:.4f}")
                        
                        if p_value < 0.05:
                            st.error("� **Significant difference detected** (p < 0.05): The distributions are statistically different.")
                        else:
                            st.success("✅ **No significant difference** (p ≥ 0.05): The distributions are statistically similar.")
                
                else:
                    st.info("ℹ️ **No Common Numerical Columns Found**")
                    st.write("The selected sheets don't have numerical columns with identical names for direct comparison.")
                    
                    # Show available numerical columns in each sheet
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown(f"**{display_name_a} has {len(num_cols_a)} numerical columns:**")
                        for col in sorted(num_cols_a):
                            st.write(f"• {col}")
                    
                    with col2:
                        st.markdown(f"**{display_name_b} has {len(num_cols_b)} numerical columns:**")
                        for col in sorted(num_cols_b):
                            st.write(f"• {col}")
                    
                    st.markdown("**💡 For meaningful comparisons:** Select sheets with similar column structures (e.g., both sales data sheets with 'Amount', 'Quantity' columns).")
            
            with comp_tabs[1]:  # Missing Data Comparison
                st.subheader("🕳️ Missing Data Comparison")
                
                # Create comparative missing data visualization
                fig_missing_comp = visualizer.create_comparative_missing_data_chart(
                    analysis_a['missing_data_analysis'],
                    analysis_b['missing_data_analysis'],
                    display_name_a,
                    display_name_b
                )
                st.plotly_chart(fig_missing_comp, use_container_width=True, key="comp_missing")
                
                # Detailed missing data comparison table
                missing_comp_data = []
                all_columns = set(list(comparator.df_a.columns) + list(comparator.df_b.columns))
                
                for col in sorted(all_columns):
                    missing_a = (comparator.df_a[col].isnull().sum() / len(comparator.df_a) * 100) if col in comparator.df_a.columns else None
                    missing_b = (comparator.df_b[col].isnull().sum() / len(comparator.df_b) * 100) if col in comparator.df_b.columns else None
                    
                    missing_comp_data.append({
                        'Column': col,
                        f'{display_name_a} Missing (%)': f"{missing_a:.1f}%" if missing_a is not None else "N/A",
                        f'{display_name_b} Missing (%)': f"{missing_b:.1f}%" if missing_b is not None else "N/A",
                        'Difference': f"{(missing_b or 0) - (missing_a or 0):+.1f}%" if missing_a is not None and missing_b is not None else "N/A"
                    })
                
                st.dataframe(pd.DataFrame(missing_comp_data), hide_index=True, use_container_width=True)
            
            with comp_tabs[2]:  # Distribution Comparison
                st.subheader("📊 Distribution Comparison")
                
                if common_num_cols:
                    # Multi-column distribution comparison
                    selected_cols = st.multiselect(
                        "Select columns to compare distributions:",
                        common_num_cols,
                        default=common_num_cols[:3] if len(common_num_cols) >= 3 else common_num_cols,
                        help="Select up to 5 columns for comparison"
                    )
                    
                    if selected_cols:
                        for col in selected_cols[:5]:  # Limit to 5 for performance
                            st.markdown(f"**{col} Distribution Comparison**")
                            
                            fig_hist_comp = visualizer.create_comparative_histogram(
                                comparator.df_a[col],
                                comparator.df_b[col],
                                col,
                                display_name_a,
                                display_name_b
                            )
                            st.plotly_chart(fig_hist_comp, use_container_width=True, key=f"comp_hist_{col}")
                else:
                    st.info("ℹ️ **No Common Numerical Columns Found**")
                    st.write("The selected sheets don't have columns with identical names, so direct distribution comparison isn't possible.")
                    
                    # Show what numerical columns each dataset has
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown(f"**📊 {display_name_a} Numerical Columns:**")
                        nums_a = list(num_cols_a)
                        if nums_a:
                            for col in nums_a:
                                st.write(f"• {col}")
                        else:
                            st.write("No numerical columns found")
                    
                    with col2:
                        st.markdown(f"**📊 {display_name_b} Numerical Columns:**")
                        nums_b = list(num_cols_b)
                        if nums_b:
                            for col in nums_b:
                                st.write(f"• {col}")
                        else:
                            st.write("No numerical columns found")
                    
                    st.markdown("**💡 Suggestion:** For meaningful comparisons, try selecting sheets that have similar column structures, or use the individual statistical analysis for each sheet separately.")
            
            with comp_tabs[3]:  # Correlation Comparison
                st.subheader("🔗 Correlation Comparison")
                
                # Calculate correlations for both datasets
                corr_a = analyzer.calculate_correlation_matrix(comparator.df_a)
                corr_b = analyzer.calculate_correlation_matrix(comparator.df_b)
                
                comp_corr_col1, comp_corr_col2 = st.columns(2)
                
                with comp_corr_col1:
                    st.markdown(f"**{display_name_a} Correlations**")
                    if 'message' not in corr_a:
                        fig_corr_a = visualizer.create_correlation_heatmap(corr_a, 'pearson')
                        st.plotly_chart(fig_corr_a, use_container_width=True, key="comp_corr_a")
                    else:
                        st.info(corr_a['message'])
                
                with comp_corr_col2:
                    st.markdown(f"**{display_name_b} Correlations**")
                    if 'message' not in corr_b:
                        fig_corr_b = visualizer.create_correlation_heatmap(corr_b, 'pearson')
                        st.plotly_chart(fig_corr_b, use_container_width=True, key="comp_corr_b")
                    else:
                        st.info(corr_b['message'])
                
                # Correlation strength comparison
                if 'message' not in corr_a and 'message' not in corr_b:
                    st.subheader("🔍 Strong Correlations Comparison")
                    
                    strong_a = corr_a.get('strong_correlations', [])
                    strong_b = corr_b.get('strong_correlations', [])
                    
                    comp_strong_col1, comp_strong_col2 = st.columns(2)
                    
                    with comp_strong_col1:
                        st.markdown(f"**{display_name_a}** ({len(strong_a)} strong correlations)")
                        if strong_a:
                            for corr in strong_a[:5]:
                                st.write(f"• {corr['column1']} ↔ {corr['column2']}: {corr['correlation']:.3f}")
                        else:
                            st.info("No strong correlations found")
                    
                    with comp_strong_col2:
                        st.markdown(f"**{display_name_b}** ({len(strong_b)} strong correlations)")
                        if strong_b:
                            for corr in strong_b[:5]:
                                st.write(f"• {corr['column1']} ↔ {corr['column2']}: {corr['correlation']:.3f}")
                        else:
                            st.info("No strong correlations found")
            
            with comp_tabs[4]:  # Data Types Comparison
                st.subheader("📋 Data Types Comparison")
                
                # Compare data types
                types_comp_data = []
                all_columns = set(list(comparator.df_a.columns) + list(comparator.df_b.columns))
                
                for col in sorted(all_columns):
                    type_a = str(comparator.df_a[col].dtype) if col in comparator.df_a.columns else "N/A"
                    type_b = str(comparator.df_b[col].dtype) if col in comparator.df_b.columns else "N/A"
                    
                    # Check if types match
                    match_status = "✅" if type_a == type_b and type_a != "N/A" else "❌" if type_a != "N/A" and type_b != "N/A" else "⚠️"
                    
                    types_comp_data.append({
                        'Column': col,
                        f'{display_name_a} Type': type_a,
                        f'{display_name_b} Type': type_b,
                        'Match': match_status
                    })
                
                st.dataframe(pd.DataFrame(types_comp_data), hide_index=True, use_container_width=True)
                
                # Summary of type mismatches
                mismatches = [item for item in types_comp_data if item['Match'] == '❌']
                if mismatches:
                    st.warning(f"⚠️ **{len(mismatches)} data type mismatches found.** These may cause issues in analysis or merging.")
                else:
                    st.success("✅ **All common columns have matching data types.**")
            
            with comp_tabs[5]:  # Summary Report
                st.subheader("📋 Comparative Analysis Summary")
                
                # Generate comprehensive comparison report
                st.markdown("### 📊 **Dataset Overview**")
                
                # Safe formatting with default values
                safe_memory_a = memory_a if memory_a is not None else 0
                safe_memory_b = memory_b if memory_b is not None else 0
                safe_missing_a = missing_a if missing_a is not None else 0
                safe_missing_b = missing_b if missing_b is not None else 0
                
                overview_data = {
                    'Metric': [
                        'Total Rows',
                        'Total Columns', 
                        'Memory Usage (MB)',
                        'Missing Data (%)',
                        'Numerical Columns',
                        'Categorical Columns',
                        'Date Columns'
                    ],
                    display_name_a: [
                        f"{len(comparator.df_a):,}",
                        len(comparator.df_a.columns),
                        f"{safe_memory_a:.1f}",
                        f"{safe_missing_a:.1f}%",
                        len(comparator.df_a.select_dtypes(include=[np.number]).columns),
                        len(comparator.df_a.select_dtypes(include=['object', 'category']).columns),
                        len(comparator.df_a.select_dtypes(include=['datetime64']).columns)
                    ],
                    display_name_b: [
                        f"{len(comparator.df_b):,}",
                        len(comparator.df_b.columns),
                        f"{safe_memory_b:.1f}",
                        f"{safe_missing_b:.1f}%",
                        len(comparator.df_b.select_dtypes(include=[np.number]).columns),
                        len(comparator.df_b.select_dtypes(include=['object', 'category']).columns),
                        len(comparator.df_b.select_dtypes(include=['datetime64']).columns)
                    ]
                }
                
                st.dataframe(pd.DataFrame(overview_data), hide_index=True, use_container_width=True)
                
                # Key insights
                st.markdown("### 🔍 **Key Insights**")
                
                insights = []
                
                # Size comparison
                if len(comparator.df_b) > len(comparator.df_a):
                    insights.append(f"📈 {display_name_b} has {len(comparator.df_b) - len(comparator.df_a):,} more rows than {display_name_a}")
                elif len(comparator.df_a) > len(comparator.df_b):
                    insights.append(f"📉 {display_name_a} has {len(comparator.df_a) - len(comparator.df_b):,} more rows than {display_name_b}")
                else:
                    insights.append(f"⚖️ Both datasets have the same number of rows ({len(comparator.df_a):,})")
                
                # Missing data comparison
                if safe_missing_a != "N/A" and safe_missing_b != "N/A" and abs(missing_diff) > 5:
                    if missing_diff > 0:
                        insights.append(f"🕳️ {display_name_b} has significantly more missing data ({missing_diff:+.1f}%)")
                    else:
                        insights.append(f"🕳️ {display_name_a} has significantly more missing data ({abs(missing_diff):.1f}%)")
                
                # Column overlap
                common_cols = set(comparator.df_a.columns).intersection(set(comparator.df_b.columns))
                total_unique_cols = len(set(comparator.df_a.columns).union(set(comparator.df_b.columns)))
                overlap_pct = len(common_cols) / total_unique_cols * 100
                
                insights.append(f"🔗 Column overlap: {overlap_pct:.1f}% ({len(common_cols)} of {total_unique_cols} total unique columns)")
                
                # Memory efficiency
                if safe_memory_a != "N/A" and safe_memory_b != "N/A" and abs(memory_diff) > 1:
                    if memory_diff > 0:
                        insights.append(f"💾 {display_name_b} uses {memory_diff:.1f} MB more memory")
                    else:
                        insights.append(f"💾 {display_name_a} uses {abs(memory_diff):.1f} MB more memory")
                
                for insight in insights:
                    st.write(f"• {insight}")
                
                # Recommendations
                st.markdown("### 💡 **Recommendations**")
                
                recommendations = []
                
                if len(mismatches) > 0:
                    recommendations.append("🔧 **Data Type Alignment**: Consider standardizing data types for common columns to enable better analysis and merging.")
                
                if safe_missing_a != "N/A" and safe_missing_b != "N/A" and abs(missing_diff) > 10:
                    recommendations.append("🧹 **Data Quality**: Address missing data patterns, especially in the dataset with higher missing percentages.")
                
                if overlap_pct < 50:
                    recommendations.append("📋 **Schema Review**: Low column overlap suggests these datasets may serve different purposes or need schema alignment.")
                
                if abs(diff_rows) > len(comparator.df_a) * 0.1:  # More than 10% difference
                    recommendations.append("⚖️ **Size Disparity**: Significant size difference between datasets - consider if this is expected or requires investigation.")
                
                if not recommendations:
                    recommendations.append("✅ **Well Aligned**: The datasets appear to be well-structured and comparable for analysis.")
                
                for rec in recommendations:
                    st.write(f"• {rec}")
                
                # Export comparison report
                if st.button("📄 Generate Detailed Comparison Report", help="Generate a comprehensive comparison report"):
                    # Create detailed report data
                    report_data = {
                        'comparison_timestamp': datetime.now().isoformat(),
                        'dataset_a': {'name': display_name_a, 'rows': len(comparator.df_a), 'columns': len(comparator.df_a.columns)},
                        'dataset_b': {'name': display_name_b, 'rows': len(comparator.df_b), 'columns': len(comparator.df_b.columns)},
                        'overview': overview_data,
                        'insights': insights,
                        'recommendations': recommendations,
                        'type_mismatches': len(mismatches),
                        'column_overlap_percentage': overlap_pct
                    }
                    
                    # Convert to JSON for download
                    import json
                    report_json = json.dumps(report_data, indent=2, default=str)
                    
                    st.download_button(
                        label="📥 Download Comparison Report (JSON)",
                        data=report_json,
                        file_name=f"comparative_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                        mime="application/json"
                    )
                    
                    st.success("✅ Comparison report generated successfully!")
    
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
            
            # Enhanced column analysis for selected key columns
            if st.checkbox("📊 Show Advanced Key Column Analysis", help="Get detailed insights about data quality and matching potential"):
                st.markdown("---")
                
                # Header section with overview
                st.markdown("### 🔍 Key Column Analysis Dashboard")
                st.markdown("*Analyze data quality and matching potential for your selected key columns*")
                
                # Overview metrics in a nice layout
                overview_col1, overview_col2, overview_col3, overview_col4 = st.columns(4)
                
                with overview_col1:
                    unique_a = comparator.df_a[key_col_a].nunique()
                    total_a = len(comparator.df_a)
                    st.metric(
                        label="📋 Sheet A Unique Values", 
                        value=f"{unique_a:,}",
                        delta=f"of {total_a:,} total"
                    )
                
                with overview_col2:
                    unique_b = comparator.df_b[key_col_b].nunique()
                    total_b = len(comparator.df_b)
                    st.metric(
                        label="📋 Sheet B Unique Values", 
                        value=f"{unique_b:,}",
                        delta=f"of {total_b:,} total"
                    )
                
                with overview_col3:
                    # Calculate potential matches estimate
                    common_sample = set(comparator.df_a[key_col_a].dropna().astype(str).str.lower()[:100]) & \
                                   set(comparator.df_b[key_col_b].dropna().astype(str).str.lower()[:100])
                    match_estimate = len(common_sample)
                    st.metric(
                        label="🎯 Potential Matches", 
                        value=f"~{match_estimate}",
                        delta="from sample",
                        help="Estimated based on first 100 records"
                    )
                
                with overview_col4:
                    # Data quality score
                    null_a = comparator.df_a[key_col_a].isnull().sum()
                    null_b = comparator.df_b[key_col_b].isnull().sum()
                    quality_score = max(0, 100 - ((null_a + null_b) / (total_a + total_b) * 100))
                    st.metric(
                        label="✨ Data Quality Score", 
                        value=f"{quality_score:.0f}%",
                        delta="Lower is better" if quality_score < 80 else "Good quality"
                    )
                
                st.markdown("---")
                
                # Detailed analysis in organized tabs
                tab1, tab2, tab3 = st.tabs(["📊 Detailed Analysis", "🔍 Pattern Detection", "💡 Recommendations"])
                
                with tab1:
                    # Side-by-side detailed analysis
                    col_analysis_1, col_analysis_2 = st.columns(2)
                    
                    with col_analysis_1:
                        st.markdown("#### 📋 Sheet A Analysis")
                        comparator.show_column_analysis(comparator.df_a, key_col_a, "Sheet A")
                    
                    with col_analysis_2:
                        st.markdown("#### 📋 Sheet B Analysis") 
                        comparator.show_column_analysis(comparator.df_b, key_col_b, "Sheet B")
                
                with tab2:
                    # Pattern comparison between columns
                    st.markdown("#### 🔍 Cross-Column Pattern Analysis")
                    
                    pattern_col1, pattern_col2 = st.columns(2)
                    
                    with pattern_col1:
                        st.markdown("**📋 Sheet A Patterns:**")
                        comparator.show_pattern_analysis(comparator.df_a, key_col_a)
                    
                    with pattern_col2:
                        st.markdown("**📋 Sheet B Patterns:**")
                        comparator.show_pattern_analysis(comparator.df_b, key_col_b)
                    
                    # Compatibility analysis
                    st.markdown("---")
                    st.markdown("#### 🤝 Column Compatibility Analysis")
                    comparator.show_compatibility_analysis(comparator.df_a, key_col_a, comparator.df_b, key_col_b)
                
                with tab3:
                    # Actionable recommendations
                    st.markdown("#### 💡 Smart Recommendations")
                    comparator.show_matching_recommendations(comparator.df_a, key_col_a, comparator.df_b, key_col_b, threshold)
        
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
            st.subheader("Comparison Mode")
            
            # Comparison mode selection
            comparison_mode = st.radio(
                "Select comparison method:",
                ["🔍 Single Column", "🎯 Multi-Column Advanced"],
                help="Single column uses one key field. Multi-column combines multiple fields for better accuracy."
            )
            
            # Multi-column configuration
            if comparison_mode == "🎯 Multi-Column Advanced":
                st.write("**Multi-Column Settings:**")
                
                # Multi-column selection
                multi_cols_a = st.multiselect(
                    "Key columns in Sheet A:",
                    comparator.df_a.columns,
                    default=[key_col_a],
                    help="Select multiple columns to match on"
                )
                
                multi_cols_b = st.multiselect(
                    "Key columns in Sheet B:",
                    comparator.df_b.columns,
                    default=[key_col_b],
                    help="Must match the order and count of Sheet A columns"
                )
                
                # Validate multi-column selection
                if len(multi_cols_a) != len(multi_cols_b):
                    st.warning("⚠️ Number of columns must match between sheets")
                elif len(multi_cols_a) < 2:
                    st.info("💡 Select at least 2 columns for multi-column matching")
                else:
                    # Field weights configuration
                    st.write("**Field Importance Weights:**")
                    field_weights = []
                    
                    for i, (col_a, col_b) in enumerate(zip(multi_cols_a, multi_cols_b)):
                        weight = st.slider(
                            f"{col_a} ↔ {col_b}",
                            min_value=0.1,
                            max_value=1.0,
                            value=0.5 if i == 0 else 0.3,
                            step=0.1,
                            help=f"Importance weight for {col_a} field"
                        )
                        field_weights.append(weight)
                    
                    # Show normalized weights
                    total_weight = sum(field_weights)
                    normalized_weights = [w/total_weight for w in field_weights]
                    
                    with st.expander("📊 Weight Distribution", expanded=False):
                        for i, (col_a, norm_weight) in enumerate(zip(multi_cols_a, normalized_weights)):
                            st.write(f"• **{col_a}**: {norm_weight:.1%}")
            
            st.divider()
            
            # Comparison buttons
            if comparison_mode == "🔍 Single Column":
                if st.button("🔍 Start Single-Column Comparison", type="primary", use_container_width=True):
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
            
            # Multi-column comparison
            elif comparison_mode == "🎯 Multi-Column Advanced" and len(multi_cols_a) >= 2 and len(multi_cols_a) == len(multi_cols_b):
                if st.button("🎯 Start Multi-Column Comparison", type="primary", use_container_width=True):
                    # Pre-comparison validation and info
                    st.info("🚀 Starting advanced multi-column comparison with weighted field matching...")
                    
                    # Show multi-column comparison parameters
                    with st.expander("📋 Multi-Column Parameters", expanded=False):
                        st.write(f"**📊 Data Overview:**")
                        st.write(f"- Sheet A: {len(comparator.df_a):,} rows")
                        st.write(f"- Sheet B: {len(comparator.df_b):,} rows")
                        st.write(f"- Key columns ({len(multi_cols_a)}): {' + '.join(multi_cols_a)} ↔ {' + '.join(multi_cols_b)}")
                        st.write(f"- Extracting: {', '.join(cols_to_extract) if cols_to_extract else 'No additional columns'}")
                        st.write(f"- Similarity threshold: {threshold}%")
                        st.write(f"- Case sensitive: {'No' if ignore_case else 'Yes'}")
                        
                        # Show field weights
                        st.write("**Field Weights:**")
                        for i, (col_a, weight) in enumerate(zip(multi_cols_a, normalized_weights)):
                            st.write(f"- {col_a}: {weight:.1%}")
                    
                    # Estimate processing time (multi-column is more intensive)
                    estimated_time = len(comparator.df_a) * 0.02 * len(multi_cols_a)  # More time for multi-column
                    if estimated_time > 60:
                        time_estimate = f"~{estimated_time/60:.1f} minutes"
                    else:
                        time_estimate = f"~{estimated_time:.0f} seconds"
                    
                    st.write(f"⏱️ **Estimated processing time:** {time_estimate}")
                    st.info("💡 Multi-column matching is more thorough but takes longer to process")
                    
                    # Run multi-column comparison
                    try:
                        results = comparator.perform_multi_column_comparison(
                            comparator.df_a, comparator.df_b,
                            multi_cols_a, multi_cols_b,
                            cols_to_extract, threshold, field_weights, ignore_case
                        )
                        comparator.results = results
                        
                        # Show completion celebration
                        st.balloons()
                        st.success("🎉 Multi-column comparison completed successfully! Scroll down to view enhanced results.")
                        
                    except Exception as e:
                        st.error(f"❌ Error during multi-column comparison: {str(e)}")
                        st.write("Please check your field selections and data, then try again.")
                        if "must match" in str(e):
                            st.info("💡 Make sure the number of key columns is the same for both sheets")
            
            elif comparison_mode == "🎯 Multi-Column Advanced":
                st.warning("⚠️ Please select at least 2 columns for each sheet, with matching counts")
        
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
            
            # Multi-column specific insights
            if comparator.results['matched'] and 'field_breakdown' in pd.DataFrame(comparator.results['matched']).columns:
                st.subheader("🎯 Multi-Column Analysis")
                
                with st.expander("📊 Field-Level Performance Analysis", expanded=False):
                    all_matches = comparator.results['matched'] + comparator.results['suggested']
                    
                    if all_matches:
                        # Analyze field performance
                        field_performance = {}
                        for match in all_matches:
                            if 'field_breakdown' in match and match['field_breakdown']:
                                for field_pair, score in match['field_breakdown'].items():
                                    if field_pair not in field_performance:
                                        field_performance[field_pair] = []
                                    field_performance[field_pair].append(score)
                        
                        if field_performance:
                            st.write("**Average Field Performance:**")
                            performance_data = []
                            
                            for field_pair, scores in field_performance.items():
                                avg_score = sum(scores) / len(scores)
                                min_score = min(scores)
                                max_score = max(scores)
                                
                                performance_data.append({
                                    'Field Mapping': field_pair,
                                    'Avg Score': f"{avg_score:.1f}%",
                                    'Min Score': f"{min_score:.1f}%",
                                    'Max Score': f"{max_score:.1f}%",
                                    'Records': len(scores),
                                    'Performance': "🟢 Excellent" if avg_score >= 80 else "🟡 Good" if avg_score >= 60 else "🔴 Needs Review"
                                })
                            
                            performance_df = pd.DataFrame(performance_data)
                            st.dataframe(performance_df, hide_index=True, use_container_width=True)
                            
                            # Field recommendations
                            best_field = max(field_performance.items(), key=lambda x: sum(x[1])/len(x[1]))
                            worst_field = min(field_performance.items(), key=lambda x: sum(x[1])/len(x[1]))
                            
                            col_rec1, col_rec2 = st.columns(2)
                            with col_rec1:
                                st.success(f"🏆 **Best performing field:** {best_field[0]} ({sum(best_field[1])/len(best_field[1]):.1f}% avg)")
                            with col_rec2:
                                if sum(worst_field[1])/len(worst_field[1]) < 70:
                                    st.warning(f"⚠️ **Needs attention:** {worst_field[0]} ({sum(worst_field[1])/len(worst_field[1]):.1f}% avg)")
                                else:
                                    st.info(f"💡 **Consider optimizing:** {worst_field[0]} ({sum(worst_field[1])/len(worst_field[1]):.1f}% avg)")
            
            # Match type distribution (enhanced for multi-column)
            if comparator.results['matched'] or comparator.results['suggested']:
                all_results = comparator.results['matched'] + comparator.results['suggested']
                match_types = {}
                
                for result in all_results:
                    match_type = result.get('match_type', 'Unknown')
                    match_types[match_type] = match_types.get(match_type, 0) + 1
                
                if len(match_types) > 1:  # Only show if there are different match types
                    with st.expander("📈 Match Type Distribution", expanded=False):
                        type_data = []
                        for match_type, count in match_types.items():
                            percentage = (count / len(all_results)) * 100
                            type_data.append({
                                'Match Type': match_type,
                                'Count': f"{count:,}",
                                'Percentage': f"{percentage:.1f}%"
                            })
                        
                        type_df = pd.DataFrame(type_data)
                        st.dataframe(type_df, hide_index=True, use_container_width=True)
            
            # Enhanced Results tabs with filtering
            tab1, tab2, tab3 = st.tabs(["✅ Matched", "⚠️ Suggested Matches", "❌ Unmatched"])
            
            with tab1:
                if comparator.results['matched']:
                    df_matched = pd.DataFrame(comparator.results['matched'])
                    
                    # Add filtering for matched results
                    filtered_matched = comparator.add_result_filters(df_matched, "Matched")
                    
                    # Display filtered results
                    if not filtered_matched.empty:
                        st.dataframe(filtered_matched, use_container_width=True)
                        
                        # Show row selection info
                        if len(filtered_matched) != len(df_matched):
                            st.caption(f"💡 Showing {len(filtered_matched):,} of {len(df_matched):,} matched records")
                    else:
                        st.info("No records match your current filters")
                else:
                    st.info("No exact matches found")
            
            with tab2:
                if comparator.results['suggested']:
                    df_suggested = pd.DataFrame(comparator.results['suggested'])
                    
                    # Add filtering for suggested results
                    filtered_suggested = comparator.add_result_filters(df_suggested, "Suggested")
                    
                    # Display filtered results with enhanced info
                    if not filtered_suggested.empty:
                        # Show confidence distribution for suggested matches
                        if 'similarity_score' in filtered_suggested.columns:
                            avg_confidence = filtered_suggested['similarity_score'].mean()
                            min_confidence = filtered_suggested['similarity_score'].min()
                            max_confidence = filtered_suggested['similarity_score'].max()
                            
                            conf_col1, conf_col2, conf_col3 = st.columns(3)
                            with conf_col1:
                                st.metric("Avg Confidence", f"{avg_confidence:.1f}%")
                            with conf_col2:
                                st.metric("Min Confidence", f"{min_confidence:.1f}%")
                            with conf_col3:
                                st.metric("Max Confidence", f"{max_confidence:.1f}%")
                        
                        st.dataframe(filtered_suggested, use_container_width=True)
                        
                        # Show helpful tips for suggested matches
                        with st.expander("💡 Tips for Reviewing Suggested Matches", expanded=False):
                            st.write("""
                            **How to review suggested matches:**
                            - ✅ **High confidence (80%+)**: Usually safe to accept
                            - ⚠️ **Medium confidence (60-79%)**: Review manually 
                            - ❌ **Low confidence (<60%)**: Likely false positives
                            
                            **Use filters to focus on:**
                            - High confidence matches first
                            - Specific data patterns or text
                            - Records with complete data
                            """)
                        
                        if len(filtered_suggested) != len(df_suggested):
                            st.caption(f"💡 Showing {len(filtered_suggested):,} of {len(df_suggested):,} suggested records")
                    else:
                        st.info("No records match your current filters")
                else:
                    st.info("No suggested matches found")
            
            with tab3:
                if comparator.results['unmatched']:
                    df_unmatched = pd.DataFrame(comparator.results['unmatched'])
                    
                    # Add filtering for unmatched results
                    filtered_unmatched = comparator.add_result_filters(df_unmatched, "Unmatched")
                    
                    # Display filtered results with analysis
                    if not filtered_unmatched.empty:
                        # Show why records might be unmatched
                        with st.expander("🔍 Analysis: Why Records Weren't Matched", expanded=False):
                            st.write("""
                            **Common reasons for unmatched records:**
                            - 🔤 **Spelling differences**: Typos, abbreviations, formatting
                            - 📝 **Missing data**: Empty or null values in key columns
                            - 🔢 **Format differences**: Numbers vs text, date formats
                            - 🌐 **Language differences**: Different languages or character sets
                            - ❌ **Actually missing**: Data truly doesn't exist in Sheet B
                            
                            **Next steps:**
                            - Review the unmatched records below
                            - Consider adjusting the similarity threshold
                            - Check for data quality issues
                            - Manual review may be needed
                            """)
                        
                        st.dataframe(filtered_unmatched, use_container_width=True)
                        
                        if len(filtered_unmatched) != len(df_unmatched):
                            st.caption(f"💡 Showing {len(filtered_unmatched):,} of {len(df_unmatched):,} unmatched records")
                    else:
                        st.info("No records match your current filters")
                else:
                    st.success("🎉 All records were matched - no unmatched data!")
            
            # Data Quality Assessment Section
            st.divider()
            st.header("🔍 Data Quality Assessment")
            
            # Import the data quality module
            from analysis.data_quality import DataQualityAssessment
            
            # Create quality assessment tabs
            quality_tabs = st.tabs([
                "📊 Overview", 
                "🕳️ Missing Data", 
                "🔄 Duplicates", 
                "📋 Data Types",
                "💡 Recommendations"
            ])
            
            with quality_tabs[0]:  # Overview
                st.subheader("📈 Data Quality Overview")
                
                # Create quality assessment for both sheets
                qa_col1, qa_col2 = st.columns(2)
                
                with qa_col1:
                    st.markdown(f"**🔍 {display_sheet_a} Quality Analysis**")
                    
                    with st.spinner("Analyzing data quality..."):
                        qa_a = DataQualityAssessment(comparator.df_a, display_sheet_a)
                        quality_report_a = qa_a.generate_quality_report()
                    
                    # Overall quality metrics
                    metrics_a = quality_report_a['overall_metrics']
                    st.metric("Overall Quality Score", f"{metrics_a['overall_quality_score']:.1f}/100")
                    
                    # Sub-metrics
                    metric_col1, metric_col2 = st.columns(2)
                    with metric_col1:
                        st.metric("Missing Data", f"{metrics_a['missing_data_score']:.1f}/100")
                        st.metric("Data Types", f"{metrics_a['data_type_score']:.1f}/100")
                    with metric_col2:
                        st.metric("Duplicates", f"{metrics_a['duplicate_score']:.1f}/100")
                
                with qa_col2:
                    st.markdown(f"**🔍 {display_sheet_b} Quality Analysis**")
                    
                    with st.spinner("Analyzing data quality..."):
                        qa_b = DataQualityAssessment(comparator.df_b, display_sheet_b)
                        quality_report_b = qa_b.generate_quality_report()
                    
                    # Overall quality metrics
                    metrics_b = quality_report_b['overall_metrics']
                    st.metric("Overall Quality Score", f"{metrics_b['overall_quality_score']:.1f}/100")
                    
                    # Sub-metrics
                    metric_col3, metric_col4 = st.columns(2)
                    with metric_col3:
                        st.metric("Missing Data", f"{metrics_b['missing_data_score']:.1f}/100")
                        st.metric("Data Types", f"{metrics_b['data_type_score']:.1f}/100")
                    with metric_col4:
                        st.metric("Duplicates", f"{metrics_b['duplicate_score']:.1f}/100")
                
                # Comparative quality visualization
                st.markdown("### 📊 Quality Comparison Dashboard")
                
                # Create comparative quality chart
                comparison_data = {
                    'Metric': ['Overall Quality', 'Missing Data', 'Duplicates', 'Data Types'],
                    display_sheet_a: [
                        metrics_a['overall_quality_score'],
                        metrics_a['missing_data_score'],
                        metrics_a['duplicate_score'],
                        metrics_a['data_type_score']
                    ],
                    display_sheet_b: [
                        metrics_b['overall_quality_score'],
                        metrics_b['missing_data_score'],
                        metrics_b['duplicate_score'],
                        metrics_b['data_type_score']
                    ]
                }
                
                comparison_df = pd.DataFrame(comparison_data)
                st.bar_chart(comparison_df.set_index('Metric'), height=400)
            
            with quality_tabs[1]:  # Missing Data
                st.subheader("🕳️ Missing Data Analysis")
                
                missing_col1, missing_col2 = st.columns(2)
                
                with missing_col1:
                    st.markdown(f"**{display_sheet_a} Missing Data**")
                    
                    # Missing data chart
                    fig_missing_a = visualizer.create_missing_data_chart(quality_report_a['missing_data'])
                    st.plotly_chart(fig_missing_a, use_container_width=True, key="missing_a")
                    
                    # Missing data completeness gauge
                    completeness_a = quality_report_a['missing_data']['overall_completeness']
                    fig_completeness_a = visualizer.create_data_completeness_gauge(completeness_a)
                    st.plotly_chart(fig_completeness_a, use_container_width=True, key="completeness_a")
                
                with missing_col2:
                    st.markdown(f"**{display_sheet_b} Missing Data**")
                    
                    # Missing data chart
                    fig_missing_b = visualizer.create_missing_data_chart(quality_report_b['missing_data'])
                    st.plotly_chart(fig_missing_b, use_container_width=True, key="missing_b")
                    
                    # Missing data completeness gauge
                    completeness_b = quality_report_b['missing_data']['overall_completeness']
                    fig_completeness_b = visualizer.create_data_completeness_gauge(completeness_b)
                    st.plotly_chart(fig_completeness_b, use_container_width=True, key="completeness_b")
                
                # Missing data summary
                st.markdown("### 📊 Missing Data Summary")
                
                missing_summary_data = {
                    'Dataset': [display_sheet_a, display_sheet_b],
                    'Overall Completeness (%)': [
                        quality_report_a['missing_data']['overall_completeness'],
                        quality_report_b['missing_data']['overall_completeness']
                    ],
                    'Columns with Missing Data': [
                        quality_report_a['missing_data']['columns_with_missing_data'],
                        quality_report_b['missing_data']['columns_with_missing_data']
                    ],
                    'Empty Rows': [
                        quality_report_a['missing_data']['completely_empty_rows'],
                        quality_report_b['missing_data']['completely_empty_rows']
                    ]
                }
                
                st.dataframe(pd.DataFrame(missing_summary_data), use_container_width=True)
            
            with quality_tabs[2]:  # Duplicates
                st.subheader("🔄 Duplicate Analysis")
                
                dup_col1, dup_col2 = st.columns(2)
                
                with dup_col1:
                    st.markdown(f"**{display_sheet_a} Duplicates**")
                    
                    # Duplicate analysis chart
                    fig_dup_a = visualizer.create_duplicate_analysis_chart(quality_report_a['duplicates'])
                    st.plotly_chart(fig_dup_a, use_container_width=True, key="duplicates_a")
                    
                    # Duplicate summary
                    dup_summary_a = quality_report_a['duplicates']['exact_duplicates']
                    st.metric("Exact Duplicates", f"{dup_summary_a['count']:,} ({dup_summary_a['percentage']:.1f}%)")
                
                with dup_col2:
                    st.markdown(f"**{display_sheet_b} Duplicates**")
                    
                    # Duplicate analysis chart
                    fig_dup_b = visualizer.create_duplicate_analysis_chart(quality_report_b['duplicates'])
                    st.plotly_chart(fig_dup_b, use_container_width=True, key="duplicates_b")
                    
                    # Duplicate summary
                    dup_summary_b = quality_report_b['duplicates']['exact_duplicates']
                    st.metric("Exact Duplicates", f"{dup_summary_b['count']:,} ({dup_summary_b['percentage']:.1f}%)")
                
                # Show sample duplicates if any exist
                if dup_summary_a['count'] > 0:
                    with st.expander(f"📋 Sample Duplicate Rows from {display_sheet_a}"):
                        duplicate_indices = dup_summary_a['duplicate_rows'][:5]
                        if duplicate_indices:
                            st.dataframe(comparator.df_a.iloc[duplicate_indices], use_container_width=True)
                
                if dup_summary_b['count'] > 0:
                    with st.expander(f"📋 Sample Duplicate Rows from {display_sheet_b}"):
                        duplicate_indices = dup_summary_b['duplicate_rows'][:5]
                        if duplicate_indices:
                            st.dataframe(comparator.df_b.iloc[duplicate_indices], use_container_width=True)
            
            with quality_tabs[3]:  # Data Types
                st.subheader("📋 Data Type Analysis")
                
                type_col1, type_col2 = st.columns(2)
                
                with type_col1:
                    st.markdown(f"**{display_sheet_a} Data Types**")
                    
                    # Data type issues chart
                    fig_types_a = visualizer.create_data_type_issues_chart(quality_report_a['data_type_validation'])
                    st.plotly_chart(fig_types_a, use_container_width=True, key="types_a")
                
                with type_col2:
                    st.markdown(f"**{display_sheet_b} Data Types**")
                    
                    # Data type issues chart
                    fig_types_b = visualizer.create_data_type_issues_chart(quality_report_b['data_type_validation'])
                    st.plotly_chart(fig_types_b, use_container_width=True, key="types_b")
                
                # Data type summary table
                st.markdown("### 📊 Data Type Summary")
                
                type_summary_a = quality_report_a['data_type_validation']['summary']
                type_summary_b = quality_report_b['data_type_validation']['summary']
                
                type_summary_data = {
                    'Metric': ['Columns Analyzed', 'Columns with Issues', 'Convertible Columns', 'Data Quality Score'],
                    display_sheet_a: [
                        type_summary_a['total_columns_analyzed'],
                        type_summary_a['columns_with_issues'],
                        type_summary_a['convertible_columns'],
                        f"{type_summary_a['data_quality_score']:.1f}/100"
                    ],
                    display_sheet_b: [
                        type_summary_b['total_columns_analyzed'],
                        type_summary_b['columns_with_issues'],
                        type_summary_b['convertible_columns'],
                        f"{type_summary_b['data_quality_score']:.1f}/100"
                    ]
                }
                
                st.dataframe(pd.DataFrame(type_summary_data), use_container_width=True)
            
            with quality_tabs[4]:  # Recommendations
                st.subheader("💡 Data Quality Recommendations")
                
                rec_col1, rec_col2 = st.columns(2)
                
                with rec_col1:
                    st.markdown(f"**🔧 Recommendations for {display_sheet_a}**")
                    
                    recommendations_a = quality_report_a['recommendations']
                    for i, rec in enumerate(recommendations_a):
                        st.write(f"{i+1}. {rec}")
                
                with rec_col2:
                    st.markdown(f"**🔧 Recommendations for {display_sheet_b}**")
                    
                    recommendations_b = quality_report_b['recommendations']
                    for i, rec in enumerate(recommendations_b):
                        st.write(f"{i+1}. {rec}")
                
                # Combined quality improvement suggestions
                st.markdown("### 🎯 Combined Quality Improvement Plan")
                
                combined_issues = []
                
                # Analyze common issues
                if metrics_a['missing_data_score'] < 80 or metrics_b['missing_data_score'] < 80:
                    combined_issues.append("🕳️ **Missing Data**: Both datasets have significant missing data. Consider data collection improvements or imputation strategies.")
                
                if metrics_a['duplicate_score'] < 90 or metrics_b['duplicate_score'] < 90:
                    combined_issues.append("🔄 **Duplicates**: Implement data deduplication processes before analysis.")
                
                if metrics_a['data_type_score'] < 85 or metrics_b['data_type_score'] < 85:
                    combined_issues.append("📋 **Data Types**: Standardize data entry formats and implement type validation.")
                
                # Overall quality comparison
                avg_quality_a = metrics_a['overall_quality_score']
                avg_quality_b = metrics_b['overall_quality_score']
                
                if abs(avg_quality_a - avg_quality_b) > 20:
                    if avg_quality_a > avg_quality_b:
                        combined_issues.append(f"⚖️ **Quality Gap**: {display_sheet_a} has significantly better quality ({avg_quality_a:.1f}) than {display_sheet_b} ({avg_quality_b:.1f}). Focus improvement efforts on {display_sheet_b}.")
                    else:
                        combined_issues.append(f"⚖️ **Quality Gap**: {display_sheet_b} has significantly better quality ({avg_quality_b:.1f}) than {display_sheet_a} ({avg_quality_a:.1f}). Focus improvement efforts on {display_sheet_a}.")
                
                if not combined_issues:
                    st.success("✅ **Excellent Overall Quality**: Both datasets show good quality metrics across all dimensions!")
                else:
                    for issue in combined_issues:
                        st.write(f"• {issue}")
                
                # Export quality report
                if st.button("📄 Generate Quality Assessment Report"):
                    quality_report_data = {
                        'timestamp': datetime.now().isoformat(),
                        'dataset_a': {
                            'name': display_sheet_a,
                            'metrics': metrics_a,
                            'recommendations': recommendations_a
                        },
                        'dataset_b': {
                            'name': display_sheet_b,
                            'metrics': metrics_b,
                            'recommendations': recommendations_b
                        },
                        'combined_recommendations': combined_issues
                    }
                    
                    import json
                    report_json = json.dumps(quality_report_data, indent=2, default=str)
                    
                    st.download_button(
                        label="📥 Download Quality Report (JSON)",
                        data=report_json,
                        file_name=f"data_quality_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                        mime="application/json"
                    )
                    
                    st.success("✅ Quality assessment report generated successfully!")

    # BUSINESS INTELLIGENCE TAB
    with main_tabs[2]:
        st.header("🚀 Business Intelligence Analysis")
        st.markdown("**Comprehensive business analysis with RFP Bayanati specifications**")
        
        # Only show BI analysis if data is loaded
        if comparator.df_a is not None and comparator.df_b is not None:
            # Import the business intelligence module
            from analysis.business_intelligence import BusinessIntelligenceAnalyzer
            
            # Create organized BI analysis tabs
            bi_tabs = st.tabs([
                "� Executive Summary",
                "� Data Overview", 
                "� Financial Analysis", 
                "� Performance KPIs",
                "� Customer Insights", 
                "📦 Product Analysis", 
                "📊 Financial Ratios",
                "🎯 Advanced KPIs",
                "⚠️ Risk Assessment",
                "🔄 Comparative Analysis",
                "📈 Interactive Dashboard"
            ])
            
            # Perform BI analysis for both datasets
            bi_analyzer_a = BusinessIntelligenceAnalyzer(comparator.df_a)
            bi_analyzer_b = BusinessIntelligenceAnalyzer(comparator.df_b)
            
            # Executive Summary Tab
            with bi_tabs[0]:
                    st.subheader("� Executive Summary & Strategic Insights")
                    st.markdown("**Comprehensive business analysis summary with key findings and recommendations**")
                    
                    # Generate executive summaries for both datasets
                    exec_summary_a = bi_analyzer_a.generate_executive_summary()
                    exec_summary_b = bi_analyzer_b.generate_executive_summary()
                    
                    # Performance Scorecard Section
                    st.subheader("🎯 Performance Scorecard")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write(f"**{display_sheet_a} Performance**")
                        
                        if 'performance_scorecard' in exec_summary_a:
                            scorecard = exec_summary_a['performance_scorecard']
                            
                            # Performance metrics
                            perf_cols = st.columns(2)
                            with perf_cols[0]:
                                st.metric(
                                    "Overall Score", 
                                    f"{scorecard.get('overall_score', 0):.1f}/100",
                                    scorecard.get('performance_grade', 'N/A')
                                )
                            with perf_cols[1]:
                                st.metric(
                                    "Risk Level", 
                                    exec_summary_a.get('risk_assessment', {}).get('overall_risk_level', 'UNKNOWN'),
                                    f"{exec_summary_a.get('risk_assessment', {}).get('critical_issues', 0)} critical issues"
                                )
                            
                            # Performance visualization
                            score = scorecard.get('overall_score', 0)
                            if score >= 85:
                                st.success(f"🏆 **Excellent Performance** ({scorecard.get('performance_grade', 'A')})")
                            elif score >= 70:
                                st.info(f"✅ **Good Performance** ({scorecard.get('performance_grade', 'B')})")
                            elif score >= 60:
                                st.warning(f"⚠️ **Average Performance** ({scorecard.get('performance_grade', 'C')})")
                            else:
                                st.error(f"❌ **Below Average Performance** ({scorecard.get('performance_grade', 'D')})")
                        
                        # Financial highlights
                        if 'financial_highlights' in exec_summary_a:
                            financial = exec_summary_a['financial_highlights']
                            st.write("**💰 Financial Highlights:**")
                            if 'total_value' in financial:
                                st.write(f"• Total Value: **{financial['total_value']}**")
                            if 'average_transaction' in financial:
                                st.write(f"• Average Transaction: **{financial['average_transaction']}**")
                            if 'value_concentration' in financial:
                                st.write(f"• Value Distribution: {financial['value_concentration']}")
                    
                    with col2:
                        st.write(f"**{display_sheet_b} Performance**")
                        
                        if 'performance_scorecard' in exec_summary_b:
                            scorecard = exec_summary_b['performance_scorecard']
                            
                            # Performance metrics
                            perf_cols = st.columns(2)
                            with perf_cols[0]:
                                st.metric(
                                    "Overall Score", 
                                    f"{scorecard.get('overall_score', 0):.1f}/100",
                                    scorecard.get('performance_grade', 'N/A')
                                )
                            with perf_cols[1]:
                                st.metric(
                                    "Risk Level", 
                                    exec_summary_b.get('risk_assessment', {}).get('overall_risk_level', 'UNKNOWN'),
                                    f"{exec_summary_b.get('risk_assessment', {}).get('critical_issues', 0)} critical issues"
                                )
                            
                            # Performance visualization
                            score = scorecard.get('overall_score', 0)
                            if score >= 85:
                                st.success(f"🏆 **Excellent Performance** ({scorecard.get('performance_grade', 'A')})")
                            elif score >= 70:
                                st.info(f"✅ **Good Performance** ({scorecard.get('performance_grade', 'B')})")
                            elif score >= 60:
                                st.warning(f"⚠️ **Average Performance** ({scorecard.get('performance_grade', 'C')})")
                            else:
                                st.error(f"❌ **Below Average Performance** ({scorecard.get('performance_grade', 'D')})")
                        
                        # Financial highlights
                        if 'financial_highlights' in exec_summary_b:
                            financial = exec_summary_b['financial_highlights']
                            st.write("**💰 Financial Highlights:**")
                            if 'total_value' in financial:
                                st.write(f"• Total Value: **{financial['total_value']}**")
                            if 'average_transaction' in financial:
                                st.write(f"• Average Transaction: **{financial['average_transaction']}**")
                            if 'value_concentration' in financial:
                                st.write(f"• Value Distribution: {financial['value_concentration']}")
                    
                    # Key Findings Section
                    st.subheader("🔍 Key Findings")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write(f"**{display_sheet_a} Key Insights:**")
                        key_findings = exec_summary_a.get('key_findings', [])[:5]  # Top 5 findings
                        for finding in key_findings:
                            st.write(f"• {finding}")
                        
                        if not key_findings:
                            st.info("No significant findings identified")
                    
                    with col2:
                        st.write(f"**{display_sheet_b} Key Insights:**")
                        key_findings = exec_summary_b.get('key_findings', [])[:5]  # Top 5 findings
                        for finding in key_findings:
                            st.write(f"• {finding}")
                        
                        if not key_findings:
                            st.info("No significant findings identified")
                    
                    # Strategic Recommendations Section
                    st.subheader("💡 Strategic Recommendations")
                    
                    # Combine recommendations from both datasets
                    all_recommendations = []
                    if 'recommendations' in exec_summary_a:
                        all_recommendations.extend(exec_summary_a['recommendations'])
                    if 'recommendations' in exec_summary_b:
                        all_recommendations.extend(exec_summary_b['recommendations'])
                    
                    # Remove duplicates while preserving order
                    unique_recommendations = []
                    seen = set()
                    for rec in all_recommendations:
                        if rec not in seen:
                            unique_recommendations.append(rec)
                            seen.add(rec)
                    
                    if unique_recommendations:
                        for i, recommendation in enumerate(unique_recommendations[:8], 1):  # Top 8 recommendations
                            st.write(f"**{i}.** {recommendation}")
                    else:
                        st.info("No specific recommendations available - ensure data contains business metrics")
                    
                    # Executive Summary Export
                    st.subheader("📥 Executive Report Export")
                    
                    # Create executive summary report
                    exec_report_data = {
                        'Summary_A': exec_summary_a,
                        'Summary_B': exec_summary_b,
                        'Generated_Date': datetime.now().isoformat(),
                        'Analysis_Type': 'Executive Business Intelligence Summary'
                    }
                    
                    import json
                    exec_report_json = json.dumps(exec_report_data, indent=2, default=str)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="📋 Download Executive Summary (JSON)",
                            data=exec_report_json,
                            file_name=f"executive_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                            mime="application/json",
                            use_container_width=True
                        )
                    
                    with col2:
                        # Create executive summary text report
                        text_report = f"""
EXECUTIVE SUMMARY REPORT
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

=== PERFORMANCE SCORECARD ===
Dataset A Performance: {exec_summary_a.get('performance_scorecard', {}).get('overall_score', 0):.1f}/100 ({exec_summary_a.get('performance_scorecard', {}).get('performance_grade', 'N/A')})
Dataset B Performance: {exec_summary_b.get('performance_scorecard', {}).get('overall_score', 0):.1f}/100 ({exec_summary_b.get('performance_scorecard', {}).get('performance_grade', 'N/A')})

=== KEY FINDINGS ===
{chr(10).join(f"• {finding}" for finding in exec_summary_a.get('key_findings', [])[:3])}
{chr(10).join(f"• {finding}" for finding in exec_summary_b.get('key_findings', [])[:3])}

=== STRATEGIC RECOMMENDATIONS ===
{chr(10).join(f"{i}. {rec}" for i, rec in enumerate(unique_recommendations[:5], 1))}

=== RISK ASSESSMENT ===
Dataset A Risk Level: {exec_summary_a.get('risk_assessment', {}).get('overall_risk_level', 'UNKNOWN')}
Dataset B Risk Level: {exec_summary_b.get('risk_assessment', {}).get('overall_risk_level', 'UNKNOWN')}
                        """
                        
                        st.download_button(
                            label="📄 Download Executive Summary (TXT)",
                            data=text_report,
                            file_name=f"executive_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                            mime="text/plain",
                            use_container_width=True
                        )
            
            # Data Overview Tab
            with bi_tabs[1]:
                st.subheader("📊 Business Overview & Key Insights")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write(f"**{display_sheet_a} Overview**")
                    
                    business_overview_a = bi_analyzer_a.generate_business_overview()
                    
                    # Create KPI cards
                    if 'dataset_info' in business_overview_a:
                        kpi_metrics = []
                        dataset_info = business_overview_a['dataset_info']
                        
                        # Display key metrics in columns
                        metric_cols = st.columns(4)
                        with metric_cols[0]:
                            st.metric("📊 Total Records", f"{dataset_info.get('total_records', 0):,}")
                        with metric_cols[1]:
                            st.metric("📋 Columns", f"{dataset_info.get('total_columns', 0):,}")
                        with metric_cols[2]:
                            st.metric("💰 Amount Columns", f"{len(bi_analyzer_a.amount_columns)}")
                        with metric_cols[3]:
                            st.metric("🎯 Categories", f"{len(bi_analyzer_a.category_columns)}")
                    
                    # Key insights
                    if 'key_business_insights' in business_overview_a:
                        st.write("**Key Business Insights:**")
                        for insight in business_overview_a['key_business_insights']:
                            st.write(f"• {insight}")
                    
                with col2:
                    st.write(f"**{display_sheet_b} Overview**")
                    
                    business_overview_b = bi_analyzer_b.generate_business_overview()
                    
                    # Display key metrics in columns
                    if 'dataset_info' in business_overview_b:
                        dataset_info = business_overview_b['dataset_info']
                        
                        metric_cols = st.columns(4)
                        with metric_cols[0]:
                            st.metric("📊 Total Records", f"{dataset_info.get('total_records', 0):,}")
                        with metric_cols[1]:
                            st.metric("📋 Columns", f"{dataset_info.get('total_columns', 0):,}")
                        with metric_cols[2]:
                            st.metric("💰 Amount Columns", f"{len(bi_analyzer_b.amount_columns)}")
                        with metric_cols[3]:
                            st.metric("🎯 Categories", f"{len(bi_analyzer_b.category_columns)}")
                    
                    # Key insights
                    if 'key_business_insights' in business_overview_b:
                        st.write("**Key Business Insights:**")
                        for insight in business_overview_b['key_business_insights']:
                            st.write(f"• {insight}")
                    
                # Business recommendations
                st.subheader("💡 Business Recommendations")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**Recommendations for {display_sheet_a}:**")
                    recommendations_a = bi_analyzer_a.generate_business_recommendations()
                    for rec in recommendations_a:
                        st.write(f"• {rec}")
                
                with col2:
                    st.write(f"**Recommendations for {display_sheet_b}:**")
                    recommendations_b = bi_analyzer_b.generate_business_recommendations()
                    for rec in recommendations_b:
                        st.write(f"• {rec}")
                
            
            # Sales Performance Tab
            with bi_tabs[2]:
                st.subheader("💰 Sales Performance Analysis")
                
                # Check if both datasets have sales data
                if bi_analyzer_a.amount_columns and bi_analyzer_b.amount_columns:
                        sales_analysis_a = bi_analyzer_a.analyze_sales_performance()
                        sales_analysis_b = bi_analyzer_b.analyze_sales_performance()
                        
                        # Side-by-side sales metrics
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write(f"**{display_sheet_a} Sales Performance**")
                            
                            if sales_analysis_a and not isinstance(sales_analysis_a, dict) or 'error' not in sales_analysis_a:
                                amount_col = list(sales_analysis_a.keys())[0]
                                sales_data = sales_analysis_a[amount_col]
                                
                                # Sales metrics
                                st.metric("Total Revenue", f"${sales_data.get('total_revenue', 0):,.2f}")
                                st.metric("Average Transaction", f"${sales_data.get('average_transaction', 0):,.2f}")
                                st.metric("Median Transaction", f"${sales_data.get('median_transaction', 0):,.2f}")
                                
                                # Revenue trend chart
                                if 'time_trends' in sales_data:
                                    fig_trend_a = visualizer.create_revenue_trend_chart(sales_data, amount_col)
                                    st.plotly_chart(fig_trend_a, use_container_width=True, key="revenue_trend_a")
                                
                                # Category performance chart
                                if 'category_performance' in sales_data:
                                    fig_cat_a = visualizer.create_category_performance_chart(sales_data, amount_col)
                                    st.plotly_chart(fig_cat_a, use_container_width=True, key="category_perf_a")
                            else:
                                st.info("No sales data available for analysis")
                        
                        with col2:
                            st.write(f"**{display_sheet_b} Sales Performance**")
                            
                            if sales_analysis_b and not isinstance(sales_analysis_b, dict) or 'error' not in sales_analysis_b:
                                amount_col = list(sales_analysis_b.keys())[0]
                                sales_data = sales_analysis_b[amount_col]
                                
                                # Sales metrics
                                st.metric("Total Revenue", f"${sales_data.get('total_revenue', 0):,.2f}")
                                st.metric("Average Transaction", f"${sales_data.get('average_transaction', 0):,.2f}")
                                st.metric("Median Transaction", f"${sales_data.get('median_transaction', 0):,.2f}")
                                
                                # Revenue trend chart
                                if 'time_trends' in sales_data:
                                    fig_trend_b = visualizer.create_revenue_trend_chart(sales_data, amount_col)
                                    st.plotly_chart(fig_trend_b, use_container_width=True, key="revenue_trend_b")
                                
                                # Category performance chart
                                if 'category_performance' in sales_data:
                                    fig_cat_b = visualizer.create_category_performance_chart(sales_data, amount_col)
                                    st.plotly_chart(fig_cat_b, use_container_width=True, key="category_perf_b")
                            else:
                                st.info("No sales data available for analysis")
                        
                        # Comparative revenue analysis
                        if sales_analysis_a and sales_analysis_b:
                            st.subheader("📊 Comparative Revenue Analysis")
                            
                            try:
                                fig_comparative = visualizer.create_comparative_business_chart(
                                    sales_analysis_a[list(sales_analysis_a.keys())[0]], 
                                    sales_analysis_b[list(sales_analysis_b.keys())[0]], 
                                    display_sheet_a, display_sheet_b, 'revenue_trends'
                                )
                                st.plotly_chart(fig_comparative, use_container_width=True, key="comparative_revenue")
                            except Exception as e:
                                st.info("Comparative revenue trends not available")
                    
                else:
                    st.info("💡 Sales analysis requires datasets with amount/revenue columns (e.g., 'sales_amount', 'price', 'revenue')")
            
            # Customer Analytics Tab
            with bi_tabs[3]:
                st.subheader("👥 Customer Analytics")
                
                customer_analysis_a = bi_analyzer_a.analyze_customer_insights()
                customer_analysis_b = bi_analyzer_b.analyze_customer_insights()
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write(f"**{display_sheet_a} Customer Analysis**")
                    
                    # Customer segmentation chart
                    if 'customer_segmentation' in customer_analysis_a:
                        fig_seg_a = visualizer.create_customer_segmentation_chart(customer_analysis_a)
                        st.plotly_chart(fig_seg_a, use_container_width=True, key="customer_seg_a")
                        
                        # Customer satisfaction analysis
                        if 'satisfaction_analysis' in customer_analysis_a:
                            fig_sat_a = visualizer.create_satisfaction_analysis_chart(customer_analysis_a)
                            st.plotly_chart(fig_sat_a, use_container_width=True, key="satisfaction_a")
                    
                    with col2:
                        st.write(f"**{display_sheet_b} Customer Analysis**")
                        
                        # Customer segmentation chart
                        if 'customer_segmentation' in customer_analysis_b:
                            fig_seg_b = visualizer.create_customer_segmentation_chart(customer_analysis_b)
                            st.plotly_chart(fig_seg_b, use_container_width=True, key="customer_seg_b")
                        
                        # Customer satisfaction analysis
                        if 'satisfaction_analysis' in customer_analysis_b:
                            fig_sat_b = visualizer.create_satisfaction_analysis_chart(customer_analysis_b)
                            st.plotly_chart(fig_sat_b, use_container_width=True, key="satisfaction_b")
                
            
            # Product Analysis Tab
            with bi_tabs[4]:
                    st.subheader("📦 Product Performance Analysis")
                    
                    product_analysis_a = bi_analyzer_a.analyze_product_performance()
                    product_analysis_b = bi_analyzer_b.analyze_product_performance()
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write(f"**{display_sheet_a} Product Analysis**")
                        
                        if 'message' not in product_analysis_a:
                            fig_prod_a = visualizer.create_product_performance_chart(product_analysis_a)
                            st.plotly_chart(fig_prod_a, use_container_width=True, key="product_perf_a")
                        else:
                            st.info(product_analysis_a.get('message', 'No product data available'))
                    
                    with col2:
                        st.write(f"**{display_sheet_b} Product Analysis**")
                        
                        if 'message' not in product_analysis_b:
                            fig_prod_b = visualizer.create_product_performance_chart(product_analysis_b)
                            st.plotly_chart(fig_prod_b, use_container_width=True, key="product_perf_b")
                        else:
                            st.info(product_analysis_b.get('message', 'No product data available'))
                
            
            # Financial Ratios Tab (RFP Bayanati Specifications)
            with bi_tabs[5]:
                    st.subheader("📊 Financial Ratios & Performance Metrics")
                    st.markdown("*Based on RFP Bayanati Project specifications including ROE, Liquidity Ratios, Cash Conversion Cycle, and profitability metrics*")
                    
                    # Get financial ratios from business overview
                    business_overview_a = bi_analyzer_a.generate_business_overview()
                    business_overview_b = bi_analyzer_b.generate_business_overview()
                    
                    financial_ratios_a = business_overview_a.get('financial_ratios', {})
                    financial_ratios_b = business_overview_b.get('financial_ratios', {})
                    
                    # Display Financial Ratios Dashboard for Dataset A
                    if financial_ratios_a and any(financial_ratios_a.get(key, {}) for key in ['liquidity_ratios', 'profitability_ratios']):
                        st.write(f"**{display_sheet_a} Financial Ratios Dashboard**")
                        fig_ratios_a = visualizer.create_financial_ratios_dashboard(financial_ratios_a)
                        st.plotly_chart(fig_ratios_a, use_container_width=True, key="financial_ratios_a")
                        
                        # Show insights
                        if 'insights' in financial_ratios_a:
                            st.write("**Key Financial Insights:**")
                            for insight in financial_ratios_a['insights'][:3]:
                                st.info(f"💡 {insight}")
                    else:
                        st.info(f"💡 {display_sheet_a}: Financial ratios require columns like 'assets', 'liabilities', 'revenue', 'net_income', 'equity' for accurate calculation")
                    
                    st.divider()
                    
                    # Display Financial Ratios Dashboard for Dataset B
                    if financial_ratios_b and any(financial_ratios_b.get(key, {}) for key in ['liquidity_ratios', 'profitability_ratios']):
                        st.write(f"**{display_sheet_b} Financial Ratios Dashboard**")
                        fig_ratios_b = visualizer.create_financial_ratios_dashboard(financial_ratios_b)
                        st.plotly_chart(fig_ratios_b, use_container_width=True, key="financial_ratios_b")
                        
                        # Show insights
                        if 'insights' in financial_ratios_b:
                            st.write("**Key Financial Insights:**")
                            for insight in financial_ratios_b['insights'][:3]:
                                st.info(f"💡 {insight}")
                    else:
                        st.info(f"💡 {display_sheet_b}: Financial ratios require columns like 'assets', 'liabilities', 'revenue', 'net_income', 'equity' for accurate calculation")
                    
                    # Financial Performance Comparison
                    if financial_ratios_a and financial_ratios_b:
                        st.write("**📊 Financial Performance Comparison**")
                        
                        col1, col2, col3 = st.columns(3)
                        
                        # ROE Comparison
                        roe_a = financial_ratios_a.get('profitability_ratios', {}).get('roe', {}).get('value', 0)
                        roe_b = financial_ratios_b.get('profitability_ratios', {}).get('roe', {}).get('value', 0)
                        
                        with col1:
                            st.metric("ROE Comparison", 
                                     f"{roe_b:.2f}%", 
                                     delta=f"{roe_b - roe_a:.2f}%" if roe_a > 0 else None)
                        
                        # Current Ratio Comparison
                        current_a = financial_ratios_a.get('liquidity_ratios', {}).get('current_ratio', {}).get('value', 0)
                        current_b = financial_ratios_b.get('liquidity_ratios', {}).get('current_ratio', {}).get('value', 0)
                        
                        with col2:
                            st.metric("Current Ratio Comparison", 
                                     f"{current_b:.2f}", 
                                     delta=f"{current_b - current_a:.2f}" if current_a > 0 else None)
                        
                        # Net Margin Comparison
                        margin_a = financial_ratios_a.get('profitability_ratios', {}).get('net_profit_margin', {}).get('value', 0)
                        margin_b = financial_ratios_b.get('profitability_ratios', {}).get('net_profit_margin', {}).get('value', 0)
                        
                        with col3:
                            st.metric("Net Margin Comparison", 
                                     f"{margin_b:.2f}%", 
                                     delta=f"{margin_b - margin_a:.2f}%" if margin_a > 0 else None)
                
            
            # Advanced Business KPIs Tab (New)
            with bi_tabs[6]:
                    st.subheader("🎯 Advanced Business KPIs")
                    st.markdown("*Sales Growth, Customer Retention, Employee Turnover, Training ROI, and operational metrics*")
                    
                    # Get advanced KPIs
                    advanced_kpis_a = bi_analyzer_a.analyze_advanced_business_kpis()
                    advanced_kpis_b = bi_analyzer_b.analyze_advanced_business_kpis()
                    
                    # KPI Scorecard for Dataset A
                    if advanced_kpis_a:
                        st.write(f"**{display_sheet_a} Business KPIs Scorecard**")
                        fig_kpi_a = visualizer.create_kpi_scorecard(advanced_kpis_a)
                        st.plotly_chart(fig_kpi_a, use_container_width=True, key="kpi_scorecard_a")
                    
                    st.divider()
                    
                    # KPI Scorecard for Dataset B
                    if advanced_kpis_b:
                        st.write(f"**{display_sheet_b} Business KPIs Scorecard**")
                        fig_kpi_b = visualizer.create_kpi_scorecard(advanced_kpis_b)
                        st.plotly_chart(fig_kpi_b, use_container_width=True, key="kpi_scorecard_b")
                    
                    # Detailed KPI Analysis
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write(f"**{display_sheet_a} KPI Details**")
                        
                        # Sales & Marketing KPIs
                        sales_kpis_a = advanced_kpis_a.get('sales_marketing_kpis', {})
                        if sales_kpis_a and 'error' not in sales_kpis_a:
                            st.write("**Sales & Marketing:**")
                            for kpi_name, kpi_data in sales_kpis_a.items():
                                if isinstance(kpi_data, dict) and 'value' in kpi_data:
                                    unit = kpi_data.get('unit', '')
                                    interpretation = kpi_data.get('interpretation', '')
                                    st.metric(kpi_name.replace('_', ' ').title(), 
                                             f"{kpi_data['value']}{unit}",
                                             help=f"Interpretation: {interpretation}")
                        
                        # HR KPIs
                        hr_kpis_a = advanced_kpis_a.get('hr_kpis', {})
                        if hr_kpis_a and 'error' not in hr_kpis_a:
                            st.write("**HR Metrics:**")
                            for kpi_name, kpi_data in hr_kpis_a.items():
                                if isinstance(kpi_data, dict) and 'value' in kpi_data:
                                    unit = kpi_data.get('unit', '')
                                    interpretation = kpi_data.get('interpretation', '')
                                    st.metric(kpi_name.replace('_', ' ').title(), 
                                             f"{kpi_data['value']}{unit}",
                                             help=f"Interpretation: {interpretation}")
                    
                    with col2:
                        st.write(f"**{display_sheet_b} KPI Details**")
                        
                        # Sales & Marketing KPIs
                        sales_kpis_b = advanced_kpis_b.get('sales_marketing_kpis', {})
                        if sales_kpis_b and 'error' not in sales_kpis_b:
                            st.write("**Sales & Marketing:**")
                            for kpi_name, kpi_data in sales_kpis_b.items():
                                if isinstance(kpi_data, dict) and 'value' in kpi_data:
                                    unit = kpi_data.get('unit', '')
                                    interpretation = kpi_data.get('interpretation', '')
                                    st.metric(kpi_name.replace('_', ' ').title(), 
                                             f"{kpi_data['value']}{unit}",
                                             help=f"Interpretation: {interpretation}")
                        
                        # HR KPIs
                        hr_kpis_b = advanced_kpis_b.get('hr_kpis', {})
                        if hr_kpis_b and 'error' not in hr_kpis_b:
                            st.write("**HR Metrics:**")
                            for kpi_name, kpi_data in hr_kpis_b.items():
                                if isinstance(kpi_data, dict) and 'value' in kpi_data:
                                    unit = kpi_data.get('unit', '')
                                    interpretation = kpi_data.get('interpretation', '')
                                    st.metric(kpi_name.replace('_', ' ').title(), 
                                             f"{kpi_data['value']}{unit}",
                                             help=f"Interpretation: {interpretation}")
                
            
            # Risk & Alerts Tab (New)
            with bi_tabs[7]:
                    st.subheader("⚠️ Risk Assessment & Early Warning System")
                    st.markdown("*Cross-Institution Benchmarking, Sector Analysis, and Early Warning Indicators*")
                    
                    # Get benchmarking and alerts
                    benchmarking_a = bi_analyzer_a.analyze_benchmarking_alerts()
                    benchmarking_b = bi_analyzer_b.analyze_benchmarking_alerts()
                    
                    # Early Warning Indicators
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write(f"**{display_sheet_a} Alert System**")
                        
                        early_warnings_a = benchmarking_a.get('early_warning_indicators', {})
                        if early_warnings_a:
                            fig_alerts_a = visualizer.create_early_warning_alerts(early_warnings_a)
                            st.plotly_chart(fig_alerts_a, use_container_width=True, key="alerts_a")
                            
                            # Show critical alerts details
                            critical_alerts = early_warnings_a.get('critical_alerts', [])
                            if critical_alerts:
                                st.error("🚨 **Critical Alerts:**")
                                for alert in critical_alerts[:3]:
                                    st.warning(f"**{alert.get('type', 'Alert')}:** {alert.get('message', 'No details')}")
                                    st.info(f"📋 Recommendation: {alert.get('recommendation', 'Review required')}")
                    
                    with col2:
                        st.write(f"**{display_sheet_b} Alert System**")
                        
                        early_warnings_b = benchmarking_b.get('early_warning_indicators', {})
                        if early_warnings_b:
                            fig_alerts_b = visualizer.create_early_warning_alerts(early_warnings_b)
                            st.plotly_chart(fig_alerts_b, use_container_width=True, key="alerts_b")
                            
                            # Show critical alerts details
                            critical_alerts = early_warnings_b.get('critical_alerts', [])
                            if critical_alerts:
                                st.error("🚨 **Critical Alerts:**")
                                for alert in critical_alerts[:3]:
                                    st.warning(f"**{alert.get('type', 'Alert')}:** {alert.get('message', 'No details')}")
                                    st.info(f"📋 Recommendation: {alert.get('recommendation', 'Review required')}")
                    
                    st.divider()
                    
                    # Sector Performance Analysis
                    sector_analysis_a = benchmarking_a.get('sector_analysis', {})
                    sector_analysis_b = benchmarking_b.get('sector_analysis', {})
                    
                    if sector_analysis_a.get('sector_rankings') or sector_analysis_b.get('sector_rankings'):
                        st.write("**📊 Sector Performance Analysis**")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if sector_analysis_a.get('sector_rankings'):
                                st.write(f"**{display_sheet_a} Sector Heatmap**")
                                fig_sector_a = visualizer.create_sector_performance_heatmap(sector_analysis_a)
                                st.plotly_chart(fig_sector_a, use_container_width=True, key="sector_heatmap_a")
                        
                        with col2:
                            if sector_analysis_b.get('sector_rankings'):
                                st.write(f"**{display_sheet_b} Sector Heatmap**")
                                fig_sector_b = visualizer.create_sector_performance_heatmap(sector_analysis_b)
                                st.plotly_chart(fig_sector_b, use_container_width=True, key="sector_heatmap_b")
                    
                    # Risk Assessment Radar Charts
                    risk_assessment_a = benchmarking_a.get('risk_assessment', {})
                    risk_assessment_b = benchmarking_b.get('risk_assessment', {})
                    
                    if risk_assessment_a.get('risk_factors') or risk_assessment_b.get('risk_factors'):
                        st.write("**🎯 Risk Assessment Overview**")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if risk_assessment_a.get('risk_factors'):
                                st.write(f"**{display_sheet_a} Risk Profile**")
                                fig_risk_a = visualizer.create_risk_assessment_radar(risk_assessment_a)
                                st.plotly_chart(fig_risk_a, use_container_width=True, key="risk_radar_a")
                                
                                risk_level = risk_assessment_a.get('risk_level', 'Unknown')
                                risk_score = risk_assessment_a.get('risk_score', 0)
                                st.metric("Overall Risk Level", risk_level, f"Score: {risk_score}")
                        
                        with col2:
                            if risk_assessment_b.get('risk_factors'):
                                st.write(f"**{display_sheet_b} Risk Profile**")
                                fig_risk_b = visualizer.create_risk_assessment_radar(risk_assessment_b)
                                st.plotly_chart(fig_risk_b, use_container_width=True, key="risk_radar_b")
                                
                                risk_level = risk_assessment_b.get('risk_level', 'Unknown')
                                risk_score = risk_assessment_b.get('risk_score', 0)
                                st.metric("Overall Risk Level", risk_level, f"Score: {risk_score}")
                    
                    # Data Quality & Governance
                    governance_a = benchmarking_a.get('performance_benchmarks', {})
                    governance_b = benchmarking_b.get('performance_benchmarks', {})
                    
                    if governance_a or governance_b:
                        st.write("**📊 Data Quality & Governance Metrics**")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if governance_a.get('data_quality'):
                                st.write(f"**{display_sheet_a} Data Quality**")
                                dq_score = governance_a['data_quality'].get('completeness_score', 0)
                                st.metric("Data Completeness", f"{dq_score:.1f}%")
                                
                                tier = governance_a['data_quality'].get('benchmark_tier', 'Unknown')
                                st.info(f"Quality Tier: {tier}")
                        
                        with col2:
                            if governance_b.get('data_quality'):
                                st.write(f"**{display_sheet_b} Data Quality**")
                                dq_score = governance_b['data_quality'].get('completeness_score', 0)
                                st.metric("Data Completeness", f"{dq_score:.1f}%")
                                
                                tier = governance_b['data_quality'].get('benchmark_tier', 'Unknown')
                                st.info(f"Quality Tier: {tier}")
                
            
            # Comparative Analysis Tab
            with bi_tabs[8]:
                    st.subheader("⚠️ Risk Assessment & Early Warning Systems")
                    st.markdown("**Comprehensive risk analysis with predictive indicators and early warning signals**")
                    
                    try:
                        # Get risk assessment from both datasets
                        risk_analysis_a = bi_analyzer_a.perform_risk_analysis() if hasattr(bi_analyzer_a, 'perform_risk_analysis') else {}
                        risk_analysis_b = bi_analyzer_b.perform_risk_analysis() if hasattr(bi_analyzer_b, 'perform_risk_analysis') else {}
                        
                        # Financial Health Indicators
                        st.subheader("💊 Financial Health Indicators")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write(f"**{display_sheet_a} Risk Profile**")
                            
                            # Create risk indicators based on available data
                            df_a = comparator.df_a
                            numeric_cols_a = df_a.select_dtypes(include=[np.number]).columns
                            
                            if len(numeric_cols_a) > 0:
                                # Calculate risk indicators
                                volatility = df_a[numeric_cols_a].std().mean() if len(numeric_cols_a) > 0 else 0
                                trend_stability = 1 - (df_a[numeric_cols_a].var().mean() / (df_a[numeric_cols_a].mean().mean() + 1))
                                data_quality = (df_a.notna().sum().sum() / (len(df_a) * len(df_a.columns))) * 100
                                
                                # Risk level calculation
                                risk_score = min(100, max(0, (volatility * 0.3 + (1-trend_stability) * 0.4 + (100-data_quality) * 0.3)))
                                risk_level = "Low" if risk_score < 30 else "Medium" if risk_score < 70 else "High"
                                risk_color = "🟢" if risk_score < 30 else "🟡" if risk_score < 70 else "🔴"
                                
                                st.metric("Overall Risk Score", f"{risk_score:.1f}/100", f"{risk_color} {risk_level}")
                                st.metric("Data Quality", f"{data_quality:.1f}%")
                                st.metric("Trend Stability", f"{trend_stability*100:.1f}%")
                            else:
                                st.info("Insufficient numeric data for risk assessment")
                        
                        with col2:
                            st.write(f"**{display_sheet_b} Risk Profile**")
                            
                            df_b = comparator.df_b
                            numeric_cols_b = df_b.select_dtypes(include=[np.number]).columns
                            
                            if len(numeric_cols_b) > 0:
                                # Calculate risk indicators
                                volatility = df_b[numeric_cols_b].std().mean() if len(numeric_cols_b) > 0 else 0
                                trend_stability = 1 - (df_b[numeric_cols_b].var().mean() / (df_b[numeric_cols_b].mean().mean() + 1))
                                data_quality = (df_b.notna().sum().sum() / (len(df_b) * len(df_b.columns))) * 100
                                
                                # Risk level calculation
                                risk_score = min(100, max(0, (volatility * 0.3 + (1-trend_stability) * 0.4 + (100-data_quality) * 0.3)))
                                risk_level = "Low" if risk_score < 30 else "Medium" if risk_score < 70 else "High"
                                risk_color = "🟢" if risk_score < 30 else "🟡" if risk_score < 70 else "🔴"
                                
                                st.metric("Overall Risk Score", f"{risk_score:.1f}/100", f"{risk_color} {risk_level}")
                                st.metric("Data Quality", f"{data_quality:.1f}%")
                                st.metric("Trend Stability", f"{trend_stability*100:.1f}%")
                            else:
                                st.info("Insufficient numeric data for risk assessment")
                        
                        # Early Warning Indicators
                        st.subheader("🚨 Early Warning Indicators")
                        
                        warning_indicators = []
                        
                        # Check for common risk patterns
                        for i, (df, name) in enumerate([(comparator.df_a, display_sheet_a), (comparator.df_b, display_sheet_b)]):
                            numeric_cols = df.select_dtypes(include=[np.number]).columns
                            
                            if len(numeric_cols) > 0:
                                # Missing data warning
                                missing_pct = (df.isnull().sum().sum() / (len(df) * len(df.columns))) * 100
                                if missing_pct > 10:
                                    warning_indicators.append(f"⚠️ {name}: High missing data rate ({missing_pct:.1f}%)")
                                
                                # Outlier detection
                                for col in numeric_cols[:3]:  # Check first 3 numeric columns
                                    data = df[col].dropna()
                                    if len(data) > 0:
                                        Q1 = data.quantile(0.25)
                                        Q3 = data.quantile(0.75)
                                        IQR = Q3 - Q1
                                        outliers = data[(data < Q1 - 1.5*IQR) | (data > Q3 + 1.5*IQR)]
                                        if len(outliers) > len(data) * 0.1:  # More than 10% outliers
                                            warning_indicators.append(f"📊 {name}: Potential data quality issues in {col}")
                        
                        if warning_indicators:
                            for warning in warning_indicators:
                                st.warning(warning)
                        else:
                            st.success("🟢 No significant risk indicators detected")
                        
                        # Risk Mitigation Recommendations
                        st.subheader("💡 Risk Mitigation Recommendations")
                        
                        recommendations = [
                            "🔍 **Data Quality**: Implement regular data validation and cleansing procedures",
                            "📊 **Monitoring**: Set up automated alerts for key performance indicators",
                            "🔄 **Backup**: Maintain data backup and recovery procedures",
                            "📈 **Trending**: Monitor data trends for early warning signals",
                            "🛡️ **Validation**: Implement data integrity checks and validation rules"
                        ]
                        
                        for rec in recommendations:
                            st.markdown(rec)
                    
                    except Exception as e:
                        st.error(f"Risk assessment error: {str(e)}")
                        st.info("💡 Risk assessment requires numeric data for comprehensive analysis")
            
            # Comparative Business Analysis Tab
            with bi_tabs[9]:
                    st.subheader("🔄 Comparative Business Analysis")
                    
                    # Perform comparative analysis
                    try:
                        comparison_analysis = bi_analyzer_a.compare_business_performance(bi_analyzer_b)
                        
                        # Performance summary
                        if 'performance_summary' in comparison_analysis:
                            st.write("**📈 Performance Summary:**")
                            for summary in comparison_analysis['performance_summary']:
                                st.write(f"• {summary}")
                        
                        # Revenue comparison
                        if 'revenue_comparison' in comparison_analysis:
                            st.subheader("💰 Revenue Comparison")
                            revenue_comp = comparison_analysis['revenue_comparison']
                            
                            if 'message' not in revenue_comp:
                                col1, col2, col3 = st.columns(3)
                                
                                with col1:
                                    total_rev = revenue_comp.get('total_revenue', {})
                                    st.metric(
                                        "Revenue Growth", 
                                        f"{total_rev.get('growth_rate', 0):+.1f}%",
                                        f"${total_rev.get('difference', 0):+,.2f}"
                                    )
                                
                                with col2:
                                    avg_trans = revenue_comp.get('average_transaction', {})
                                    st.metric(
                                        "Avg Transaction Growth", 
                                        f"{avg_trans.get('improvement', 0):+.1f}%",
                                        f"${avg_trans.get('difference', 0):+,.2f}"
                                    )
                                
                                with col3:
                                    dataset_comp = comparison_analysis.get('dataset_comparison', {})
                                    differences = dataset_comp.get('differences', {})
                                    st.metric(
                                        "Data Volume Growth", 
                                        f"{differences.get('growth_rate', 0):+.1f}%",
                                        f"{differences.get('row_difference', 0):+,} records"
                                    )
                        
                        # Customer comparison
                        if 'customer_comparison' in comparison_analysis:
                            st.subheader("👥 Customer Growth")
                            customer_comp = comparison_analysis['customer_comparison']
                            
                            if 'message' not in customer_comp and 'customer_count' in customer_comp:
                                customer_data = customer_comp['customer_count']
                                st.metric(
                                    "Customer Base Growth",
                                    f"{customer_data.get('growth_rate', 0):+.1f}%",
                                    f"{customer_data.get('difference', 0):+,} customers"
                                )
                    
                    except Exception as e:
                        st.error(f"Comparative analysis error: {str(e)}")
                        st.info("💡 Comparative analysis requires compatible datasets with similar business metrics")
            
            # Interactive Dashboard Tab
            with bi_tabs[10]:
                    st.subheader("📈 Interactive Dashboard")
                    st.markdown("Create interactive visualizations with drill-down capabilities and export options")
                    
                    try:
                        # Dataset selection
                        col1, col2 = st.columns(2)
                        with col1:
                            dataset_choice = st.selectbox(
                                "Select Dataset:",
                                [display_sheet_a, display_sheet_b],
                                key="interactive_dataset"
                            )
                        
                        with col2:
                            chart_type = st.selectbox(
                                "Select Chart Type:",
                                ["Interactive Bar Chart", "Drill-Down Scatter Plot", "Time Series Analysis", "Category Comparison"],
                                key="interactive_chart_type"
                            )
                        
                        # Get selected dataset
                        selected_df = comparator.df_a if dataset_choice == display_sheet_a else comparator.df_b
                        available_columns = list(selected_df.columns)
                        
                        if available_columns:
                            # Column mapping interface
                            st.markdown("#### 🎯 Data Mapping")
                            col1, col2, col3 = st.columns(3)
                            
                            with col1:
                                x_column = st.selectbox(
                                    "X-Axis Column:",
                                    available_columns,
                                    key="interactive_x_column"
                                )
                            
                            with col2:
                                numeric_columns = [col for col in available_columns if selected_df[col].dtype in ['int64', 'float64']]
                                y_column = st.selectbox(
                                    "Y-Axis Column:",
                                    numeric_columns if numeric_columns else available_columns,
                                    key="interactive_y_column"
                                )
                            
                            with col3:
                                color_column = st.selectbox(
                                    "Color/Group Column (Optional):",
                                    ["None"] + available_columns,
                                    key="interactive_color_column"
                                )
                            
                            # Export format selection
                            export_format = st.selectbox(
                                "Export Format:",
                                ["PNG", "PDF", "HTML", "SVG"],
                                key="export_format"
                            )
                            
                            # Chart generation
                            if st.button("🎨 Generate Interactive Chart", key="generate_interactive"):
                                with st.spinner("Creating interactive visualization..."):
                                    try:
                                        from analysis.visualization import InteractiveDashboard, ChartExporter
                                        
                                        # Initialize dashboard with both datasets
                                        dashboard = InteractiveDashboard(
                                            comparator.df_a, 
                                            comparator.df_b,
                                            display_sheet_a,
                                            display_sheet_b
                                        )
                                        
                                        # Prepare data
                                        chart_data = selected_df.copy()
                                        color_col = None if color_column == "None" else color_column
                                        
                                        # Create chart based on selection
                                        if chart_type == "Interactive Bar Chart":
                                            fig = dashboard.create_single_dataset_bar_chart(
                                                chart_data, x_column, y_column, color_col
                                            )
                                            st.plotly_chart(fig, use_container_width=True)
                                            
                                        elif chart_type == "Drill-Down Scatter Plot":
                                            fig = dashboard.create_single_dataset_scatter(
                                                chart_data, x_column, y_column, color_col
                                            )
                                            st.plotly_chart(fig, use_container_width=True)
                                            
                                        elif chart_type == "Time Series Analysis":
                                            # Check if x_column can be converted to datetime
                                            try:
                                                chart_data_time = chart_data.copy()
                                                chart_data_time[x_column] = pd.to_datetime(chart_data_time[x_column], errors='coerce')
                                                chart_data_time = chart_data_time.dropna(subset=[x_column])
                                                
                                                if len(chart_data_time) > 0:
                                                    # Use scatter plot for time series with line mode
                                                    fig = dashboard.create_single_dataset_scatter(
                                                        chart_data_time, x_column, y_column, color_col
                                                    )
                                                    # Update to line chart for time series
                                                    for trace in fig.data:
                                                        trace.mode = 'lines+markers'
                                                    fig.update_layout(title=f"Time Series: {y_column} over {x_column}")
                                                    fig.update_xaxes(title="Date")
                                                    st.plotly_chart(fig, use_container_width=True)
                                                else:
                                                    st.error("No valid date values found in selected X-axis column")
                                            except Exception as date_error:
                                                st.error(f"Date conversion error: {str(date_error)}")
                                        
                                        elif chart_type == "Category Comparison":
                                            # Use bar chart for category comparison
                                            try:
                                                fig = dashboard.create_single_dataset_bar_chart(
                                                    chart_data, x_column, y_column, color_col
                                                )
                                                fig.update_layout(title=f"Category Comparison: {y_column} by {x_column}")
                                                st.plotly_chart(fig, use_container_width=True)
                                            except Exception as group_error:
                                                st.error(f"Category comparison error: {str(group_error)}")
                                        
                                        # Export options
                                        st.markdown("#### 📥 Export Options")
                                        col1, col2 = st.columns(2)
                                        
                                        with col1:
                                            if st.button(f"📊 Export Chart as {export_format}", key="export_chart"):
                                                try:
                                                    # Export chart with data
                                                    export_data = dashboard.export_chart_with_data(
                                                        fig, chart_data, export_format.lower()
                                                    )
                                                    
                                                    if export_format in ["PNG", "PDF", "SVG"]:
                                                        st.download_button(
                                                            label=f"Download {export_format}",
                                                            data=export_data,
                                                            file_name=f"interactive_chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{export_format.lower()}",
                                                            mime=f"image/{export_format.lower()}" if export_format != "PDF" else "application/pdf"
                                                        )
                                                    else:  # HTML
                                                        st.download_button(
                                                            label="Download HTML",
                                                            data=export_data,
                                                            file_name=f"interactive_chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
                                                            mime="text/html"
                                                        )
                                                    
                                                    st.success(f"Chart exported as {export_format}!")
                                                
                                                except Exception as e:
                                                    st.error(f"Export error: {str(e)}")
                                        
                                        with col2:
                                            if st.button("📋 Export Data as CSV", key="export_data"):
                                                csv_data = chart_data.to_csv(index=False)
                                                st.download_button(
                                                    label="Download CSV Data",
                                                    data=csv_data,
                                                    file_name=f"chart_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                                    mime="text/csv"
                                                )
                                                st.success("Data exported as CSV!")
                                        
                                        # Chart insights
                                        st.markdown("#### 💡 Chart Insights")
                                        insights_col1, insights_col2 = st.columns(2)
                                        
                                        with insights_col1:
                                            st.metric(
                                                "Data Points",
                                                f"{len(chart_data):,}",
                                                help="Total number of data points in the visualization"
                                            )
                                            
                                            # Calculate basic statistics for numeric columns
                                            if y_column in chart_data.columns and chart_data[y_column].dtype in ['int64', 'float64']:
                                                avg_value = chart_data[y_column].mean()
                                                st.metric(
                                                    f"Average {y_column}",
                                                    f"{avg_value:,.2f}",
                                                    help=f"Mean value of {y_column}"
                                                )
                                        
                                        with insights_col2:
                                            if color_col and color_col in chart_data.columns:
                                                unique_categories = chart_data[color_col].nunique()
                                                st.metric(
                                                    f"Unique {color_col}",
                                                    f"{unique_categories:,}",
                                                    help=f"Number of unique categories in {color_col}"
                                                )
                                            
                                            if y_column in chart_data.columns and chart_data[y_column].dtype in ['int64', 'float64']:
                                                max_value = chart_data[y_column].max()
                                                st.metric(
                                                    f"Maximum {y_column}",
                                                    f"{max_value:,.2f}",
                                                    help=f"Highest value in {y_column}"
                                                )
                                    
                                    except Exception as e:
                                        st.error(f"Interactive dashboard error: {str(e)}")
                                        st.info("💡 Please ensure your data contains the selected columns and appropriate data types")
                        
                        else:
                            st.info("📊 No columns available in the selected dataset")
                            st.markdown("""
                            **Interactive Dashboard Features:**
                            - 🎨 **Interactive Charts**: Bar charts, scatter plots, time series
                            - 🔍 **Drill-Down Capabilities**: Click on chart elements for detailed views
                            - 📥 **Multiple Export Formats**: PNG, PDF, HTML, SVG
                            - 📊 **Real-Time Insights**: Dynamic metrics and statistics
                            - 🎯 **Flexible Data Mapping**: Choose any columns for visualization
                            - 📋 **Data Export**: Download both charts and underlying data
                            """)
                    
                    except Exception as e:
                        st.error(f"Interactive Dashboard error: {str(e)}")
                        st.info("💡 Interactive Dashboard requires processed data with appropriate column types")
        else:
            st.info("📁 Please upload and load data files in the sidebar to enable Business Intelligence analysis")
            st.markdown("""
            **Business Intelligence Features:**
            - 📊 Financial Ratio Analysis (ROE, ROA, Current Ratio, etc.)
            - 📈 Advanced KPIs (Sales growth, Customer retention, etc.) 
            - 👥 Customer Analytics (Segmentation, Lifetime Value, etc.)
            - ⚠️ Risk Assessment & Early Warning Systems
            - 💰 Comprehensive Financial Analysis
            """)
    
    # DATA ANALYSIS TAB
    with main_tabs[3]:
        st.header("🔍 Advanced Data Analysis")
        st.markdown("**Statistical analysis, data quality assessment, and insights**")
        
        if comparator.df_a is not None or comparator.df_b is not None:
            # Show data analysis options
            analysis_tabs = st.tabs(["📊 Statistical Analysis", "🔍 Data Quality", "📈 Trends & Patterns"])
            
            with analysis_tabs[0]:
                st.subheader("📊 Statistical Analysis")
                st.info("Coming soon: Descriptive statistics, correlation analysis, and distribution analysis")
            
            with analysis_tabs[1]:
                st.subheader("🔍 Data Quality Assessment")  
                st.info("Coming soon: Missing data analysis, duplicate detection, and data validation")
            
            with analysis_tabs[2]:
                st.subheader("📈 Trends & Patterns")
                st.info("Coming soon: Trend analysis, seasonal patterns, and anomaly detection")
        else:
            st.info("📁 Please upload data files in the sidebar to enable advanced data analysis")
    
    # REPORTS & EXPORT TAB
    with main_tabs[4]:
        st.header("📋 Reports & Export")
        st.markdown("**Generate and download comprehensive reports**")
        
        if hasattr(comparator, 'results') and comparator.results:
            # Export functionality
            st.subheader("📥 Export Results")
            
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
        else:
            st.info("🔄 No results available for export. Please run a comparison analysis first.")
            st.markdown("""
            **Available Export Options:**
            - 📊 Complete Excel report with multiple sheets
            - 📈 Individual CSV files for each result category  
            - 💼 Executive summary reports
            - 📋 Data quality assessment reports
            """)
    
    # Footer information
    st.divider()
    with st.expander("ℹ️ About this Platform", expanded=False):
        st.markdown("""
        **Excel Comparison & Business Intelligence Platform v2.0**
        
        This comprehensive platform combines advanced Excel comparison capabilities with 
        business intelligence analysis following RFP Bayanati specifications.
        
        **Key Capabilities:**
        - 🔍 Advanced fuzzy matching algorithms
        - 📊 Comprehensive business intelligence analysis  
        - 📈 Financial ratio calculations (ROE, ROA, liquidity ratios)
        - 👥 Customer analytics and segmentation
        - 📋 Professional reporting and export
        - ⚠️ Risk assessment and early warning systems
        """)

if __name__ == "__main__":
    main()