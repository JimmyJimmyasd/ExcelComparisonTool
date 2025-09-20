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
        """Enhanced column analysis with actionable insights"""
        
        st.subheader(f"📊 Data Quality Analysis: {column_name}")
        
        # Basic statistics in metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Rows", f"{len(df):,}")
        with col2:  
            unique_count = df[column_name].nunique()
            st.metric("Unique Values", f"{unique_count:,}")
        with col3:
            null_count = df[column_name].isnull().sum()
            null_pct = (null_count / len(df)) * 100
            st.metric("Missing Values", f"{null_count:,} ({null_pct:.1f}%)")
        with col4:
            duplicate_count = len(df) - unique_count
            st.metric("Duplicates", f"{duplicate_count:,}")
        
        # Data quality alerts
        alerts = []
        if null_pct > 10:
            alerts.append(f"⚠️ High missing data rate ({null_pct:.1f}%). Consider data cleaning.")
        
        if duplicate_count > len(df) * 0.5:
            alerts.append(f"⚠️ High duplicate rate ({(duplicate_count/len(df)*100):.1f}%). May affect matching accuracy.")
        
        if unique_count == len(df) and null_count == 0:
            alerts.append(f"✅ Perfect key column - all values are unique!")
        
        # Show alerts
        if alerts:
            for alert in alerts:
                if "✅" in alert:
                    st.success(alert)
                else:
                    st.warning(alert)
        
        # Advanced analysis in expandable sections
        col_left, col_right = st.columns(2)
        
        with col_left:
            with st.expander("📋 Sample Data Preview", expanded=False):
                # Sample data with metadata
                sample_data = df[column_name].dropna().head(10).tolist()
                
                if sample_data:
                    sample_df = pd.DataFrame({
                        'Sample Values': sample_data,
                        'Length': [len(str(x)) for x in sample_data],
                        'Type': [type(x).__name__ for x in sample_data]
                    })
                    st.dataframe(sample_df, hide_index=True, use_container_width=True)
                else:
                    st.info("No non-null values found")
        
        with col_right:
            with st.expander("🔍 Pattern Detection", expanded=False):
                # Pattern analysis for text columns
                if df[column_name].dtype == 'object':
                    patterns_found = []
                    
                    # Email pattern
                    email_pattern = df[column_name].str.contains(r'@\w+\.\w+', na=False, regex=True).sum()
                    if email_pattern > 0:
                        patterns_found.append(f"📧 Email addresses: {email_pattern} ({(email_pattern/len(df)*100):.1f}%)")
                    
                    # Phone pattern
                    phone_pattern = df[column_name].str.contains(r'\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b', na=False, regex=True).sum()
                    if phone_pattern > 0:
                        patterns_found.append(f"📞 Phone numbers: {phone_pattern} ({(phone_pattern/len(df)*100):.1f}%)")
                    
                    # Date pattern
                    date_pattern = df[column_name].str.contains(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', na=False, regex=True).sum()
                    if date_pattern > 0:
                        patterns_found.append(f"📅 Date-like values: {date_pattern} ({(date_pattern/len(df)*100):.1f}%)")
                    
                    # ID pattern (alphanumeric)
                    id_pattern = df[column_name].str.contains(r'^[A-Za-z0-9]+$', na=False, regex=True).sum()
                    if id_pattern > 0:
                        patterns_found.append(f"🆔 ID-like values: {id_pattern} ({(id_pattern/len(df)*100):.1f}%)")
                    
                    # Number pattern
                    number_pattern = df[column_name].str.contains(r'^\d+$', na=False, regex=True).sum()
                    if number_pattern > 0:
                        patterns_found.append(f"🔢 Numeric strings: {number_pattern} ({(number_pattern/len(df)*100):.1f}%)")
                    
                    if patterns_found:
                        for pattern in patterns_found:
                            st.info(pattern)
                    else:
                        st.write("No common patterns detected")
                
                else:
                    # Numeric column analysis
                    st.write("**Numeric Column Statistics:**")
                    try:
                        col_stats = df[column_name].describe()
                        st.write(f"• Mean: {col_stats['mean']:.2f}")
                        st.write(f"• Median: {col_stats['50%']:.2f}")
                        st.write(f"• Min: {col_stats['min']:.2f}")
                        st.write(f"• Max: {col_stats['max']:.2f}")
                        st.write(f"• Std Dev: {col_stats['std']:.2f}")
                    except:
                        st.write("Unable to calculate numeric statistics")
        
        # Data distribution insights
        with st.expander("📊 Data Distribution Insights", expanded=False):
            # Value length distribution for text
            if df[column_name].dtype == 'object':
                lengths = df[column_name].dropna().astype(str).str.len()
                if len(lengths) > 0:
                    st.write("**Text Length Statistics:**")
                    col_a, col_b, col_c = st.columns(3)
                    with col_a:
                        st.metric("Min Length", f"{lengths.min()}")
                    with col_b:
                        st.metric("Avg Length", f"{lengths.mean():.1f}")
                    with col_c:
                        st.metric("Max Length", f"{lengths.max()}")
                    
                    # Show most common lengths
                    length_counts = lengths.value_counts().head(5)
                    if len(length_counts) > 0:
                        st.write("**Most Common Text Lengths:**")
                        for length, count in length_counts.items():
                            percentage = (count / len(lengths)) * 100
                            st.write(f"• {length} characters: {count:,} values ({percentage:.1f}%)")
            
            # Top values
            st.write("**Most Frequent Values:**")
            top_values = df[column_name].value_counts().head(10)
            if len(top_values) > 0:
                for value, count in top_values.items():
                    percentage = (count / len(df)) * 100
                    # Truncate long values for display
                    display_value = str(value)
                    if len(display_value) > 50:
                        display_value = display_value[:47] + "..."
                    st.write(f"• `{display_value}`: {count:,} times ({percentage:.1f}%)")
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
                    result_row.update(b_data['data'])
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
                result_row.update(best_match['data'])
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
            
            # Enhanced column analysis for selected key columns
            if st.checkbox("📊 Show Advanced Key Column Analysis", help="Get detailed insights about data quality and matching potential"):
                st.write("---")
                col_analysis_1, col_analysis_2 = st.columns(2)
                
                with col_analysis_1:
                    st.write("### 📋 Sheet A Analysis")
                    comparator.show_column_analysis(comparator.df_a, key_col_a, "Sheet A")
                
                with col_analysis_2:
                    st.write("### 📋 Sheet B Analysis") 
                    comparator.show_column_analysis(comparator.df_b, key_col_b, "Sheet B")
        
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