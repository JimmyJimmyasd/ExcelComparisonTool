import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Any, Optional
import re
from fuzzywuzzy import fuzz
from collections import Counter

# Handle visualization imports gracefully for deployment compatibility
try:
    import seaborn as sns
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    print("Warning: matplotlib/seaborn not available. Some visualizations will be disabled.")

try:
    import plotly.express as px
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False
    print("Warning: plotly not available. Interactive charts will be disabled.")

class DataQualityAssessment:
    """Comprehensive data quality analysis for Excel data"""
    
    def __init__(self, df: pd.DataFrame, sheet_name: str = "Data"):
        self.df = df.copy()
        self.sheet_name = sheet_name
        self.total_rows = len(df)
        self.total_columns = len(df.columns)
    
    def analyze_missing_data(self) -> Dict[str, Any]:
        """Comprehensive missing data analysis"""
        
        # Column-wise missing data
        missing_counts = self.df.isnull().sum()
        missing_percentages = (missing_counts / self.total_rows * 100).round(2)
        
        # Row-wise missing data
        rows_missing_counts = self.df.isnull().sum(axis=1)
        completely_empty_rows = (rows_missing_counts == self.total_columns).sum()
        
        # Missing data patterns
        missing_patterns = self._identify_missing_patterns()
        
        # Overall completeness
        total_cells = self.total_rows * self.total_columns
        missing_cells = self.df.isnull().sum().sum()
        overall_completeness = ((total_cells - missing_cells) / total_cells * 100).round(2)
        
        return {
            'column_missing_counts': missing_counts.to_dict(),
            'column_missing_percentage': missing_percentages.to_dict(),
            'rows_with_missing_data': (rows_missing_counts > 0).sum(),
            'completely_empty_rows': completely_empty_rows,
            'missing_patterns': missing_patterns,
            'overall_completeness': overall_completeness,
            'total_missing_cells': int(missing_cells),
            'columns_with_missing_data': (missing_counts > 0).sum(),
            'worst_columns': missing_percentages.nlargest(5).to_dict(),
            'best_columns': missing_percentages[missing_percentages == 0].index.tolist()
        }
    
    def _identify_missing_patterns(self) -> Dict[str, Any]:
        """Identify patterns in missing data"""
        
        # Check for systematic missing data (entire columns or rows)
        systematic_patterns = {}
        
        # Columns that are mostly empty (>90% missing)
        mostly_empty_cols = []
        for col in self.df.columns:
            missing_pct = (self.df[col].isnull().sum() / self.total_rows * 100)
            if missing_pct > 90:
                mostly_empty_cols.append(col)
        
        # Check for missing data clusters (consecutive missing values)
        cluster_patterns = {}
        for col in self.df.select_dtypes(include=[np.number]).columns:
            if self.df[col].isnull().any():
                # Find consecutive missing value runs
                is_null = self.df[col].isnull()
                runs = []
                current_run = 0
                
                for value in is_null:
                    if value:
                        current_run += 1
                    else:
                        if current_run > 0:
                            runs.append(current_run)
                        current_run = 0
                
                if current_run > 0:
                    runs.append(current_run)
                
                if runs:
                    cluster_patterns[col] = {
                        'max_consecutive_missing': max(runs),
                        'total_runs': len(runs),
                        'avg_run_length': np.mean(runs)
                    }
        
        return {
            'mostly_empty_columns': mostly_empty_cols,
            'consecutive_missing_clusters': cluster_patterns
        }
    
    def detect_duplicates(self, similarity_threshold: float = 0.9) -> Dict[str, Any]:
        """Detect exact and near-duplicate records"""
        
        # Exact duplicates
        exact_duplicates = self.df.duplicated()
        exact_duplicate_count = exact_duplicates.sum()
        exact_duplicate_rows = self.df[exact_duplicates].index.tolist()
        
        # Near duplicates using fuzzy matching (for string columns)
        near_duplicates = self._find_near_duplicates(similarity_threshold)
        
        # Duplicate analysis by columns
        column_duplicates = {}
        for col in self.df.columns:
            if self.df[col].dtype == 'object':  # String columns
                col_duplicates = self.df[col].duplicated()
                column_duplicates[col] = {
                    'duplicate_count': col_duplicates.sum(),
                    'unique_values': self.df[col].nunique(),
                    'duplicate_percentage': (col_duplicates.sum() / self.total_rows * 100).round(2)
                }
        
        # Subset duplicates (duplicates based on key columns)
        subset_duplicates = self._analyze_subset_duplicates()
        
        return {
            'exact_duplicates': {
                'count': int(exact_duplicate_count),
                'percentage': (exact_duplicate_count / self.total_rows * 100).round(2),
                'duplicate_rows': exact_duplicate_rows[:10]  # Show first 10
            },
            'near_duplicates': near_duplicates,
            'column_duplicates': column_duplicates,
            'subset_duplicates': subset_duplicates,
            'duplicate_summary': {
                'total_unique_rows': self.df.drop_duplicates().shape[0],
                'data_reduction_potential': (exact_duplicate_count / self.total_rows * 100).round(2)
            }
        }
    
    def _find_near_duplicates(self, threshold: float) -> Dict[str, Any]:
        """Find near-duplicate records using fuzzy string matching"""
        
        near_duplicates = []
        string_columns = self.df.select_dtypes(include=['object']).columns
        
        if len(string_columns) == 0:
            return {'pairs': [], 'summary': 'No string columns for fuzzy matching'}
        
        # Sample data for performance (if dataset is large)
        sample_size = min(1000, len(self.df))
        df_sample = self.df.sample(n=sample_size) if len(self.df) > sample_size else self.df
        
        # Create composite string for comparison
        df_sample['_composite'] = df_sample[string_columns].astype(str).apply(
            lambda x: ' '.join(x.fillna('')), axis=1
        )
        
        # Compare pairs
        pairs_checked = 0
        max_pairs = 500  # Limit for performance
        
        for i, row1 in df_sample.iterrows():
            if pairs_checked >= max_pairs:
                break
                
            for j, row2 in df_sample.iterrows():
                if i >= j or pairs_checked >= max_pairs:
                    continue
                
                similarity = fuzz.ratio(row1['_composite'], row2['_composite']) / 100
                
                if similarity >= threshold and similarity < 1.0:  # Near duplicates, not exact
                    near_duplicates.append({
                        'row1_index': i,
                        'row2_index': j,
                        'similarity_score': round(similarity, 3),
                        'row1_data': row1[string_columns].to_dict(),
                        'row2_data': row2[string_columns].to_dict()
                    })
                
                pairs_checked += 1
        
        return {
            'pairs': near_duplicates[:20],  # Show top 20
            'total_found': len(near_duplicates),
            'summary': f'Found {len(near_duplicates)} near-duplicate pairs with ≥{threshold*100}% similarity'
        }
    
    def _analyze_subset_duplicates(self) -> Dict[str, Any]:
        """Analyze duplicates based on key column combinations"""
        
        subset_analysis = {}
        
        # Try different column combinations for duplicate detection
        potential_key_columns = []
        
        # Identify potential key columns (low cardinality or ID-like names)
        for col in self.df.columns:
            # Safely convert column name to string and then to lowercase
            col_str = str(col).lower() if col is not None else ''
            unique_ratio = self.df[col].nunique() / self.total_rows
            
            # ID columns or columns with reasonable uniqueness
            if any(keyword in col_str for keyword in ['id', 'key', 'code', 'number']) or unique_ratio > 0.7:
                potential_key_columns.append(col)
        
        # Analyze duplicates for key column combinations
        if potential_key_columns:
            for col in potential_key_columns[:3]:  # Limit to first 3 key columns
                duplicates_in_col = self.df.duplicated(subset=[col])
                subset_analysis[col] = {
                    'duplicate_count': duplicates_in_col.sum(),
                    'unique_values': self.df[col].nunique(),
                    'duplicate_percentage': (duplicates_in_col.sum() / self.total_rows * 100).round(2)
                }
        
        return subset_analysis
    
    def validate_data_types(self) -> Dict[str, Any]:
        """Comprehensive data type validation and recommendations"""
        
        validation_results = {}
        
        for col in self.df.columns:
            col_analysis = {
                'current_type': str(self.df[col].dtype),
                'recommended_type': None,
                'issues': [],
                'conversion_possible': False,
                'sample_values': self.df[col].dropna().head(5).tolist()
            }
            
            # Analyze column content
            non_null_values = self.df[col].dropna()
            
            if len(non_null_values) == 0:
                col_analysis['issues'].append('Column is completely empty')
                validation_results[col] = col_analysis
                continue
            
            # Check for numeric data stored as text
            if self.df[col].dtype == 'object':
                numeric_conversion = self._check_numeric_conversion(non_null_values)
                if numeric_conversion['convertible']:
                    col_analysis['recommended_type'] = numeric_conversion['recommended_type']
                    col_analysis['conversion_possible'] = True
                    col_analysis['issues'].append(f"Numeric data stored as text - can convert to {numeric_conversion['recommended_type']}")
                
                # Check for date data stored as text
                date_conversion = self._check_date_conversion(non_null_values)
                if date_conversion['convertible']:
                    col_analysis['recommended_type'] = 'datetime64[ns]'
                    col_analysis['conversion_possible'] = True
                    col_analysis['issues'].append(f"Date data stored as text - can convert to datetime")
                
                # Check for boolean data stored as text
                bool_conversion = self._check_boolean_conversion(non_null_values)
                if bool_conversion['convertible']:
                    col_analysis['recommended_type'] = 'bool'
                    col_analysis['conversion_possible'] = True
                    col_analysis['issues'].append("Boolean data stored as text - can convert to bool")
            
            # Check for inconsistent data in numeric columns
            elif self.df[col].dtype in ['int64', 'float64']:
                # Check for mixed data types
                string_like_values = []
                for val in non_null_values.head(100):  # Sample check
                    if isinstance(val, str):
                        string_like_values.append(val)
                
                if string_like_values:
                    col_analysis['issues'].append(f"Mixed data types detected - found text values: {string_like_values[:3]}")
            
            # Check data range and outliers for numeric columns
            if self.df[col].dtype in ['int64', 'float64']:
                outlier_analysis = self._detect_outliers(self.df[col])
                if outlier_analysis['outlier_count'] > 0:
                    col_analysis['issues'].append(f"Contains {outlier_analysis['outlier_count']} potential outliers")
            
            validation_results[col] = col_analysis
        
        # Overall assessment
        total_issues = sum(len(col_data['issues']) for col_data in validation_results.values())
        convertible_columns = sum(1 for col_data in validation_results.values() if col_data['conversion_possible'])
        
        summary = {
            'total_columns_analyzed': len(validation_results),
            'columns_with_issues': sum(1 for col_data in validation_results.values() if col_data['issues']),
            'total_issues_found': total_issues,
            'convertible_columns': convertible_columns,
            'data_quality_score': max(0, 100 - (total_issues * 10))  # Simple scoring
        }
        
        return {
            'column_analysis': validation_results,
            'summary': summary
        }
    
    def _check_numeric_conversion(self, series: pd.Series) -> Dict[str, Any]:
        """Check if string column can be converted to numeric"""
        
        convertible_count = 0
        sample_values = series.head(50)  # Sample for performance
        
        for value in sample_values:
            if isinstance(value, str):
                # Remove common formatting
                cleaned_value = re.sub(r'[,$%\s]', '', str(value))
                try:
                    float(cleaned_value)
                    convertible_count += 1
                except (ValueError, TypeError):
                    pass
        
        conversion_rate = convertible_count / len(sample_values)
        
        if conversion_rate > 0.8:  # 80% convertible
            # Determine if int or float
            has_decimals = any('.' in str(val) for val in sample_values if isinstance(val, str))
            recommended_type = 'float64' if has_decimals else 'int64'
            
            return {
                'convertible': True,
                'recommended_type': recommended_type,
                'conversion_rate': conversion_rate
            }
        
        return {'convertible': False, 'conversion_rate': conversion_rate}
    
    def _check_date_conversion(self, series: pd.Series) -> Dict[str, Any]:
        """Check if string column can be converted to datetime"""
        
        convertible_count = 0
        sample_values = series.head(50)
        
        for value in sample_values:
            if isinstance(value, str):
                try:
                    pd.to_datetime(value)
                    convertible_count += 1
                except (ValueError, TypeError):
                    pass
        
        conversion_rate = convertible_count / len(sample_values)
        
        return {
            'convertible': conversion_rate > 0.7,  # 70% convertible
            'conversion_rate': conversion_rate
        }
    
    def _check_boolean_conversion(self, series: pd.Series) -> Dict[str, Any]:
        """Check if string column represents boolean data"""
        
        unique_values = set(str(val).lower().strip() for val in series.unique())
        boolean_values = {'true', 'false', 'yes', 'no', '1', '0', 'y', 'n', 't', 'f'}
        
        is_boolean = unique_values.issubset(boolean_values) and len(unique_values) <= 4
        
        return {
            'convertible': is_boolean,
            'unique_values': list(unique_values)
        }
    
    def _detect_outliers(self, series: pd.Series) -> Dict[str, Any]:
        """Detect outliers using IQR method"""
        
        if series.dtype not in ['int64', 'float64']:
            return {'outlier_count': 0, 'outliers': []}
        
        Q1 = series.quantile(0.25)
        Q3 = series.quantile(0.75)
        IQR = Q3 - Q1
        
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        
        outliers = series[(series < lower_bound) | (series > upper_bound)]
        
        return {
            'outlier_count': len(outliers),
            'outliers': outliers.head(10).tolist(),  # Show first 10
            'lower_bound': lower_bound,
            'upper_bound': upper_bound
        }
    
    def create_missing_data_heatmap(self) -> go.Figure:
        """Create an interactive heatmap showing missing data patterns"""
        
        # Sample data if too large for visualization
        df_viz = self.df.head(500) if len(self.df) > 500 else self.df
        
        # Create missing data matrix
        missing_matrix = df_viz.isnull().astype(int)
        
        fig = go.Figure(data=go.Heatmap(
            z=missing_matrix.values,
            x=missing_matrix.columns,
            y=list(range(len(missing_matrix))),
            colorscale=[[0, 'lightblue'], [1, 'red']],
            showscale=True,
            colorbar=dict(
                title="Missing Data",
                tickvals=[0, 1],
                ticktext=["Present", "Missing"]
            )
        ))
        
        fig.update_layout(
            title=f"Missing Data Heatmap - {self.sheet_name}",
            xaxis_title="Columns",
            yaxis_title="Row Index",
            height=400,
            xaxis={'tickangle': 45}
        )
        
        return fig
    
    def create_data_quality_summary_chart(self, quality_results: Dict[str, Any]) -> go.Figure:
        """Create a comprehensive data quality summary visualization"""
        
        # Create subplots
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=(
                'Missing Data by Column',
                'Data Type Issues',
                'Duplicate Analysis',
                'Overall Quality Score'
            ),
            specs=[[{"type": "bar"}, {"type": "bar"}],
                   [{"type": "bar"}, {"type": "indicator"}]]
        )
        
        # Missing data chart
        missing_data = quality_results['missing_data']['column_missing_percentage']
        cols_with_missing = {k: v for k, v in missing_data.items() if v > 0}
        
        if cols_with_missing:
            fig.add_trace(
                go.Bar(
                    x=list(cols_with_missing.keys()),
                    y=list(cols_with_missing.values()),
                    name="Missing %",
                    marker_color='red'
                ),
                row=1, col=1
            )
        
        # Data type issues
        type_issues = quality_results['data_type_validation']['column_analysis']
        issue_counts = {col: len(data['issues']) for col, data in type_issues.items() if data['issues']}
        
        if issue_counts:
            fig.add_trace(
                go.Bar(
                    x=list(issue_counts.keys()),
                    y=list(issue_counts.values()),
                    name="Issues",
                    marker_color='orange'
                ),
                row=1, col=2
            )
        
        # Duplicate analysis
        duplicate_data = quality_results['duplicates']
        duplicate_types = ['Exact', 'Column-based']
        duplicate_counts = [
            duplicate_data['exact_duplicates']['count'],
            sum(col_data['duplicate_count'] for col_data in duplicate_data['column_duplicates'].values())
        ]
        
        fig.add_trace(
            go.Bar(
                x=duplicate_types,
                y=duplicate_counts,
                name="Duplicates",
                marker_color='yellow'
            ),
            row=2, col=1
        )
        
        # Overall quality score
        quality_score = quality_results['data_type_validation']['summary']['data_quality_score']
        
        fig.add_trace(
            go.Indicator(
                mode="gauge+number",
                value=quality_score,
                domain={'x': [0, 1], 'y': [0, 1]},
                title={'text': "Quality Score"},
                gauge={
                    'axis': {'range': [None, 100]},
                    'bar': {'color': "darkblue"},
                    'steps': [
                        {'range': [0, 50], 'color': "lightgray"},
                        {'range': [50, 80], 'color': "gray"}],
                    'threshold': {
                        'line': {'color': "red", 'width': 4},
                        'thickness': 0.75,
                        'value': 90}
                }
            ),
            row=2, col=2
        )
        
        fig.update_layout(height=600, showlegend=False, title_text="Data Quality Assessment Summary")
        
        return fig
    
    def generate_quality_report(self) -> Dict[str, Any]:
        """Generate comprehensive data quality report"""
        
        # Analyzing missing data...
        missing_data = self.analyze_missing_data()
        
        # Detecting duplicates...
        duplicates = self.detect_duplicates()
        
        # Validating data types...
        data_type_validation = self.validate_data_types()
        
        # Calculate overall quality metrics
        missing_score = max(0, 100 - missing_data['column_missing_percentage'].get('overall', 0))
        duplicate_score = max(0, 100 - duplicates['exact_duplicates']['percentage'])
        type_score = data_type_validation['summary']['data_quality_score']
        
        overall_score = (missing_score + duplicate_score + type_score) / 3
        
        return {
            'missing_data': missing_data,
            'duplicates': duplicates,
            'data_type_validation': data_type_validation,
            'overall_metrics': {
                'missing_data_score': round(missing_score, 1),
                'duplicate_score': round(duplicate_score, 1),
                'data_type_score': round(type_score, 1),
                'overall_quality_score': round(overall_score, 1)
            },
            'recommendations': self._generate_recommendations(missing_data, duplicates, data_type_validation)
        }
    
    def _generate_recommendations(self, missing_data: Dict, duplicates: Dict, type_validation: Dict) -> List[str]:
        """Generate actionable recommendations based on quality analysis"""
        
        recommendations = []
        
        # Missing data recommendations
        if missing_data['overall_completeness'] < 90:
            recommendations.append(f"📊 **Data Completeness**: Only {missing_data['overall_completeness']}% complete. Consider investigating missing data sources.")
        
        if missing_data['worst_columns']:
            worst_col = max(missing_data['worst_columns'], key=missing_data['worst_columns'].get)
            worst_pct = missing_data['worst_columns'][worst_col]
            recommendations.append(f"🕳️ **Critical Missing Data**: Column '{worst_col}' is {worst_pct}% missing. Consider removing or imputing values.")
        
        # Duplicate recommendations
        if duplicates['exact_duplicates']['percentage'] > 5:
            recommendations.append(f"🔄 **Duplicate Cleanup**: {duplicates['exact_duplicates']['percentage']}% exact duplicates found. Consider deduplication to save {duplicates['duplicate_summary']['data_reduction_potential']}% space.")
        
        # Data type recommendations
        convertible_count = type_validation['summary']['convertible_columns']
        if convertible_count > 0:
            recommendations.append(f"🔧 **Data Type Optimization**: {convertible_count} columns can be converted to more appropriate data types for better performance.")
        
        # Performance recommendations
        mostly_empty_cols = missing_data.get('missing_patterns', {}).get('mostly_empty_columns', [])
        if len(mostly_empty_cols) > 0:
            recommendations.append(f"🗑️ **Column Cleanup**: {len(mostly_empty_cols)} columns are mostly empty (>90% missing). Consider removing them.")
        
        # General recommendations
        if not recommendations:
            recommendations.append("✅ **Excellent Quality**: Your data shows good quality metrics across all dimensions!")
        
        return recommendations