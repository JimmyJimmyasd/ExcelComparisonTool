# Statistical Analysis Module for Excel Comparison Tool
# This module provides comprehensive statistical analysis capabilities

import pandas as pd
import numpy as np
from scipy import stats
from typing import Dict, List, Tuple, Optional
import warnings
warnings.filterwarnings('ignore')

class StatisticalAnalyzer:
    """
    Comprehensive statistical analysis for Excel data
    """
    
    def __init__(self):
        self.analysis_results = {}
    
    def analyze_dataframe(self, df: pd.DataFrame, sheet_name: str = "Sheet") -> Dict:
        """
        Perform comprehensive statistical analysis on a DataFrame
        
        Args:
            df: DataFrame to analyze
            sheet_name: Name of the sheet for reporting
            
        Returns:
            Dictionary containing all statistical analysis results
        """
        results = {
            'sheet_name': sheet_name,
            'basic_info': self._get_basic_info(df),
            'descriptive_stats': self._calculate_descriptive_stats(df),
            'missing_data_analysis': self._analyze_missing_data(df),
            'data_types_analysis': self._analyze_data_types(df),
            'numerical_analysis': self._analyze_numerical_columns(df),
            'categorical_analysis': self._analyze_categorical_columns(df)
        }
        
        self.analysis_results[sheet_name] = results
        return results
    
    def _get_basic_info(self, df: pd.DataFrame) -> Dict:
        """Get basic information about the dataset"""
        return {
            'total_rows': len(df),
            'total_columns': len(df.columns),
            'total_cells': len(df) * len(df.columns),
            'memory_usage': df.memory_usage(deep=True).sum(),
            'shape': df.shape
        }
    
    def _calculate_descriptive_stats(self, df: pd.DataFrame) -> Dict:
        """Calculate descriptive statistics for numerical columns"""
        numerical_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(numerical_cols) == 0:
            return {'message': 'No numerical columns found for analysis'}
        
        desc_stats = {}
        
        for col in numerical_cols:
            series = df[col].dropna()
            
            if len(series) == 0:
                continue
                
            desc_stats[col] = {
                'count': len(series),
                'mean': series.mean(),
                'median': series.median(),
                'mode': series.mode().iloc[0] if len(series.mode()) > 0 else None,
                'std_dev': series.std(),
                'variance': series.var(),
                'min': series.min(),
                'max': series.max(),
                'range': series.max() - series.min(),
                'q1': series.quantile(0.25),
                'q3': series.quantile(0.75),
                'iqr': series.quantile(0.75) - series.quantile(0.25),
                'skewness': stats.skew(series),
                'kurtosis': stats.kurtosis(series),
                'coefficient_of_variation': (series.std() / series.mean()) * 100 if series.mean() != 0 else 0
            }
        
        return desc_stats
    
    def _analyze_missing_data(self, df: pd.DataFrame) -> Dict:
        """Analyze missing data patterns"""
        missing_info = {}
        
        for col in df.columns:
            missing_count = df[col].isnull().sum()
            missing_percentage = (missing_count / len(df)) * 100
            
            missing_info[col] = {
                'missing_count': missing_count,
                'missing_percentage': missing_percentage,
                'present_count': len(df) - missing_count,
                'present_percentage': 100 - missing_percentage
            }
        
        # Overall missing data statistics
        total_missing = df.isnull().sum().sum()
        total_cells = len(df) * len(df.columns)
        
        missing_info['_summary'] = {
            'total_missing_cells': total_missing,
            'total_cells': total_cells,
            'overall_missing_percentage': (total_missing / total_cells) * 100,
            'columns_with_missing_data': len([col for col in df.columns if df[col].isnull().sum() > 0]),
            'complete_rows': len(df.dropna()),
            'incomplete_rows': len(df) - len(df.dropna())
        }
        
        return missing_info
    
    def _analyze_data_types(self, df: pd.DataFrame) -> Dict:
        """Analyze data types and suggest optimizations"""
        type_analysis = {}
        
        for col in df.columns:
            dtype = str(df[col].dtype)
            unique_count = df[col].nunique()
            null_count = df[col].isnull().sum()
            
            # Determine recommended type
            recommended_type = self._suggest_data_type(df[col])
            
            type_analysis[col] = {
                'current_type': dtype,
                'unique_values': unique_count,
                'null_values': null_count,
                'recommended_type': recommended_type,
                'memory_usage': df[col].memory_usage(deep=True),
                'sample_values': df[col].dropna().head(3).tolist()
            }
        
        return type_analysis
    
    def _suggest_data_type(self, series: pd.Series) -> str:
        """Suggest optimal data type for a series"""
        # Skip if mostly null
        if series.isnull().sum() / len(series) > 0.5:
            return "High null percentage - consider data cleaning"
        
        # Try to convert to numeric
        try:
            pd.to_numeric(series.dropna(), errors='raise')
            return "numeric"
        except (ValueError, TypeError):
            pass
        
        # Check if it's datetime
        try:
            pd.to_datetime(series.dropna(), errors='raise')
            return "datetime"
        except (ValueError, TypeError):
            pass
        
        # Check if it should be categorical
        unique_ratio = series.nunique() / len(series)
        if unique_ratio < 0.1 and series.nunique() < 50:
            return "categorical"
        
        return "text"
    
    def _analyze_numerical_columns(self, df: pd.DataFrame) -> Dict:
        """Detailed analysis of numerical columns"""
        numerical_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(numerical_cols) == 0:
            return {'message': 'No numerical columns found'}
        
        numerical_analysis = {}
        
        for col in numerical_cols:
            series = df[col].dropna()
            
            if len(series) == 0:
                continue
            
            # Outlier detection using IQR method
            q1 = series.quantile(0.25)
            q3 = series.quantile(0.75)
            iqr = q3 - q1
            lower_bound = q1 - 1.5 * iqr
            upper_bound = q3 + 1.5 * iqr
            outliers = series[(series < lower_bound) | (series > upper_bound)]
            
            # Distribution analysis
            numerical_analysis[col] = {
                'outliers_count': len(outliers),
                'outliers_percentage': (len(outliers) / len(series)) * 100,
                'outlier_bounds': {
                    'lower': lower_bound,
                    'upper': upper_bound
                },
                'distribution_info': {
                    'is_normal': self._test_normality(series),
                    'histogram_data': self._create_histogram_data(series),
                    'unique_values': series.nunique(),
                    'zeros_count': (series == 0).sum(),
                    'negative_count': (series < 0).sum(),
                    'positive_count': (series > 0).sum()
                }
            }
        
        return numerical_analysis
    
    def _analyze_categorical_columns(self, df: pd.DataFrame) -> Dict:
        """Analyze categorical/text columns"""
        categorical_cols = df.select_dtypes(include=['object', 'category']).columns
        
        if len(categorical_cols) == 0:
            return {'message': 'No categorical columns found'}
        
        categorical_analysis = {}
        
        for col in categorical_cols:
            series = df[col].dropna()
            
            if len(series) == 0:
                continue
            
            value_counts = series.value_counts()
            
            categorical_analysis[col] = {
                'unique_values': series.nunique(),
                'most_frequent': value_counts.index[0] if len(value_counts) > 0 else None,
                'most_frequent_count': value_counts.iloc[0] if len(value_counts) > 0 else 0,
                'least_frequent': value_counts.index[-1] if len(value_counts) > 0 else None,
                'least_frequent_count': value_counts.iloc[-1] if len(value_counts) > 0 else 0,
                'top_values': value_counts.head(10).to_dict(),
                'average_length': series.astype(str).str.len().mean(),
                'max_length': series.astype(str).str.len().max(),
                'min_length': series.astype(str).str.len().min(),
                'empty_strings': (series.astype(str).str.strip() == '').sum()
            }
        
        return categorical_analysis
    
    def _test_normality(self, series: pd.Series, alpha: float = 0.05) -> Dict:
        """Test if data follows normal distribution"""
        if len(series) < 8:
            return {'test': 'insufficient_data', 'is_normal': False}
        
        try:
            # Shapiro-Wilk test for normality
            stat, p_value = stats.shapiro(series.sample(min(5000, len(series))))  # Limit sample size
            
            return {
                'test': 'shapiro_wilk',
                'statistic': stat,
                'p_value': p_value,
                'is_normal': p_value > alpha,
                'alpha': alpha
            }
        except Exception:
            return {'test': 'failed', 'is_normal': False}
    
    def _create_histogram_data(self, series: pd.Series, bins: int = 20) -> Dict:
        """Create histogram data for plotting"""
        try:
            counts, bin_edges = np.histogram(series, bins=bins)
            
            return {
                'counts': counts.tolist(),
                'bin_edges': bin_edges.tolist(),
                'bin_centers': ((bin_edges[:-1] + bin_edges[1:]) / 2).tolist()
            }
        except Exception:
            return {'error': 'Could not create histogram data'}
    
    def calculate_correlation_matrix(self, df: pd.DataFrame) -> Dict:
        """Calculate correlation matrix for numerical columns"""
        numerical_df = df.select_dtypes(include=[np.number])
        
        if len(numerical_df.columns) < 2:
            return {'message': 'Need at least 2 numerical columns for correlation analysis'}
        
        # Calculate different types of correlations
        correlations = {
            'pearson': numerical_df.corr(method='pearson'),
            'spearman': numerical_df.corr(method='spearman'),
            'kendall': numerical_df.corr(method='kendall')
        }
        
        # Find strong correlations
        pearson_corr = correlations['pearson']
        strong_correlations = []
        
        for i in range(len(pearson_corr.columns)):
            for j in range(i+1, len(pearson_corr.columns)):
                col1, col2 = pearson_corr.columns[i], pearson_corr.columns[j]
                corr_value = pearson_corr.iloc[i, j]
                
                if abs(corr_value) > 0.7:  # Strong correlation threshold
                    strong_correlations.append({
                        'column1': col1,
                        'column2': col2,
                        'correlation': corr_value,
                        'strength': 'very_strong' if abs(corr_value) > 0.9 else 'strong'
                    })
        
        return {
            'correlation_matrices': correlations,
            'strong_correlations': strong_correlations,
            'correlation_summary': {
                'total_pairs': len(pearson_corr.columns) * (len(pearson_corr.columns) - 1) // 2,
                'strong_correlations_count': len(strong_correlations),
                'average_correlation': pearson_corr.values[np.triu_indices_from(pearson_corr.values, k=1)].mean()
            }
        }
    
    def generate_statistical_summary(self, df: pd.DataFrame, sheet_name: str = "Sheet") -> str:
        """Generate a human-readable statistical summary"""
        analysis = self.analyze_dataframe(df, sheet_name)
        
        summary_parts = []
        summary_parts.append(f"📊 Statistical Analysis Summary for {sheet_name}")
        summary_parts.append("=" * 50)
        
        # Basic info
        basic = analysis['basic_info']
        summary_parts.append(f"📋 Dataset Overview:")
        summary_parts.append(f"   • Rows: {basic['total_rows']:,}")
        summary_parts.append(f"   • Columns: {basic['total_columns']:,}")
        summary_parts.append(f"   • Total Cells: {basic['total_cells']:,}")
        summary_parts.append(f"   • Memory Usage: {basic['memory_usage']:,} bytes")
        
        # Missing data summary
        missing = analysis['missing_data_analysis']['_summary']
        summary_parts.append(f"\n🕳️ Missing Data:")
        summary_parts.append(f"   • Missing Cells: {missing['total_missing_cells']:,} ({missing['overall_missing_percentage']:.1f}%)")
        summary_parts.append(f"   • Complete Rows: {missing['complete_rows']:,}")
        summary_parts.append(f"   • Columns with Missing Data: {missing['columns_with_missing_data']}")
        
        # Numerical analysis
        if 'message' not in analysis['descriptive_stats']:
            num_cols = len(analysis['descriptive_stats'])
            summary_parts.append(f"\n📈 Numerical Analysis ({num_cols} columns):")
            
            # Find most variable column
            most_variable = None
            highest_cv = 0
            for col, stats in analysis['descriptive_stats'].items():
                if stats['coefficient_of_variation'] > highest_cv:
                    highest_cv = stats['coefficient_of_variation']
                    most_variable = col
            
            if most_variable:
                summary_parts.append(f"   • Most Variable Column: {most_variable} (CV: {highest_cv:.1f}%)")
        
        return "\n".join(summary_parts)