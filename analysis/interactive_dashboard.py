"""
Interactive Dashboard Module for Excel Comparison Tool
Provides comprehensive interactive visualizations using Plotly for data analysis and comparison.
"""

import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.figure_factory as ff
from typing import Dict, List, Any, Optional, Tuple, Union
import warnings
warnings.filterwarnings('ignore')

class InteractiveDashboard:
    """
    Create interactive dashboards and visualizations for Excel data analysis
    """
    
    def __init__(self):
        """Initialize the Interactive Dashboard"""
        self.color_palette = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
                             '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
        self.charts = {}
        
    def create_comprehensive_dashboard(self, analysis_results: Dict[str, Any], 
                                    comparison_data: Optional[Dict] = None) -> Dict[str, go.Figure]:
        """
        Create a comprehensive dashboard with multiple interactive charts
        
        Args:
            analysis_results: Results from statistical analysis
            comparison_data: Optional comparison data between datasets
            
        Returns:
            Dictionary of Plotly figures for different aspects of analysis
        """
        dashboard_charts = {}
        
        # 1. Overview Summary Chart
        dashboard_charts['overview'] = self._create_overview_chart(analysis_results)
        
        # 2. Missing Data Visualization
        if 'missing_data_analysis' in analysis_results:
            dashboard_charts['missing_data'] = self._create_missing_data_chart(analysis_results['missing_data_analysis'])
        
        # 3. Correlation Heatmap
        dashboard_charts['correlation'] = self._create_correlation_heatmap(analysis_results)
        
        # 4. Distribution Analysis
        dashboard_charts['distributions'] = self._create_distribution_charts(analysis_results)
        
        # 5. Outlier Analysis
        dashboard_charts['outliers'] = self._create_outlier_charts(analysis_results)
        
        # 6. Data Quality Dashboard
        dashboard_charts['data_quality'] = self._create_data_quality_dashboard(analysis_results)
        
        # 7. Comparison Charts (if comparison data provided)
        if comparison_data:
            dashboard_charts['comparison'] = self._create_comparison_charts(comparison_data)
        
        self.charts = dashboard_charts
        return dashboard_charts
    
    def _create_overview_chart(self, analysis_results: Dict[str, Any]) -> go.Figure:
        """Create an overview summary chart"""
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=['Dataset Shape', 'Data Types Distribution', 'Missing Data Overview', 'Memory Usage'],
            specs=[[{'type': 'bar'}, {'type': 'pie'}],
                   [{'type': 'bar'}, {'type': 'indicator'}]]
        )
        
        # Dataset shape
        if 'basic_info' in analysis_results:
            basic_info = analysis_results['basic_info']
            fig.add_trace(
                go.Bar(
                    x=['Rows', 'Columns'],
                    y=[basic_info['total_rows'], basic_info['total_columns']],
                    name='Dataset Shape',
                    marker_color=['#1f77b4', '#ff7f0e']
                ),
                row=1, col=1
            )
        
        # Data types distribution
        if 'data_types_analysis' in analysis_results:
            type_counts = {}
            for col, analysis in analysis_results['data_types_analysis'].items():
                dtype = analysis['current_type']
                type_counts[dtype] = type_counts.get(dtype, 0) + 1
            
            fig.add_trace(
                go.Pie(
                    labels=list(type_counts.keys()),
                    values=list(type_counts.values()),
                    name='Data Types'
                ),
                row=1, col=2
            )
        
        # Missing data overview
        if 'missing_data_analysis' in analysis_results:
            missing_data = analysis_results['missing_data_analysis']
            cols_with_missing = []
            missing_percentages = []
            
            for col, data in missing_data.items():
                if col != '_summary' and data['missing_percentage'] > 0:
                    cols_with_missing.append(col[:20] + '...' if len(col) > 20 else col)
                    missing_percentages.append(data['missing_percentage'])
            
            if cols_with_missing:
                fig.add_trace(
                    go.Bar(
                        x=cols_with_missing[:10],  # Top 10 columns with missing data
                        y=missing_percentages[:10],
                        name='Missing %',
                        marker_color='#d62728'
                    ),
                    row=2, col=1
                )
        
        # Memory usage indicator
        if 'basic_info' in analysis_results:
            memory_mb = analysis_results['basic_info']['memory_usage'] / 1024 / 1024
            fig.add_trace(
                go.Indicator(
                    mode='gauge+number',
                    value=memory_mb,
                    title={'text': 'Memory (MB)'},
                    gauge={
                        'axis': {'range': [0, max(100, memory_mb * 1.2)]},
                        'bar': {'color': '#2ca02c'},
                        'steps': [
                            {'range': [0, 50], 'color': 'lightgray'},
                            {'range': [50, 100], 'color': 'gray'}
                        ],
                        'threshold': {
                            'line': {'color': 'red', 'width': 4},
                            'thickness': 0.75,
                            'value': 100
                        }
                    }
                ),
                row=2, col=2
            )
        
        fig.update_layout(
            title_text="📊 Dataset Overview Dashboard",
            showlegend=False,
            height=600
        )
        
        return fig
    
    def _create_missing_data_chart(self, missing_data_analysis: Dict[str, Any]) -> go.Figure:
        """Create interactive missing data visualization"""
        # Prepare data
        columns = []
        missing_percentages = []
        missing_counts = []
        
        for col, data in missing_data_analysis.items():
            if col != '_summary':
                columns.append(col)
                missing_percentages.append(data['missing_percentage'])
                missing_counts.append(data['missing_count'])
        
        # Create subplot with secondary y-axis
        fig = make_subplots(
            rows=2, cols=1,
            subplot_titles=['Missing Data by Percentage', 'Missing Data Heatmap Pattern'],
            vertical_spacing=0.1,
            specs=[[{'secondary_y': True}], [{'type': 'heatmap'}]]
        )
        
        # Bar chart for missing percentages
        fig.add_trace(
            go.Bar(
                x=columns,
                y=missing_percentages,
                name='Missing %',
                marker_color='#d62728',
                hovertemplate='<b>%{x}</b><br>Missing: %{y:.1f}%<extra></extra>'
            ),
            row=1, col=1
        )
        
        # Line chart for missing counts
        fig.add_trace(
            go.Scatter(
                x=columns,
                y=missing_counts,
                mode='lines+markers',
                name='Missing Count',
                line=dict(color='#ff7f0e', width=2),
                marker=dict(size=6),
                yaxis='y2',
                hovertemplate='<b>%{x}</b><br>Count: %{y}<extra></extra>'
            ),
            row=1, col=1,
            secondary_y=True
        )
        
        # Create missing data pattern heatmap (simplified)
        if len(columns) > 0:
            # Create a simple pattern representation
            pattern_data = [[missing_percentages[i] if j == 0 else 0 for j in range(min(10, len(columns)))] 
                           for i in range(min(10, len(columns)))]
            
            fig.add_trace(
                go.Heatmap(
                    z=pattern_data,
                    x=columns[:10],
                    y=['Pattern'] * len(columns[:10]),
                    colorscale='Reds',
                    showscale=True,
                    hovertemplate='<b>%{x}</b><br>Missing: %{z:.1f}%<extra></extra>'
                ),
                row=2, col=1
            )
        
        # Update layout
        fig.update_xaxes(title_text="Columns", row=1, col=1)
        fig.update_yaxes(title_text="Missing Percentage (%)", row=1, col=1)
        fig.update_yaxes(title_text="Missing Count", secondary_y=True, row=1, col=1)
        
        fig.update_layout(
            title_text="🕳️ Missing Data Analysis",
            height=700,
            showlegend=True
        )
        
        return fig
    
    def _create_correlation_heatmap(self, analysis_results: Dict[str, Any]) -> go.Figure:
        """Create interactive correlation heatmap"""
        if 'correlation_matrices' not in analysis_results:
            # Create empty figure with message
            fig = go.Figure()
            fig.add_annotation(
                text="No correlation data available<br>Need at least 2 numeric columns",
                xref="paper", yref="paper",
                x=0.5, y=0.5, xanchor='center', yanchor='middle',
                showarrow=False,
                font=dict(size=16)
            )
            fig.update_layout(title="🔗 Correlation Analysis")
            return fig
        
        corr_data = analysis_results['correlation_matrices']['pearson']
        
        # Create correlation heatmap
        fig = go.Figure(data=go.Heatmap(
            z=corr_data.values,
            x=corr_data.columns,
            y=corr_data.index,
            colorscale='RdBu',
            zmid=0,
            text=np.round(corr_data.values, 3),
            texttemplate='%{text}',
            textfont={'size': 10},
            hovertemplate='<b>%{x} vs %{y}</b><br>Correlation: %{z:.3f}<extra></extra>'
        ))
        
        # Add annotations for strong correlations
        strong_corrs = analysis_results.get('strong_correlations', [])
        if strong_corrs:
            annotations_text = f"Strong Correlations Found: {len(strong_corrs)}"
            fig.add_annotation(
                text=annotations_text,
                xref="paper", yref="paper",
                x=0.02, y=0.98,
                xanchor='left', yanchor='top',
                bgcolor="rgba(255,255,255,0.8)",
                bordercolor="gray",
                borderwidth=1
            )
        
        fig.update_layout(
            title="🔗 Correlation Matrix (Pearson)",
            xaxis_title="Variables",
            yaxis_title="Variables",
            height=600
        )
        
        return fig
    
    def _create_distribution_charts(self, analysis_results: Dict[str, Any]) -> go.Figure:
        """Create distribution analysis charts"""
        if 'numerical_analysis' not in analysis_results or 'message' in analysis_results['numerical_analysis']:
            fig = go.Figure()
            fig.add_annotation(
                text="No numerical data available for distribution analysis",
                xref="paper", yref="paper", x=0.5, y=0.5,
                xanchor='center', yanchor='middle',
                showarrow=False, font=dict(size=16)
            )
            fig.update_layout(title="📊 Distribution Analysis")
            return fig
        
        numerical_data = analysis_results['numerical_analysis']
        
        # Create subplots for multiple distributions
        cols_count = min(3, len(numerical_data))
        rows_count = (len(numerical_data) + cols_count - 1) // cols_count
        
        # Calculate appropriate vertical spacing based on number of rows
        # Maximum spacing is (1 / (rows - 1)) if rows > 1, else use default
        if rows_count > 1:
            max_vertical_spacing = 1.0 / (rows_count - 1)
            vertical_spacing = min(0.1, max_vertical_spacing * 0.8)  # Use 80% of max to be safe
        else:
            vertical_spacing = 0.1
        
        # Create enhanced subplot titles with normality information
        enhanced_titles = []
        for col_name, col_data in numerical_data.items():
            normality_status = "✓ Normal" if col_data['distribution_info']['is_normal']['is_normal'] else "✗ Non-normal"
            enhanced_titles.append(f"{col_name} ({normality_status})")
        
        fig = make_subplots(
            rows=rows_count,
            cols=cols_count,
            subplot_titles=enhanced_titles,
            vertical_spacing=vertical_spacing
        )
        
        for idx, (col_name, col_data) in enumerate(numerical_data.items()):
            row = idx // cols_count + 1
            col = idx % cols_count + 1
            
            if 'distribution_info' in col_data and 'histogram_data' in col_data['distribution_info']:
                hist_data = col_data['distribution_info']['histogram_data']
                
                fig.add_trace(
                    go.Histogram(
                        x=hist_data['bin_centers'],
                        y=hist_data['counts'],
                        name=col_name,
                        marker_color=self.color_palette[idx % len(self.color_palette)],
                        opacity=0.7,
                        hovertemplate=f'<b>{col_name}</b><br>Value: %{{x}}<br>Count: %{{y}}<extra></extra>'
                    ),
                    row=row, col=col
                )
                
                # Normality information is now included in the subplot titles
        
        fig.update_layout(
            title="📊 Distribution Analysis",
            height=200 * rows_count + 100,
            showlegend=False
        )
        
        return fig
    
    def _create_outlier_charts(self, analysis_results: Dict[str, Any]) -> go.Figure:
        """Create outlier analysis visualization"""
        if 'numerical_analysis' not in analysis_results or 'message' in analysis_results['numerical_analysis']:
            fig = go.Figure()
            fig.add_annotation(
                text="No numerical data available for outlier analysis",
                xref="paper", yref="paper", x=0.5, y=0.5,
                xanchor='center', yanchor='middle',
                showarrow=False, font=dict(size=16)
            )
            fig.update_layout(title="⚠️ Outlier Analysis")
            return fig
        
        numerical_data = analysis_results['numerical_analysis']
        
        # Prepare data for outlier summary
        columns = []
        outlier_counts = []
        outlier_percentages = []
        
        for col_name, col_data in numerical_data.items():
            columns.append(col_name)
            outlier_counts.append(col_data['outliers_count'])
            outlier_percentages.append(col_data['outliers_percentage'])
        
        # Create subplot
        fig = make_subplots(
            rows=1, cols=2,
            subplot_titles=['Outlier Counts by Column', 'Outlier Percentages'],
            specs=[[{'type': 'bar'}, {'type': 'scatter'}]]
        )
        
        # Outlier counts bar chart
        fig.add_trace(
            go.Bar(
                x=columns,
                y=outlier_counts,
                name='Outlier Count',
                marker_color='#d62728',
                hovertemplate='<b>%{x}</b><br>Outliers: %{y}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # Outlier percentages scatter plot
        fig.add_trace(
            go.Scatter(
                x=columns,
                y=outlier_percentages,
                mode='markers+lines',
                name='Outlier %',
                marker=dict(
                    size=10,
                    color=outlier_percentages,
                    colorscale='Reds',
                    showscale=True,
                    colorbar=dict(title="Outlier %")
                ),
                line=dict(color='red', width=2),
                hovertemplate='<b>%{x}</b><br>Outlier %: %{y:.2f}%<extra></extra>'
            ),
            row=1, col=2
        )
        
        fig.update_layout(
            title="⚠️ Outlier Analysis Summary",
            height=500,
            showlegend=False
        )
        
        return fig
    
    def _create_data_quality_dashboard(self, analysis_results: Dict[str, Any]) -> go.Figure:
        """Create comprehensive data quality dashboard"""
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=[
                'Data Completeness Score',
                'Column Quality Scores',
                'Data Type Recommendations',
                'Overall Quality Metrics'
            ],
            specs=[
                [{'type': 'indicator'}, {'type': 'bar'}],
                [{'type': 'pie'}, {'type': 'table'}]
            ]
        )
        
        # 1. Data Completeness Score
        if 'missing_data_analysis' in analysis_results:
            missing_summary = analysis_results['missing_data_analysis']['_summary']
            completeness_score = 100 - missing_summary['overall_missing_percentage']
            
            fig.add_trace(
                go.Indicator(
                    mode='gauge+number+delta',
                    value=completeness_score,
                    title={'text': 'Completeness Score'},
                    delta={'reference': 95},
                    gauge={
                        'axis': {'range': [None, 100]},
                        'bar': {'color': 'green' if completeness_score >= 95 else 'orange' if completeness_score >= 80 else 'red'},
                        'steps': [
                            {'range': [0, 70], 'color': 'lightgray'},
                            {'range': [70, 85], 'color': 'yellow'},
                            {'range': [85, 100], 'color': 'lightgreen'}
                        ],
                        'threshold': {
                            'line': {'color': 'red', 'width': 4},
                            'thickness': 0.75,
                            'value': 90
                        }
                    }
                ),
                row=1, col=1
            )
        
        # 2. Column Quality Scores
        if 'missing_data_analysis' in analysis_results:
            columns = []
            quality_scores = []
            
            for col, data in analysis_results['missing_data_analysis'].items():
                if col != '_summary':
                    columns.append(col[:15] + '...' if len(col) > 15 else col)
                    quality_score = 100 - data['missing_percentage']
                    quality_scores.append(quality_score)
            
            colors = ['green' if score >= 95 else 'orange' if score >= 80 else 'red' for score in quality_scores]
            
            fig.add_trace(
                go.Bar(
                    x=columns[:10],  # Top 10 columns
                    y=quality_scores[:10],
                    name='Quality Score',
                    marker_color=colors[:10],
                    hovertemplate='<b>%{x}</b><br>Quality: %{y:.1f}%<extra></extra>'
                ),
                row=1, col=2
            )
        
        # 3. Data Type Recommendations
        if 'data_types_analysis' in analysis_results:
            type_recommendations = {'Optimal': 0, 'Needs Optimization': 0, 'Needs Review': 0}
            
            for col, analysis in analysis_results['data_types_analysis'].items():
                # Check if suggestions exist and handle safely
                suggestions = analysis.get('suggestions', [])
                if suggestions and len(suggestions) > 0:
                    if len(suggestions) > 1:
                        type_recommendations['Needs Review'] += 1
                    else:
                        type_recommendations['Needs Optimization'] += 1
                else:
                    type_recommendations['Optimal'] += 1
            
            fig.add_trace(
                go.Pie(
                    labels=list(type_recommendations.keys()),
                    values=list(type_recommendations.values()),
                    marker_colors=['green', 'orange', 'red'],
                    name='Type Quality'
                ),
                row=2, col=1
            )
        
        # 4. Quality Metrics Table
        metrics_data = []
        if 'basic_info' in analysis_results:
            basic_info = analysis_results['basic_info']
            metrics_data.append(['Dataset Size', f"{basic_info['total_rows']:,} rows × {basic_info['total_columns']} cols"])
        
        if 'missing_data_analysis' in analysis_results:
            missing_summary = analysis_results['missing_data_analysis']['_summary']
            metrics_data.append(['Completeness', f"{100 - missing_summary['overall_missing_percentage']:.1f}%"])
            metrics_data.append(['Complete Rows', f"{missing_summary['complete_rows']:,} ({missing_summary['complete_rows']/basic_info['total_rows']*100:.1f}%)"])
        
        if 'numerical_analysis' in analysis_results and 'message' not in analysis_results['numerical_analysis']:
            total_outliers = sum([data['outliers_count'] for data in analysis_results['numerical_analysis'].values()])
            metrics_data.append(['Outliers Detected', f"{total_outliers:,}"])
        
        if metrics_data:
            fig.add_trace(
                go.Table(
                    header=dict(values=['Metric', 'Value'], fill_color='lightblue'),
                    cells=dict(values=list(zip(*metrics_data)), fill_color='white')
                ),
                row=2, col=2
            )
        
        fig.update_layout(
            title="🎯 Data Quality Dashboard",
            height=700,
            showlegend=False
        )
        
        return fig
    
    def _create_comparison_charts(self, comparison_data: Dict[str, Any]) -> go.Figure:
        """Create comparison charts between datasets"""
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=[
                'Dataset Size Comparison',
                'Data Quality Comparison',
                'Statistical Test Results',
                'Missing Data Comparison'
            ],
            specs=[
                [{'type': 'bar'}, {'type': 'bar'}],
                [{'type': 'heatmap'}, {'type': 'scatter'}]
            ]
        )
        
        datasets = comparison_data['datasets']
        dataset_names = list(datasets.keys())
        
        # 1. Dataset Size Comparison
        sizes = [datasets[name]['basic_info']['total_cells'] for name in dataset_names]
        fig.add_trace(
            go.Bar(
                x=dataset_names,
                y=sizes,
                name='Dataset Size',
                marker_color=['#1f77b4', '#ff7f0e'],
                hovertemplate='<b>%{x}</b><br>Total Cells: %{y:,}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # 2. Data Quality Comparison (Completeness)
        completeness_scores = []
        for name in dataset_names:
            missing_pct = datasets[name]['missing_data_analysis']['_summary']['overall_missing_percentage']
            completeness_scores.append(100 - missing_pct)
        
        fig.add_trace(
            go.Bar(
                x=dataset_names,
                y=completeness_scores,
                name='Completeness %',
                marker_color=['green' if score >= 95 else 'orange' if score >= 80 else 'red' for score in completeness_scores],
                hovertemplate='<b>%{x}</b><br>Completeness: %{y:.1f}%<extra></extra>'
            ),
            row=1, col=2
        )
        
        # 3. Statistical Test Results Heatmap
        if 'statistical_tests' in comparison_data:
            test_results = comparison_data['statistical_tests']
            if test_results:
                columns = list(test_results.keys())
                p_values = []
                significance = []
                
                for col in columns:
                    if 'means_comparison' in test_results[col]:
                        p_val = test_results[col]['means_comparison']['t_test_p_value']
                        p_values.append(p_val)
                        significance.append(1 if p_val < 0.05 else 0)
                
                if p_values:
                    fig.add_trace(
                        go.Heatmap(
                            z=[significance],
                            x=columns,
                            y=['Significant Difference'],
                            colorscale=[[0, 'green'], [1, 'red']],
                            showscale=True,
                            hovertemplate='<b>%{x}</b><br>Significant: %{z}<extra></extra>'
                        ),
                        row=2, col=1
                    )
        
        # 4. Missing Data Comparison Scatter
        missing_data_comparison = []
        columns_list = []
        
        for name in dataset_names:
            missing_data = datasets[name]['missing_data_analysis']
            for col, data in missing_data.items():
                if col != '_summary':
                    if col not in columns_list:
                        columns_list.append(col)
                        missing_data_comparison.append([data['missing_percentage']])
                    else:
                        idx = columns_list.index(col)
                        missing_data_comparison[idx].append(data['missing_percentage'])
        
        if missing_data_comparison:
            for i, name in enumerate(dataset_names):
                y_values = [row[i] if len(row) > i else 0 for row in missing_data_comparison]
                fig.add_trace(
                    go.Scatter(
                        x=columns_list[:10],  # Limit to first 10 columns
                        y=y_values[:10],
                        mode='markers+lines',
                        name=f'{name} Missing %',
                        marker=dict(size=8),
                        hovertemplate=f'<b>{name}</b><br>Column: %{{x}}<br>Missing: %{{y:.1f}}%<extra></extra>'
                    ),
                    row=2, col=2
                )
        
        fig.update_layout(
            title="🔄 Dataset Comparison Analysis",
            height=700,
            showlegend=True
        )
        
        return fig
    
    def export_dashboard_data(self, charts: Dict[str, go.Figure]) -> Dict[str, Any]:
        """Export dashboard data for external use"""
        export_data = {}
        
        for chart_name, fig in charts.items():
            export_data[chart_name] = {
                'figure_json': fig.to_json(),
                'data_summary': {
                    'traces_count': len(fig.data),
                    'chart_type': chart_name,
                    'has_data': len(fig.data) > 0
                }
            }
        
        return export_data
    
    def create_summary_metrics(self, analysis_results: Dict[str, Any]) -> Dict[str, Any]:
        """Create summary metrics for dashboard overview"""
        metrics = {}
        
        # Basic metrics
        if 'basic_info' in analysis_results:
            basic = analysis_results['basic_info']
            metrics['dataset_size'] = {
                'rows': basic['total_rows'],
                'columns': basic['total_columns'],
                'cells': basic['total_cells'],
                'memory_mb': basic['memory_usage'] / 1024 / 1024
            }
        
        # Quality metrics
        if 'missing_data_analysis' in analysis_results:
            missing = analysis_results['missing_data_analysis']['_summary']
            metrics['data_quality'] = {
                'completeness_percentage': 100 - missing['overall_missing_percentage'],
                'complete_rows': missing['complete_rows'],
                'missing_cells': missing['total_missing_cells']
            }
        
        # Statistical metrics
        if 'descriptive_stats' in analysis_results and 'message' not in analysis_results['descriptive_stats']:
            numeric_cols = len(analysis_results['descriptive_stats'])
            metrics['statistical_summary'] = {
                'numeric_columns': numeric_cols,
                'categorical_columns': len(analysis_results.get('categorical_analysis', {}))
            }
        
        # Outlier metrics
        if 'numerical_analysis' in analysis_results and 'message' not in analysis_results['numerical_analysis']:
            total_outliers = sum([data['outliers_count'] for data in analysis_results['numerical_analysis'].values()])
            metrics['outlier_summary'] = {
                'total_outliers': total_outliers,
                'columns_with_outliers': len([col for col, data in analysis_results['numerical_analysis'].items() 
                                            if data['outliers_count'] > 0])
            }
        
        return metrics