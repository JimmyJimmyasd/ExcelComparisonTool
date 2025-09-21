# Visualization Module for Statistical Analysis
# Provides interactive charts and plots using Plotly

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any
import streamlit as st
import base64
import io
from datetime import datetime

class StatisticalVisualizer:
    """
    Create interactive visualizations for statistical analysis
    """
    
    def __init__(self):
        self.color_palette = px.colors.qualitative.Set3
        self.sequential_colors = px.colors.sequential.Viridis
    
    def create_descriptive_stats_chart(self, desc_stats: Dict, title: str = "Descriptive Statistics") -> go.Figure:
        """Create a comprehensive descriptive statistics chart"""
        
        if not desc_stats or 'message' in desc_stats:
            return self._create_no_data_chart("No numerical data available for analysis")
        
        # Prepare data for visualization
        columns = list(desc_stats.keys())
        metrics = ['mean', 'median', 'std_dev', 'min', 'max']
        
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=('Central Tendency', 'Variability', 'Range', 'Distribution Shape'),
            specs=[[{"secondary_y": False}, {"secondary_y": False}],
                   [{"secondary_y": False}, {"secondary_y": False}]]
        )
        
        # Central Tendency (Mean vs Median)
        means = [desc_stats[col]['mean'] for col in columns]
        medians = [desc_stats[col]['median'] for col in columns]
        
        fig.add_trace(
            go.Bar(name='Mean', x=columns, y=means, marker_color='lightblue'),
            row=1, col=1
        )
        fig.add_trace(
            go.Bar(name='Median', x=columns, y=medians, marker_color='lightcoral'),
            row=1, col=1
        )
        
        # Variability (Standard Deviation)
        std_devs = [desc_stats[col]['std_dev'] for col in columns]
        fig.add_trace(
            go.Bar(name='Std Dev', x=columns, y=std_devs, marker_color='lightgreen'),
            row=1, col=2
        )
        
        # Range (Min to Max)
        mins = [desc_stats[col]['min'] for col in columns]
        maxs = [desc_stats[col]['max'] for col in columns]
        
        for i, col in enumerate(columns):
            fig.add_trace(
                go.Scatter(
                    x=[col, col], y=[mins[i], maxs[i]],
                    mode='lines+markers',
                    name=f'{col} Range',
                    line=dict(width=3),
                    showlegend=False
                ),
                row=2, col=1
            )
        
        # Distribution Shape (Skewness)
        skewness = [desc_stats[col]['skewness'] for col in columns]
        colors = ['red' if x > 0 else 'blue' for x in skewness]
        
        fig.add_trace(
            go.Bar(name='Skewness', x=columns, y=skewness, marker_color=colors),
            row=2, col=2
        )
        
        fig.update_layout(
            height=600,
            title_text=title,
            showlegend=True,
            title_x=0.5
        )
        
        return fig
    
    def create_missing_data_heatmap(self, df: pd.DataFrame, title: str = "Missing Data Pattern") -> go.Figure:
        """Create a heatmap showing missing data patterns"""
        
        # Create missing data matrix
        missing_matrix = df.isnull().astype(int)
        
        if missing_matrix.sum().sum() == 0:
            return self._create_no_data_chart("No missing data found - all cells are complete! ✅")
        
        fig = go.Figure(data=go.Heatmap(
            z=missing_matrix.values,
            x=missing_matrix.columns,
            y=list(range(len(missing_matrix))),
            colorscale=[[0, 'lightblue'], [1, 'red']],
            showscale=True,
            colorbar=dict(
                title="Missing Data",
                tickvals=[0, 1],
                ticktext=['Present', 'Missing']
            )
        ))
        
        fig.update_layout(
            title=title,
            xaxis_title="Columns",
            yaxis_title="Row Index",
            height=400,
            title_x=0.5
        )
        
        return fig
    
    def create_missing_data_summary_chart(self, missing_analysis: Dict) -> go.Figure:
        """Create a summary chart of missing data by column"""
        
        if '_summary' not in missing_analysis:
            return self._create_no_data_chart("No missing data analysis available")
        
        # Extract data (exclude summary)
        columns = [col for col in missing_analysis.keys() if col != '_summary']
        missing_percentages = [missing_analysis[col]['missing_percentage'] for col in columns]
        
        # Sort by missing percentage
        sorted_data = sorted(zip(columns, missing_percentages), key=lambda x: x[1], reverse=True)
        sorted_columns, sorted_percentages = zip(*sorted_data) if sorted_data else ([], [])
        
        # Color code based on severity
        colors = []
        for pct in sorted_percentages:
            if pct > 50:
                colors.append('red')
            elif pct > 20:
                colors.append('orange')
            elif pct > 5:
                colors.append('yellow')
            else:
                colors.append('green')
        
        fig = go.Figure(data=go.Bar(
            x=list(sorted_columns),
            y=list(sorted_percentages),
            marker_color=colors,
            text=[f'{pct:.1f}%' for pct in sorted_percentages],
            textposition='auto'
        ))
        
        fig.update_layout(
            title="Missing Data by Column",
            xaxis_title="Columns",
            yaxis_title="Missing Percentage (%)",
            height=400,
            title_x=0.5
        )
        
        # Add horizontal lines for severity levels
        fig.add_hline(y=50, line_dash="dash", line_color="red", 
                     annotation_text="High (>50%)", annotation_position="top right")
        fig.add_hline(y=20, line_dash="dash", line_color="orange",
                     annotation_text="Medium (>20%)", annotation_position="top right")
        fig.add_hline(y=5, line_dash="dash", line_color="yellow",
                     annotation_text="Low (>5%)", annotation_position="top right")
        
        return fig
    
    def create_correlation_heatmap(self, correlation_data: Dict, correlation_type: str = 'pearson') -> go.Figure:
        """Create correlation heatmap"""
        
        if 'message' in correlation_data:
            return self._create_no_data_chart(correlation_data['message'])
        
        corr_matrix = correlation_data['correlation_matrices'][correlation_type]
        
        fig = go.Figure(data=go.Heatmap(
            z=corr_matrix.values,
            x=corr_matrix.columns,
            y=corr_matrix.columns,
            colorscale='RdBu_r',
            zmid=0,
            text=np.round(corr_matrix.values, 2),
            texttemplate="%{text}",
            textfont={"size": 10},
            showscale=True,
            colorbar=dict(
                title="Correlation Coefficient"
            )
        ))
        
        fig.update_layout(
            title=f"Correlation Matrix ({correlation_type.title()})",
            xaxis_title="Variables",
            yaxis_title="Variables",
            height=500,
            title_x=0.5
        )
        
        return fig
    
    def create_distribution_plots(self, df: pd.DataFrame, columns: List[str] = None) -> go.Figure:
        """Create distribution plots for numerical columns"""
        
        numerical_cols = df.select_dtypes(include=[np.number]).columns
        if columns:
            numerical_cols = [col for col in columns if col in numerical_cols]
        
        if len(numerical_cols) == 0:
            return self._create_no_data_chart("No numerical columns available for distribution analysis")
        
        # Limit to first 6 columns for readability
        numerical_cols = numerical_cols[:6]
        
        # Calculate subplot layout
        n_cols = min(3, len(numerical_cols))
        n_rows = (len(numerical_cols) + n_cols - 1) // n_cols
        
        fig = make_subplots(
            rows=n_rows, cols=n_cols,
            subplot_titles=numerical_cols,
            vertical_spacing=0.1
        )
        
        for i, col in enumerate(numerical_cols):
            row = (i // n_cols) + 1
            col_pos = (i % n_cols) + 1
            
            # Clean data
            data = df[col].dropna()
            
            if len(data) > 0:
                fig.add_trace(
                    go.Histogram(
                        x=data,
                        name=col,
                        nbinsx=30,
                        opacity=0.7,
                        showlegend=False
                    ),
                    row=row, col=col_pos
                )
        
        fig.update_layout(
            height=300 * n_rows,
            title_text="Distribution Analysis",
            title_x=0.5,
            showlegend=False
        )
        
        return fig
    
    def create_box_plots(self, df: pd.DataFrame, columns: List[str] = None) -> go.Figure:
        """Create box plots for outlier detection"""
        
        numerical_cols = df.select_dtypes(include=[np.number]).columns
        if columns:
            numerical_cols = [col for col in columns if col in numerical_cols]
        
        if len(numerical_cols) == 0:
            return self._create_no_data_chart("No numerical columns available for box plot analysis")
        
        fig = go.Figure()
        
        for col in numerical_cols[:8]:  # Limit to 8 columns
            data = df[col].dropna()
            if len(data) > 0:
                fig.add_trace(go.Box(
                    y=data,
                    name=col,
                    boxpoints='outliers',
                    jitter=0.3,
                    pointpos=-1.8
                ))
        
        fig.update_layout(
            title="Box Plots - Outlier Detection",
            yaxis_title="Values",
            height=500,
            title_x=0.5
        )
        
        return fig
    
    def create_data_types_pie_chart(self, type_analysis: Dict) -> go.Figure:
        """Create pie chart showing data type distribution"""
        
        type_counts = {}
        for col, info in type_analysis.items():
            current_type = info['current_type']
            if current_type in type_counts:
                type_counts[current_type] += 1
            else:
                type_counts[current_type] = 1
        
        fig = go.Figure(data=[go.Pie(
            labels=list(type_counts.keys()),
            values=list(type_counts.values()),
            hole=0.3,
            textinfo='label+percent',
            textposition='auto'
        )])
        
        fig.update_layout(
            title="Data Types Distribution",
            height=400,
            title_x=0.5
        )
        
        return fig
    
    def create_summary_metrics_cards(self, basic_info: Dict, missing_summary: Dict) -> List[Dict]:
        """Create summary metric cards for dashboard"""
        
        cards = [
            {
                'title': 'Total Rows',
                'value': f"{basic_info['total_rows']:,}",
                'icon': '📊'
            },
            {
                'title': 'Total Columns', 
                'value': f"{basic_info['total_columns']:,}",
                'icon': '📋'
            },
            {
                'title': 'Missing Data',
                'value': f"{missing_summary['overall_missing_percentage']:.1f}%",
                'icon': '🕳️'
            },
            {
                'title': 'Complete Rows',
                'value': f"{missing_summary['complete_rows']:,}",
                'icon': '✅'
            },
            {
                'title': 'Memory Usage',
                'value': f"{basic_info['memory_usage'] / (1024*1024):.1f} MB",
                'icon': '💾'
            }
        ]
        
        return cards
    
    def _create_no_data_chart(self, message: str) -> go.Figure:
        """Create a placeholder chart when no data is available"""
        
        fig = go.Figure()
        
        fig.add_annotation(
            text=message,
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            xanchor='center', yanchor='middle',
            font=dict(size=16, color="gray"),
            showarrow=False
        )
        
        fig.update_layout(
            xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
            yaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
            height=300,
            plot_bgcolor='white'
        )
        
        return fig
    
    def create_comparative_distribution_plot(self, data_a: pd.Series, data_b: pd.Series, 
                                           title: str, label_a: str, label_b: str) -> go.Figure:
        """Create comparative distribution plot for two datasets"""
        
        fig = go.Figure()
        
        # Add histograms for both datasets
        fig.add_trace(go.Histogram(
            x=data_a.dropna(),
            name=label_a,
            opacity=0.7,
            nbinsx=30,
            marker_color='blue'
        ))
        
        fig.add_trace(go.Histogram(
            x=data_b.dropna(),
            name=label_b,
            opacity=0.7,
            nbinsx=30,
            marker_color='red'
        ))
        
        fig.update_layout(
            title=title,
            xaxis_title="Value",
            yaxis_title="Frequency",
            barmode='overlay',
            height=500,
            title_x=0.5,
            legend=dict(x=0.7, y=0.9)
        )
        
        return fig
    
    def create_comparative_missing_data_chart(self, missing_a: Dict, missing_b: Dict, 
                                            label_a: str, label_b: str) -> go.Figure:
        """Create comparative missing data chart"""
        
        # Extract column-wise missing percentages
        columns_a = list(missing_a.get('column_missing_percentage', {}).keys())
        missing_pct_a = list(missing_a.get('column_missing_percentage', {}).values())
        
        columns_b = list(missing_b.get('column_missing_percentage', {}).keys())
        missing_pct_b = list(missing_b.get('column_missing_percentage', {}).values())
        
        # Get all unique columns
        all_columns = sorted(set(columns_a + columns_b))
        
        # Align data for both datasets
        aligned_a = []
        aligned_b = []
        
        for col in all_columns:
            if col in columns_a:
                aligned_a.append(missing_a['column_missing_percentage'][col])
            else:
                aligned_a.append(0)
                
            if col in columns_b:
                aligned_b.append(missing_b['column_missing_percentage'][col])
            else:
                aligned_b.append(0)
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=all_columns,
            y=aligned_a,
            name=label_a,
            marker_color='blue',
            opacity=0.7
        ))
        
        fig.add_trace(go.Bar(
            x=all_columns,
            y=aligned_b,
            name=label_b,
            marker_color='red',
            opacity=0.7
        ))
        
        fig.update_layout(
            title="Missing Data Comparison by Column",
            xaxis_title="Columns",
            yaxis_title="Missing Data (%)",
            barmode='group',
            height=500,
            title_x=0.5,
            xaxis_tickangle=-45
        )
        
        return fig
    
    def create_comparative_histogram(self, data_a: pd.Series, data_b: pd.Series,
                                   column_name: str, label_a: str, label_b: str) -> go.Figure:
        """Create comparative histogram for a specific column"""
        
        fig = go.Figure()
        
        # Calculate appropriate bins
        combined_data = pd.concat([data_a.dropna(), data_b.dropna()])
        bin_edges = np.histogram_bin_edges(combined_data, bins=30)
        
        fig.add_trace(go.Histogram(
            x=data_a.dropna(),
            name=label_a,
            opacity=0.6,
            xbins=dict(start=bin_edges[0], end=bin_edges[-1], size=(bin_edges[-1] - bin_edges[0])/30),
            marker_color='blue',
            legendgroup='group1'
        ))
        
        fig.add_trace(go.Histogram(
            x=data_b.dropna(),
            name=label_b,
            opacity=0.6,
            xbins=dict(start=bin_edges[0], end=bin_edges[-1], size=(bin_edges[-1] - bin_edges[0])/30),
            marker_color='red',
            legendgroup='group2'
        ))
        
        # Add mean lines
        fig.add_vline(
            x=data_a.mean(),
            line_dash="dash",
            line_color="blue",
            annotation_text=f"{label_a} Mean: {data_a.mean():.2f}"
        )
        
        fig.add_vline(
            x=data_b.mean(),
            line_dash="dash", 
            line_color="red",
            annotation_text=f"{label_b} Mean: {data_b.mean():.2f}"
        )
        
        fig.update_layout(
            title=f"{column_name} - Distribution Comparison",
            xaxis_title=column_name,
            yaxis_title="Frequency",
            barmode='overlay',
            height=500,
            title_x=0.5,
            legend=dict(x=0.7, y=0.9)
        )
        
        return fig
    
    def create_missing_data_chart(self, missing_data_analysis: dict) -> go.Figure:
        """Create missing data visualization chart"""
        
        column_missing = missing_data_analysis['column_missing_percentage']
        
        # Filter columns with missing data
        cols_with_missing = {k: v for k, v in column_missing.items() if v > 0}
        
        if not cols_with_missing:
            return self._create_no_data_chart("No missing data found - excellent data quality!")
        
        fig = go.Figure()
        
        # Sort by missing percentage
        sorted_missing = dict(sorted(cols_with_missing.items(), key=lambda x: x[1], reverse=True))
        
        colors = ['red' if v > 50 else 'orange' if v > 20 else 'yellow' for v in sorted_missing.values()]
        
        fig.add_trace(go.Bar(
            x=list(sorted_missing.keys()),
            y=list(sorted_missing.values()),
            marker_color=colors,
            text=[f"{v:.1f}%" for v in sorted_missing.values()],
            textposition='auto'
        ))
        
        fig.update_layout(
            title="Missing Data Analysis by Column",
            xaxis_title="Columns",
            yaxis_title="Missing Data Percentage",
            height=400,
            xaxis={'tickangle': 45}
        )
        
        return fig
    
    def create_data_completeness_gauge(self, completeness_percentage: float) -> go.Figure:
        """Create data completeness gauge chart"""
        
        # Determine color based on completeness
        if completeness_percentage >= 95:
            color = "green"
        elif completeness_percentage >= 80:
            color = "yellow"
        else:
            color = "red"
        
        fig = go.Figure(go.Indicator(
            mode = "gauge+number+delta",
            value = completeness_percentage,
            domain = {'x': [0, 1], 'y': [0, 1]},
            title = {'text': "Data Completeness"},
            delta = {'reference': 100, 'increasing': {'color': "green"}},
            gauge = {
                'axis': {'range': [None, 100]},
                'bar': {'color': color},
                'steps': [
                    {'range': [0, 50], 'color': "lightgray"},
                    {'range': [50, 80], 'color': "gray"},
                    {'range': [80, 95], 'color': "lightgreen"},
                    {'range': [95, 100], 'color': "green"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 90
                }
            }
        ))
        
        fig.update_layout(height=300, font={'color': "darkblue", 'family': "Arial"})
        
        return fig
    
    def create_duplicate_analysis_chart(self, duplicate_analysis: dict) -> go.Figure:
        """Create duplicate analysis chart"""
        
        exact_count = duplicate_analysis.get('exact_duplicates', {}).get('count', 0)
        near_count = duplicate_analysis.get('near_duplicates', {}).get('total_found', 0)
        
        if exact_count == 0 and near_count == 0:
            return self._create_no_data_chart("No duplicates found - excellent data quality!")
        
        fig = go.Figure()
        
        categories = ['Exact Duplicates', 'Near Duplicates']
        values = [exact_count, near_count]
        colors = ['red', 'orange']
        
        fig.add_trace(go.Bar(
            x=categories,
            y=values,
            marker_color=colors,
            text=[str(v) for v in values],
            textposition='auto'
        ))
        
        fig.update_layout(
            title="Duplicate Records Analysis",
            xaxis_title="Duplicate Type",
            yaxis_title="Number of Records",
            height=400
        )
        
        return fig
    
    def create_data_type_issues_chart(self, data_type_analysis: dict) -> go.Figure:
        """Create data type issues chart"""
        
        # Count issues by type
        issue_counts = {'Mixed Types': 0, 'Invalid Values': 0, 'Conversion Issues': 0}
        
        for column, analysis in data_type_analysis.items():
            if analysis.get('has_mixed_types', False):
                issue_counts['Mixed Types'] += 1
            if analysis.get('conversion_issues', 0) > 0:
                issue_counts['Conversion Issues'] += 1
            if analysis.get('invalid_values', 0) > 0:
                issue_counts['Invalid Values'] += 1
        
        # Filter out zero counts
        issue_counts = {k: v for k, v in issue_counts.items() if v > 0}
        
        if not issue_counts:
            return self._create_no_data_chart("No data type issues found - excellent data quality!")
        
        fig = go.Figure()
        
        colors = ['red', 'orange', 'yellow'][:len(issue_counts)]
        
        fig.add_trace(go.Bar(
            x=list(issue_counts.keys()),
            y=list(issue_counts.values()),
            marker_color=colors,
            text=[str(v) for v in issue_counts.values()],
            textposition='auto'
        ))
        
        fig.update_layout(
            title="Data Type Issues by Category",
            xaxis_title="Issue Type",
            yaxis_title="Number of Columns Affected",
            height=400
        )
        
        return fig
    
    # ============================================================================
    # BUSINESS INTELLIGENCE VISUALIZATIONS
    # ============================================================================
    
    def create_revenue_trend_chart(self, sales_data: Dict, amount_column: str) -> go.Figure:
        """Create revenue trend chart over time"""
        
        if 'time_trends' not in sales_data or not sales_data['time_trends']:
            return self._create_no_data_chart("No time trend data available")
        
        fig = go.Figure()
        
        for date_col, trend_data in sales_data['time_trends'].items():
            if 'monthly_revenue' in trend_data:
                monthly_data = trend_data['monthly_revenue']
                
                # Convert period index to strings for plotting
                periods = list(monthly_data.keys())
                revenues = list(monthly_data.values())
                
                fig.add_trace(go.Scatter(
                    x=[str(p) for p in periods],
                    y=revenues,
                    mode='lines+markers',
                    name=f'Revenue Trend ({date_col})',
                    line=dict(width=3),
                    marker=dict(size=8)
                ))
        
        fig.update_layout(
            title=f"Revenue Trends - {amount_column}",
            xaxis_title="Month",
            yaxis_title="Revenue ($)",
            height=500,
            hovermode='x unified',
            title_x=0.5
        )
        
        return fig
    
    def create_category_performance_chart(self, sales_data: Dict, amount_column: str) -> go.Figure:
        """Create category performance comparison chart"""
        
        if 'category_performance' not in sales_data:
            return self._create_no_data_chart("No category performance data available")
        
        # Use the first category for visualization
        category_data = list(sales_data['category_performance'].values())[0]
        
        if 'revenue_by_category' not in category_data:
            return self._create_no_data_chart("No revenue by category data available")
        
        revenue_data = category_data['revenue_by_category']
        transaction_data = category_data.get('transaction_count', {})
        
        # Sort by revenue
        sorted_categories = sorted(revenue_data.items(), key=lambda x: x[1], reverse=True)
        categories, revenues = zip(*sorted_categories) if sorted_categories else ([], [])
        
        fig = make_subplots(
            rows=1, cols=2,
            subplot_titles=['Revenue by Category', 'Transaction Volume by Category'],
            specs=[[{"secondary_y": False}, {"secondary_y": False}]]
        )
        
        # Revenue chart
        fig.add_trace(
            go.Bar(
                x=list(categories),
                y=list(revenues),
                name='Revenue',
                marker_color='blue',
                text=[f'${v:,.0f}' for v in revenues],
                textposition='auto'
            ),
            row=1, col=1
        )
        
        # Transaction volume chart
        transaction_counts = [transaction_data.get(cat, 0) for cat in categories]
        fig.add_trace(
            go.Bar(
                x=list(categories),
                y=transaction_counts,
                name='Transactions',
                marker_color='green',
                text=[str(v) for v in transaction_counts],
                textposition='auto'
            ),
            row=1, col=2
        )
        
        fig.update_layout(
            title=f"Category Performance Analysis - {amount_column}",
            height=500,
            showlegend=False,
            title_x=0.5
        )
        
        fig.update_xaxes(tickangle=45)
        
        return fig
    
    def create_customer_segmentation_chart(self, customer_data: Dict) -> go.Figure:
        """Create customer segmentation visualization"""
        
        if 'customer_segmentation' not in customer_data:
            return self._create_no_data_chart("No customer segmentation data available")
        
        segmentation = customer_data['customer_segmentation']
        
        fig = make_subplots(
            rows=1, cols=2,
            subplot_titles=['Value Segments', 'Frequency Segments'],
            specs=[[{"type": "domain"}, {"type": "domain"}]]
        )
        
        # Value segments pie chart
        if 'value_segments' in segmentation:
            value_seg = segmentation['value_segments']
            fig.add_trace(
                go.Pie(
                    labels=['High Value', 'Medium Value', 'Low Value'],
                    values=[
                        value_seg.get('high_value', 0),
                        value_seg.get('medium_value', 0),
                        value_seg.get('low_value', 0)
                    ],
                    name="Value Segments",
                    marker_colors=['gold', 'lightblue', 'lightgray']
                ),
                row=1, col=1
            )
        
        # Frequency segments pie chart
        if 'frequency_segments' in segmentation:
            freq_seg = segmentation['frequency_segments']
            fig.add_trace(
                go.Pie(
                    labels=['Repeat Customers', 'One-time Customers'],
                    values=[
                        freq_seg.get('repeat_customers', 0),
                        freq_seg.get('one_time_customers', 0)
                    ],
                    name="Frequency Segments",
                    marker_colors=['darkgreen', 'lightgreen']
                ),
                row=1, col=2
            )
        
        fig.update_layout(
            title="Customer Segmentation Analysis",
            height=500,
            title_x=0.5
        )
        
        return fig
    
    def create_product_performance_chart(self, product_data: Dict) -> go.Figure:
        """Create product performance analysis chart"""
        
        if not product_data or 'message' in product_data:
            return self._create_no_data_chart("No product performance data available")
        
        # Get the first product analysis
        product_analysis = list(product_data.values())[0]
        
        if 'product_performance' not in product_analysis:
            return self._create_no_data_chart("No product performance metrics available")
        
        # Get the first amount column's performance data
        performance_data = list(product_analysis['product_performance'].values())[0]
        
        revenue_data = performance_data.get('revenue_by_product', {})
        volume_data = performance_data.get('sales_volume_by_product', {})
        
        # Sort by revenue
        sorted_products = sorted(revenue_data.items(), key=lambda x: x[1], reverse=True)[:10]  # Top 10
        products, revenues = zip(*sorted_products) if sorted_products else ([], [])
        
        fig = make_subplots(
            rows=2, cols=1,
            subplot_titles=['Top Products by Revenue', 'Top Products by Sales Volume'],
            vertical_spacing=0.15
        )
        
        # Revenue chart
        fig.add_trace(
            go.Bar(
                x=list(products),
                y=list(revenues),
                name='Revenue',
                marker_color='blue',
                text=[f'${v:,.0f}' for v in revenues],
                textposition='auto'
            ),
            row=1, col=1
        )
        
        # Volume chart
        volumes = [volume_data.get(prod, 0) for prod in products]
        fig.add_trace(
            go.Bar(
                x=list(products),
                y=volumes,
                name='Volume',
                marker_color='orange',
                text=[str(v) for v in volumes],
                textposition='auto'
            ),
            row=2, col=1
        )
        
        fig.update_layout(
            title="Product Performance Analysis",
            height=700,
            showlegend=False,
            title_x=0.5
        )
        
        fig.update_xaxes(tickangle=45)
        
        return fig
    
    def create_satisfaction_analysis_chart(self, customer_data: Dict) -> go.Figure:
        """Create customer satisfaction analysis chart"""
        
        if 'satisfaction_analysis' not in customer_data:
            return self._create_no_data_chart("No satisfaction data available")
        
        satisfaction = customer_data['satisfaction_analysis']
        
        if not satisfaction or 'message' in satisfaction:
            return self._create_no_data_chart("No rating columns detected for satisfaction analysis")
        
        # Get the first rating column's data
        rating_data = list(satisfaction.values())[0]
        
        fig = make_subplots(
            rows=1, cols=2,
            subplot_titles=['Rating Distribution', 'Satisfaction Level'],
            specs=[[{"type": "xy"}, {"type": "indicator"}]]
        )
        
        # Rating distribution histogram
        if 'rating_distribution' in rating_data:
            ratings = rating_data['rating_distribution']
            fig.add_trace(
                go.Bar(
                    x=list(ratings.keys()),
                    y=list(ratings.values()),
                    name='Rating Distribution',
                    marker_color='skyblue',
                    text=[str(v) for v in ratings.values()],
                    textposition='auto'
                ),
                row=1, col=1
            )
        
        # Satisfaction gauge
        avg_rating = rating_data.get('average_rating', 0)
        satisfaction_level = rating_data.get('satisfaction_level', 'Unknown')
        
        fig.add_trace(
            go.Indicator(
                mode="gauge+number+delta",
                value=avg_rating,
                domain={'x': [0.6, 1], 'y': [0, 1]},
                title={'text': f"Average Rating<br>{satisfaction_level}"},
                gauge={
                    'axis': {'range': [None, 5]},
                    'bar': {'color': "darkgreen" if avg_rating > 4 else "orange" if avg_rating > 3 else "red"},
                    'steps': [
                        {'range': [0, 2], 'color': "lightgray"},
                        {'range': [2, 3], 'color': "yellow"},
                        {'range': [3, 4], 'color': "lightgreen"},
                        {'range': [4, 5], 'color': "green"}
                    ],
                    'threshold': {
                        'line': {'color': "red", 'width': 4},
                        'thickness': 0.75,
                        'value': 4.0
                    }
                }
            ),
            row=1, col=2
        )
        
        fig.update_layout(
            title="Customer Satisfaction Analysis",
            height=500,
            title_x=0.5
        )
        
        return fig
    
    def create_business_kpi_cards(self, business_overview: Dict, sales_data: Dict) -> List[Dict]:
        """Create business KPI cards for dashboard"""
        
        cards = []
        
        # Dataset overview cards
        dataset_info = business_overview.get('dataset_info', {})
        cards.append({
            'title': 'Total Records',
            'value': f"{dataset_info.get('total_records', 0):,}",
            'icon': '📊'
        })
        
        # Revenue cards
        if sales_data:
            amount_col = list(sales_data.keys())[0]
            sales_info = sales_data[amount_col]
            
            cards.extend([
                {
                    'title': 'Total Revenue',
                    'value': f"${sales_info.get('total_revenue', 0):,.2f}",
                    'icon': '💰'
                },
                {
                    'title': 'Average Transaction',
                    'value': f"${sales_info.get('average_transaction', 0):,.2f}",
                    'icon': '🧾'
                },
                {
                    'title': 'Total Transactions',
                    'value': f"{sales_info.get('performance_metrics', {}).get('total_transactions', 0):,}",
                    'icon': '🛒'
                }
            ])
        
        # Time range card
        date_range = dataset_info.get('date_range', {})
        if isinstance(date_range, dict) and date_range:
            first_date_col = list(date_range.keys())[0]
            duration = date_range[first_date_col].get('duration_days', 0)
            cards.append({
                'title': 'Data Duration',
                'value': f"{duration} days",
                'icon': '📅'
            })
        
        # Business dimensions card
        column_class = business_overview.get('column_classification', {})
        business_dims = len(column_class.get('category_columns', [])) + len(column_class.get('amount_columns', []))
        cards.append({
            'title': 'Business Dimensions',
            'value': str(business_dims),
            'icon': '🎯'
        })
        
        return cards
    
    def create_revenue_distribution_chart(self, sales_data: Dict, amount_column: str) -> go.Figure:
        """Create revenue distribution analysis chart"""
        
        if 'revenue_distribution' not in sales_data:
            return self._create_no_data_chart("No revenue distribution data available")
        
        revenue_dist = sales_data['revenue_distribution']
        
        fig = make_subplots(
            rows=1, cols=2,
            subplot_titles=['Revenue Quartiles', 'Revenue Percentiles'],
            specs=[[{"type": "bar"}, {"type": "bar"}]]
        )
        
        # Quartiles chart
        if 'quartiles' in revenue_dist:
            quartiles = revenue_dist['quartiles']
            fig.add_trace(
                go.Bar(
                    x=['Q1', 'Q2 (Median)', 'Q3', 'Q4 (Max)'],
                    y=[quartiles.get('Q1', 0), quartiles.get('Q2_median', 0), 
                       quartiles.get('Q3', 0), quartiles.get('Q4_max', 0)],
                    name='Quartiles',
                    marker_color=['lightblue', 'blue', 'darkblue', 'navy'],
                    text=[f'${v:,.0f}' for v in [quartiles.get('Q1', 0), quartiles.get('Q2_median', 0), 
                                                  quartiles.get('Q3', 0), quartiles.get('Q4_max', 0)]],
                    textposition='auto'
                ),
                row=1, col=1
            )
        
        # Percentiles chart
        if 'percentiles' in revenue_dist:
            percentiles = revenue_dist['percentiles']
            fig.add_trace(
                go.Bar(
                    x=['P90', 'P95', 'P99'],
                    y=[percentiles.get('P90', 0), percentiles.get('P95', 0), percentiles.get('P99', 0)],
                    name='Percentiles',
                    marker_color=['orange', 'red', 'darkred'],
                    text=[f'${v:,.0f}' for v in [percentiles.get('P90', 0), percentiles.get('P95', 0), percentiles.get('P99', 0)]],
                    textposition='auto'
                ),
                row=1, col=2
            )
        
        fig.update_layout(
            title=f"Revenue Distribution Analysis - {amount_column}",
            height=500,
            showlegend=False,
            title_x=0.5
        )
        
        return fig
    
    def create_comparative_business_chart(self, data_a: Dict, data_b: Dict, 
                                        label_a: str, label_b: str, metric: str) -> go.Figure:
        """Create comparative business metrics chart"""
        
        if metric == 'revenue_trends' and 'time_trends' in data_a and 'time_trends' in data_b:
            return self._create_comparative_revenue_trends(data_a, data_b, label_a, label_b)
        elif metric == 'category_performance' and 'category_performance' in data_a and 'category_performance' in data_b:
            return self._create_comparative_category_performance(data_a, data_b, label_a, label_b)
        else:
            return self._create_no_data_chart(f"Comparative {metric} analysis not available")
    
    def _create_comparative_revenue_trends(self, data_a: Dict, data_b: Dict, 
                                         label_a: str, label_b: str) -> go.Figure:
        """Create comparative revenue trends chart"""
        
        fig = go.Figure()
        
        # Extract revenue trends from both datasets
        trends_a = list(data_a['time_trends'].values())[0].get('monthly_revenue', {})
        trends_b = list(data_b['time_trends'].values())[0].get('monthly_revenue', {})
        
        # Plot trends for dataset A
        if trends_a:
            periods_a = [str(p) for p in trends_a.keys()]
            revenues_a = list(trends_a.values())
            
            fig.add_trace(go.Scatter(
                x=periods_a,
                y=revenues_a,
                mode='lines+markers',
                name=label_a,
                line=dict(width=3, color='blue'),
                marker=dict(size=8)
            ))
        
        # Plot trends for dataset B
        if trends_b:
            periods_b = [str(p) for p in trends_b.keys()]
            revenues_b = list(trends_b.values())
            
            fig.add_trace(go.Scatter(
                x=periods_b,
                y=revenues_b,
                mode='lines+markers',
                name=label_b,
                line=dict(width=3, color='red'),
                marker=dict(size=8)
            ))
        
        fig.update_layout(
            title="Comparative Revenue Trends",
            xaxis_title="Month",
            yaxis_title="Revenue ($)",
            height=500,
            hovermode='x unified',
            title_x=0.5,
            legend=dict(x=0.7, y=0.9)
        )
        
        return fig
    
    def _create_comparative_category_performance(self, data_a: Dict, data_b: Dict, 
                                               label_a: str, label_b: str) -> go.Figure:
        """Create comparative category performance chart"""
        
        # Extract category performance data
        cat_data_a = list(data_a['category_performance'].values())[0]['revenue_by_category']
        cat_data_b = list(data_b['category_performance'].values())[0]['revenue_by_category']
        
        # Get all unique categories
        all_categories = sorted(set(list(cat_data_a.keys()) + list(cat_data_b.keys())))
        
        # Align data for both datasets
        revenues_a = [cat_data_a.get(cat, 0) for cat in all_categories]
        revenues_b = [cat_data_b.get(cat, 0) for cat in all_categories]
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=all_categories,
            y=revenues_a,
            name=label_a,
            marker_color='blue',
            opacity=0.7
        ))
        
        fig.add_trace(go.Bar(
            x=all_categories,
            y=revenues_b,
            name=label_b,
            marker_color='red',
            opacity=0.7
        ))
        
        fig.update_layout(
            title="Comparative Category Performance",
            xaxis_title="Categories",
            yaxis_title="Revenue ($)",
            barmode='group',
            height=500,
            title_x=0.5,
            xaxis_tickangle=-45
        )
        
        return fig
    
    def create_financial_ratios_dashboard(self, financial_ratios: Dict) -> go.Figure:
        """Create comprehensive financial ratios dashboard with gauge charts"""
        
        fig = make_subplots(
            rows=2, cols=3,
            subplot_titles=('Current Ratio', 'ROE (%)', 'Net Profit Margin (%)', 
                          'Quick Ratio', 'ROA (%)', 'Gross Profit Margin (%)'),
            specs=[[{'type': 'indicator'}, {'type': 'indicator'}, {'type': 'indicator'}],
                   [{'type': 'indicator'}, {'type': 'indicator'}, {'type': 'indicator'}]]
        )
        
        # Extract ratio values
        liquidity = financial_ratios.get('liquidity_ratios', {})
        profitability = financial_ratios.get('profitability_ratios', {})
        
        # Current Ratio Gauge
        current_ratio = liquidity.get('current_ratio', {}).get('value', 0)
        fig.add_trace(go.Indicator(
            mode="gauge+number+delta",
            value=current_ratio,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Current Ratio"},
            gauge={
                'axis': {'range': [None, 3]},
                'bar': {'color': "darkblue"},
                'steps': [
                    {'range': [0, 1], 'color': "lightgray"},
                    {'range': [1, 1.5], 'color': "yellow"},
                    {'range': [1.5, 3], 'color': "lightgreen"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 1.5
                }
            }
        ), row=1, col=1)
        
        # ROE Gauge
        roe = profitability.get('roe', {}).get('value', 0)
        fig.add_trace(go.Indicator(
            mode="gauge+number",
            value=roe,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "ROE (%)"},
            gauge={
                'axis': {'range': [None, 30]},
                'bar': {'color': "darkgreen"},
                'steps': [
                    {'range': [0, 10], 'color': "lightgray"},
                    {'range': [10, 15], 'color': "yellow"},
                    {'range': [15, 30], 'color': "lightgreen"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 15
                }
            }
        ), row=1, col=2)
        
        # Net Profit Margin Gauge
        net_margin = profitability.get('net_profit_margin', {}).get('value', 0)
        fig.add_trace(go.Indicator(
            mode="gauge+number",
            value=net_margin,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Net Margin (%)"},
            gauge={
                'axis': {'range': [None, 50]},
                'bar': {'color': "purple"},
                'steps': [
                    {'range': [0, 10], 'color': "lightgray"},
                    {'range': [10, 20], 'color': "yellow"},
                    {'range': [20, 50], 'color': "lightgreen"}
                ]
            }
        ), row=1, col=3)
        
        # Quick Ratio Gauge
        quick_ratio = liquidity.get('quick_ratio', {}).get('value', 0)
        fig.add_trace(go.Indicator(
            mode="gauge+number",
            value=quick_ratio,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Quick Ratio"},
            gauge={
                'axis': {'range': [None, 2]},
                'bar': {'color': "orange"},
                'steps': [
                    {'range': [0, 0.5], 'color': "lightgray"},
                    {'range': [0.5, 1], 'color': "yellow"},
                    {'range': [1, 2], 'color': "lightgreen"}
                ]
            }
        ), row=2, col=1)
        
        # ROA Gauge
        roa = profitability.get('roa', {}).get('value', 0)
        fig.add_trace(go.Indicator(
            mode="gauge+number",
            value=roa,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "ROA (%)"},
            gauge={
                'axis': {'range': [None, 15]},
                'bar': {'color': "teal"},
                'steps': [
                    {'range': [0, 2], 'color': "lightgray"},
                    {'range': [2, 5], 'color': "yellow"},
                    {'range': [5, 15], 'color': "lightgreen"}
                ]
            }
        ), row=2, col=2)
        
        # Gross Profit Margin Gauge
        gross_margin = profitability.get('gross_profit_margin', {}).get('value', 0)
        fig.add_trace(go.Indicator(
            mode="gauge+number",
            value=gross_margin,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "Gross Margin (%)"},
            gauge={
                'axis': {'range': [None, 80]},
                'bar': {'color': "crimson"},
                'steps': [
                    {'range': [0, 25], 'color': "lightgray"},
                    {'range': [25, 40], 'color': "yellow"},
                    {'range': [40, 80], 'color': "lightgreen"}
                ]
            }
        ), row=2, col=3)
        
        fig.update_layout(
            title="Financial Ratios Dashboard",
            height=600,
            title_x=0.5
        )
        
        return fig
    
    def create_kpi_scorecard(self, business_kpis: Dict) -> go.Figure:
        """Create KPI scorecard with progress bars and alerts"""
        
        fig = go.Figure()
        
        kpi_data = []
        colors = []
        
        # Extract KPIs from different categories
        categories = ['sales_marketing_kpis', 'operational_kpis', 'hr_kpis', 'banking_cash_kpis']
        
        for category in categories:
            if category in business_kpis:
                category_data = business_kpis[category]
                for kpi_name, kpi_info in category_data.items():
                    if isinstance(kpi_info, dict) and 'value' in kpi_info:
                        kpi_data.append({
                            'name': kpi_name.replace('_', ' ').title(),
                            'value': kpi_info['value'],
                            'interpretation': kpi_info.get('interpretation', 'Unknown'),
                            'unit': kpi_info.get('unit', '')
                        })
                        
                        # Color based on interpretation
                        interp = kpi_info.get('interpretation', '').lower()
                        if 'excellent' in interp or 'good' in interp or 'strong' in interp:
                            colors.append('green')
                        elif 'concerning' in interp or 'poor' in interp or 'critical' in interp:
                            colors.append('red')
                        else:
                            colors.append('orange')
        
        if kpi_data:
            names = [item['name'] for item in kpi_data]
            values = [item['value'] for item in kpi_data]
            interpretations = [item['interpretation'] for item in kpi_data]
            
            fig.add_trace(go.Bar(
                y=names,
                x=values,
                orientation='h',
                marker_color=colors,
                text=[f"{val}{item['unit']} - {interp}" for val, item, interp in zip(values, kpi_data, interpretations)],
                textposition='auto'
            ))
        
        fig.update_layout(
            title="Business KPIs Scorecard",
            xaxis_title="KPI Values",
            height=max(400, len(kpi_data) * 40),
            title_x=0.5,
            showlegend=False
        )
        
        return fig
    
    def create_early_warning_alerts(self, alerts_data: Dict) -> go.Figure:
        """Create early warning indicators dashboard"""
        
        fig = go.Figure()
        
        # Extract alerts
        critical_alerts = alerts_data.get('critical_alerts', [])
        warning_alerts = alerts_data.get('warning_alerts', [])
        
        all_alerts = []
        colors = []
        
        # Add critical alerts
        for alert in critical_alerts:
            all_alerts.append(f"🚨 {alert.get('type', 'Critical Alert')}")
            colors.append('red')
        
        # Add warning alerts
        for alert in warning_alerts:
            all_alerts.append(f"⚠️ {alert.get('type', 'Warning Alert')}")
            colors.append('orange')
        
        # Add some dummy positive indicators if no alerts
        if not all_alerts:
            all_alerts = ['✅ No Critical Issues', '✅ Systems Normal', '✅ Performance Good']
            colors = ['green', 'green', 'green']
        
        # Create alert status chart
        fig.add_trace(go.Bar(
            y=all_alerts,
            x=[1] * len(all_alerts),
            orientation='h',
            marker_color=colors,
            text=all_alerts,
            textposition='outside',
            showlegend=False
        ))
        
        fig.update_layout(
            title="Early Warning System - Alert Status",
            xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
            yaxis=dict(showgrid=False),
            height=max(300, len(all_alerts) * 50),
            title_x=0.5,
            plot_bgcolor='white'
        )
        
        return fig
    
    def create_sector_performance_heatmap(self, sector_data: Dict) -> go.Figure:
        """Create sector performance heatmap"""
        
        sector_rankings = sector_data.get('sector_rankings', [])
        
        if not sector_rankings:
            # Create empty heatmap
            fig = go.Figure(data=go.Heatmap(
                z=[[0]],
                x=['No Data'],
                y=['No Sectors'],
                colorscale='RdYlGn',
                showscale=True
            ))
        else:
            sectors = [item['sector'] for item in sector_rankings]
            revenues = [item['revenue'] for item in sector_rankings]
            market_shares = [item['market_share'] for item in sector_rankings]
            
            # Create heatmap matrix
            heatmap_data = []
            metrics = ['Revenue', 'Market Share', 'Avg Transaction']
            
            for sector_info in sector_rankings:
                row = [
                    sector_info['revenue'],
                    sector_info['market_share'],
                    sector_info['avg_transaction']
                ]
                heatmap_data.append(row)
            
            fig = go.Figure(data=go.Heatmap(
                z=heatmap_data,
                x=metrics,
                y=sectors,
                colorscale='RdYlGn',
                showscale=True,
                text=[[f"${val:,.0f}" if i==0 else f"{val:.1f}%" if i==1 else f"${val:,.0f}" 
                       for i, val in enumerate(row)] for row in heatmap_data],
                texttemplate="%{text}",
                textfont={"size": 10}
            ))
        
        fig.update_layout(
            title="Sector Performance Heatmap",
            height=max(400, len(sector_rankings) * 40),
            title_x=0.5
        )
        
        return fig
    
    def create_risk_assessment_radar(self, risk_data: Dict) -> go.Figure:
        """Create risk assessment radar chart"""
        
        risk_factors = risk_data.get('risk_factors', [])
        
        if not risk_factors:
            # Create empty radar
            categories = ['No Risk Factors Identified']
            values = [0]
        else:
            categories = [factor.get('type', 'Unknown Risk') for factor in risk_factors]
            # Convert severity to numeric values
            severity_map = {'Low': 1, 'Medium': 2, 'High': 3}
            values = [severity_map.get(factor.get('severity', 'Low'), 1) for factor in risk_factors]
        
        fig = go.Figure()
        
        fig.add_trace(go.Scatterpolar(
            r=values,
            theta=categories,
            fill='toself',
            name='Risk Level',
            marker_color='red',
            opacity=0.6
        ))
        
        fig.update_layout(
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[0, 3],
                    ticktext=['None', 'Low', 'Medium', 'High'],
                    tickvals=[0, 1, 2, 3]
                )
            ),
            title="Business Risk Assessment",
            height=500,
            title_x=0.5
        )
        
        return fig
    
    def create_data_quality_progress_bar(self, governance_data: Dict) -> go.Figure:
        """Create data quality progress bar"""
        
        fig = go.Figure()
        
        data_quality = governance_data.get('data_quality_score', {})
        score = data_quality.get('score', 0)
        
        # Create progress bar
        fig.add_trace(go.Bar(
            x=[score],
            y=['Data Quality Score'],
            orientation='h',
            marker=dict(
                color='green' if score >= 95 else 'orange' if score >= 85 else 'red',
                line=dict(color='black', width=1)
            ),
            text=f"{score:.1f}%",
            textposition='outside'
        ))
        
        # Add benchmark lines
        fig.add_vline(x=95, line_dash="dash", line_color="green", 
                      annotation_text="Excellent (95%)", annotation_position="top")
        fig.add_vline(x=85, line_dash="dash", line_color="orange", 
                      annotation_text="Good (85%)", annotation_position="top")
        
        fig.update_layout(
            title="Data Quality Score",
            xaxis=dict(range=[0, 100], title="Score (%)"),
            height=200,
            title_x=0.5,
            showlegend=False
        )
        
        return fig


class InteractiveDashboard:
    """
    Enhanced Interactive Dashboard for Business Intelligence Analysis
    Provides real-time charts with drill-down capabilities and export options
    """
    
    def __init__(self, df_a: pd.DataFrame, df_b: pd.DataFrame, 
                 sheet_name_a: str = "Dataset A", sheet_name_b: str = "Dataset B"):
        """
        Initialize the interactive dashboard
        """
        self.df_a = df_a.copy()
        self.df_b = df_b.copy()
        self.sheet_name_a = sheet_name_a
        self.sheet_name_b = sheet_name_b
        
        # Identify column types
        self.numeric_cols_a = self._identify_numeric_columns(df_a)
        self.numeric_cols_b = self._identify_numeric_columns(df_b)
        self.categorical_cols_a = self._identify_categorical_columns(df_a)
        self.categorical_cols_b = self._identify_categorical_columns(df_b)
        
        # Chart configuration
        self.color_palette = [
            '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
            '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'
        ]
        
    def _identify_numeric_columns(self, df: pd.DataFrame) -> List[str]:
        """Identify numeric columns in dataframe"""
        return df.select_dtypes(include=[np.number]).columns.tolist()
    
    def _identify_categorical_columns(self, df: pd.DataFrame) -> List[str]:
        """Identify categorical columns in dataframe"""
        return df.select_dtypes(include=['object', 'category']).columns.tolist()
    
    def create_interactive_comparison_chart(self, column: str, chart_type: str = "bar") -> go.Figure:
        """
        Create interactive comparison chart with drill-down capabilities
        """
        title = f"Interactive {chart_type.title()} Chart: {column}"
        
        data_a = self.df_a[column].dropna() if column in self.df_a.columns else pd.Series([])
        data_b = self.df_b[column].dropna() if column in self.df_b.columns else pd.Series([])
        
        fig = go.Figure()
        
        if column in self.numeric_cols_a or column in self.numeric_cols_b:
            # Numeric comparison with detailed hover
            stats_a = {
                'mean': data_a.mean() if len(data_a) > 0 else 0,
                'median': data_a.median() if len(data_a) > 0 else 0,
                'std': data_a.std() if len(data_a) > 0 else 0,
                'count': len(data_a)
            }
            stats_b = {
                'mean': data_b.mean() if len(data_b) > 0 else 0,
                'median': data_b.median() if len(data_b) > 0 else 0,
                'std': data_b.std() if len(data_b) > 0 else 0,
                'count': len(data_b)
            }
            
            metrics = ['Mean', 'Median', 'Std Dev']
            values_a = [stats_a['mean'], stats_a['median'], stats_a['std']]
            values_b = [stats_b['mean'], stats_b['median'], stats_b['std']]
            
            fig.add_trace(go.Bar(
                name=self.sheet_name_a,
                x=metrics,
                y=values_a,
                marker_color=self.color_palette[0],
                hovertemplate='<b>%{fullData.name}</b><br>' +
                             'Metric: %{x}<br>' +
                             'Value: %{y:.2f}<br>' +
                             f'Sample Size: {stats_a["count"]}<br>' +
                             '<extra></extra>',
                text=[f'{v:.2f}' for v in values_a],
                textposition='auto'
            ))
            
            fig.add_trace(go.Bar(
                name=self.sheet_name_b,
                x=metrics,
                y=values_b,
                marker_color=self.color_palette[1],
                hovertemplate='<b>%{fullData.name}</b><br>' +
                             'Metric: %{x}<br>' +
                             'Value: %{y:.2f}<br>' +
                             f'Sample Size: {stats_b["count"]}<br>' +
                             '<extra></extra>',
                text=[f'{v:.2f}' for v in values_b],
                textposition='auto'
            ))
        else:
            # Categorical comparison
            value_counts_a = data_a.value_counts().head(10)
            value_counts_b = data_b.value_counts().head(10)
            
            all_categories = sorted(set(value_counts_a.index.tolist() + value_counts_b.index.tolist()))
            
            counts_a = [value_counts_a.get(cat, 0) for cat in all_categories]
            counts_b = [value_counts_b.get(cat, 0) for cat in all_categories]
            
            fig.add_trace(go.Bar(
                name=self.sheet_name_a,
                x=all_categories,
                y=counts_a,
                marker_color=self.color_palette[0],
                hovertemplate='<b>%{fullData.name}</b><br>' +
                             'Category: %{x}<br>' +
                             'Count: %{y}<br>' +
                             'Percentage: %{customdata:.1f}%<br>' +
                             '<extra></extra>',
                customdata=[100 * c / sum(counts_a) if sum(counts_a) > 0 else 0 for c in counts_a],
                text=counts_a,
                textposition='auto'
            ))
            
            fig.add_trace(go.Bar(
                name=self.sheet_name_b,
                x=all_categories,
                y=counts_b,
                marker_color=self.color_palette[1],
                hovertemplate='<b>%{fullData.name}</b><br>' +
                             'Category: %{x}<br>' +
                             'Count: %{y}<br>' +
                             'Percentage: %{customdata:.1f}%<br>' +
                             '<extra></extra>',
                customdata=[100 * c / sum(counts_b) if sum(counts_b) > 0 else 0 for c in counts_b],
                text=counts_b,
                textposition='auto'
            ))
        
        # Enhanced layout with interaction capabilities
        fig.update_layout(
            title=title,
            barmode='group',
            hovermode='x unified',
            template='plotly_white',
            height=500,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            # Add custom buttons for interactivity
            updatemenus=[
                dict(
                    type="buttons",
                    direction="left",
                    buttons=list([
                        dict(
                            args=[{"visible": [True, True]}],
                            label="Show Both",
                            method="restyle"
                        ),
                        dict(
                            args=[{"visible": [True, False]}],
                            label=f"Show {self.sheet_name_a}",
                            method="restyle"
                        ),
                        dict(
                            args=[{"visible": [False, True]}],
                            label=f"Show {self.sheet_name_b}",
                            method="restyle"
                        )
                    ]),
                    pad={"r": 10, "t": 10},
                    showactive=True,
                    x=0.01,
                    xanchor="left",
                    y=1.15,
                    yanchor="top"
                ),
            ]
        )
        
        return fig
    
    def create_drill_down_scatter(self, x_col: str, y_col: str, color_col: str = None) -> go.Figure:
        """Create interactive scatter plot with drill-down"""
        
        title = f"Interactive Scatter: {y_col} vs {x_col}"
        if color_col:
            title += f" (colored by {color_col})"
        
        fig = go.Figure()
        
        # Dataset A scatter
        if x_col in self.df_a.columns and y_col in self.df_a.columns:
            df_a_clean = self.df_a.dropna(subset=[x_col, y_col])
            
            fig.add_trace(go.Scatter(
                x=df_a_clean[x_col],
                y=df_a_clean[y_col],
                mode='markers',
                name=self.sheet_name_a,
                marker=dict(
                    size=8,
                    color=df_a_clean[color_col] if color_col and color_col in self.df_a.columns else self.color_palette[0],
                    colorscale='Blues' if color_col else None,
                    showscale=True if color_col else False,
                    colorbar=dict(title=color_col, x=1.02) if color_col else None
                ),
                hovertemplate='<b>%{fullData.name}</b><br>' +
                             f'{x_col}: %{{x}}<br>' +
                             f'{y_col}: %{{y}}<br>' +
                             (f'{color_col}: %{{marker.color}}<br>' if color_col else '') +
                             '<extra></extra>'
            ))
        
        # Dataset B scatter
        if x_col in self.df_b.columns and y_col in self.df_b.columns:
            df_b_clean = self.df_b.dropna(subset=[x_col, y_col])
            
            fig.add_trace(go.Scatter(
                x=df_b_clean[x_col],
                y=df_b_clean[y_col],
                mode='markers',
                name=self.sheet_name_b,
                marker=dict(
                    size=8,
                    color=df_b_clean[color_col] if color_col and color_col in self.df_b.columns else self.color_palette[1],
                    colorscale='Reds' if color_col else None,
                    showscale=True if color_col else False,
                    colorbar=dict(title=color_col, x=1.1) if color_col else None
                ),
                hovertemplate='<b>%{fullData.name}</b><br>' +
                             f'{x_col}: %{{x}}<br>' +
                             f'{y_col}: %{{y}}<br>' +
                             (f'{color_col}: %{{marker.color}}<br>' if color_col else '') +
                             '<extra></extra>'
            ))
        
        fig.update_layout(
            title=title,
            xaxis_title=x_col,
            yaxis_title=y_col,
            template='plotly_white',
            height=500,
            hovermode='closest'
        )
        
        return fig
    
    def export_chart_with_data(self, fig: go.Figure, include_data: bool = True) -> Dict[str, Any]:
        """Export chart with optional underlying data"""
        
        # Generate chart image
        img_bytes = fig.to_image(format="png", width=1200, height=800)
        img_b64 = base64.b64encode(img_bytes).decode()
        
        export_data = {
            'chart_image': img_b64,
            'chart_html': fig.to_html(include_plotlyjs='cdn'),
            'export_timestamp': datetime.now().isoformat()
        }
        
        if include_data:
            # Extract data from figure traces
            chart_data = []
            for trace in fig.data:
                trace_data = {
                    'name': trace.name,
                    'type': trace.type,
                    'x': list(trace.x) if hasattr(trace, 'x') and trace.x is not None else [],
                    'y': list(trace.y) if hasattr(trace, 'y') and trace.y is not None else []
                }
                chart_data.append(trace_data)
            
            export_data['underlying_data'] = chart_data
        
        return export_data
    
    def create_single_dataset_bar_chart(self, df: pd.DataFrame, x_col: str, y_col: str, color_col: str = None) -> go.Figure:
        """Create interactive bar chart for a single dataset"""
        
        title = f"Interactive Bar Chart: {y_col} by {x_col}"
        
        fig = go.Figure()
        
        # Clean data
        df_clean = df.dropna(subset=[x_col, y_col])
        
        if len(df_clean) == 0:
            return self._create_no_data_chart("No valid data for selected columns")
        
        # Group data if needed
        if color_col and color_col in df.columns:
            # Grouped bar chart
            for category in df_clean[color_col].unique():
                category_data = df_clean[df_clean[color_col] == category]
                
                if len(category_data) > 20:  # Aggregate if too many points
                    grouped = category_data.groupby(x_col)[y_col].agg(['mean', 'sum', 'count']).reset_index()
                    fig.add_trace(go.Bar(
                        x=grouped[x_col],
                        y=grouped['sum'],
                        name=str(category),
                        hovertemplate='<b>%{fullData.name}</b><br>' +
                                     f'{x_col}: %{{x}}<br>' +
                                     f'Total {y_col}: %{{y:,.2f}}<br>' +
                                     f'Count: %{{customdata}}<br>' +
                                     '<extra></extra>',
                        customdata=grouped['count']
                    ))
                else:
                    fig.add_trace(go.Bar(
                        x=category_data[x_col],
                        y=category_data[y_col],
                        name=str(category),
                        hovertemplate='<b>%{fullData.name}</b><br>' +
                                     f'{x_col}: %{{x}}<br>' +
                                     f'{y_col}: %{{y:,.2f}}<br>' +
                                     '<extra></extra>'
                    ))
        else:
            # Simple bar chart
            if len(df_clean) > 20:  # Aggregate if too many points
                grouped = df_clean.groupby(x_col)[y_col].agg(['mean', 'sum', 'count']).reset_index()
                fig.add_trace(go.Bar(
                    x=grouped[x_col],
                    y=grouped['sum'],
                    name=f"Total {y_col}",
                    marker_color=self.color_palette[0],
                    hovertemplate=f'{x_col}: %{{x}}<br>' +
                                 f'Total {y_col}: %{{y:,.2f}}<br>' +
                                 f'Count: %{{customdata}}<br>' +
                                 '<extra></extra>',
                    customdata=grouped['count']
                ))
            else:
                fig.add_trace(go.Bar(
                    x=df_clean[x_col],
                    y=df_clean[y_col],
                    name=f"{y_col}",
                    marker_color=self.color_palette[0],
                    hovertemplate=f'{x_col}: %{{x}}<br>' +
                                 f'{y_col}: %{{y:,.2f}}<br>' +
                                 '<extra></extra>'
                ))
        
        # Update layout
        fig.update_layout(
            title=title,
            xaxis_title=x_col,
            yaxis_title=y_col,
            hovermode='closest',
            height=500,
            title_x=0.5,
            showlegend=bool(color_col)
        )
        
        return fig
    
    def create_single_dataset_scatter(self, df: pd.DataFrame, x_col: str, y_col: str, color_col: str = None) -> go.Figure:
        """Create interactive scatter plot for a single dataset"""
        
        title = f"Interactive Scatter: {y_col} vs {x_col}"
        if color_col:
            title += f" (colored by {color_col})"
        
        fig = go.Figure()
        
        # Clean data
        df_clean = df.dropna(subset=[x_col, y_col])
        
        if len(df_clean) == 0:
            return self._create_no_data_chart("No valid data for selected columns")
        
        if color_col and color_col in df.columns:
            # Colored scatter plot
            unique_colors = df_clean[color_col].unique()
            
            for i, category in enumerate(unique_colors):
                category_data = df_clean[df_clean[color_col] == category]
                
                fig.add_trace(go.Scatter(
                    x=category_data[x_col],
                    y=category_data[y_col],
                    mode='markers',
                    name=str(category),
                    marker=dict(
                        size=8,
                        color=self.color_palette[i % len(self.color_palette)]
                    ),
                    hovertemplate='<b>%{fullData.name}</b><br>' +
                                 f'{x_col}: %{{x}}<br>' +
                                 f'{y_col}: %{{y}}<br>' +
                                 f'{color_col}: {category}<br>' +
                                 '<extra></extra>'
                ))
        else:
            # Simple scatter plot
            fig.add_trace(go.Scatter(
                x=df_clean[x_col],
                y=df_clean[y_col],
                mode='markers',
                name=f"{y_col} vs {x_col}",
                marker=dict(
                    size=8,
                    color=self.color_palette[0]
                ),
                hovertemplate=f'{x_col}: %{{x}}<br>' +
                             f'{y_col}: %{{y}}<br>' +
                             '<extra></extra>'
            ))
        
        # Update layout
        fig.update_layout(
            title=title,
            xaxis_title=x_col,
            yaxis_title=y_col,
            hovermode='closest',
            height=500,
            title_x=0.5,
            showlegend=bool(color_col)
        )
        
        return fig


class ChartExporter:
    """Enhanced utility class for chart export functionality with multiple format support"""
    
    @staticmethod
    def export_single_chart(fig: go.Figure, filename: str = None, 
                           format: str = "png", width: int = 1200, height: int = 800) -> bytes:
        """Export single chart to specified format"""
        
        if filename is None:
            filename = f"chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{format}"
        
        if format.lower() == "html":
            html_str = fig.to_html(include_plotlyjs='cdn', div_id="chart")
            return html_str.encode('utf-8')
        else:
            return fig.to_image(format=format, width=width, height=height)
    
    @staticmethod
    def create_download_link(file_bytes: bytes, filename: str, format: str) -> str:
        """Create Streamlit download link for exported chart"""
        
        b64 = base64.b64encode(file_bytes).decode()
        
        mime_types = {
            'png': 'image/png',
            'jpg': 'image/jpeg',
            'pdf': 'application/pdf',
            'svg': 'image/svg+xml',
            'html': 'text/html'
        }
        
        mime_type = mime_types.get(format.lower(), 'application/octet-stream')
        
        return f'<a href="data:{mime_type};base64,{b64}" download="{filename}">Download {format.upper()}</a>'