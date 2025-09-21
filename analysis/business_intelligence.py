# Business Intelligence Analysis Module
# Provides comprehensive business insights and content-focused analysis

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

class BusinessIntelligenceAnalyzer:
    """
    Comprehensive Business Intelligence Analysis
    Focuses on business content, insights, and performance metrics
    """
    
    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self.total_rows = len(df)
        self.columns = df.columns.tolist()
        
        # Auto-detect common business columns
        self.date_columns = self._detect_date_columns()
        self.amount_columns = self._detect_amount_columns()
        self.category_columns = self._detect_category_columns()
        self.rating_columns = self._detect_rating_columns()
        self.id_columns = self._detect_id_columns()
        
    def _detect_date_columns(self) -> List[str]:
        """Auto-detect date/time columns"""
        date_cols = []
        for col in self.df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['date', 'time', 'created', 'updated', 'order']):
                # Check if column contains date-like data
                sample = self.df[col].dropna().head(10)
                if len(sample) > 0:
                    try:
                        pd.to_datetime(sample)
                        date_cols.append(col)
                    except:
                        pass
            elif self.df[col].dtype in ['datetime64[ns]', 'datetime64']:
                date_cols.append(col)
        return date_cols
    
    def _detect_amount_columns(self) -> List[str]:
        """Auto-detect monetary/amount columns"""
        amount_cols = []
        for col in self.df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['amount', 'price', 'cost', 'revenue', 'sales', 'value', 'total']):
                if pd.api.types.is_numeric_dtype(self.df[col]):
                    amount_cols.append(col)
        return amount_cols
    
    def _detect_category_columns(self) -> List[str]:
        """Auto-detect categorical columns"""
        category_cols = []
        for col in self.df.columns:
            col_lower = str(col).lower()
            unique_ratio = self.df[col].nunique() / len(self.df)
            
            # Low cardinality or specific business categories
            if (unique_ratio < 0.1 or 
                any(keyword in col_lower for keyword in ['category', 'type', 'status', 'region', 'product', 'department', 'rep'])):
                if self.df[col].dtype == 'object' or unique_ratio < 0.1:
                    category_cols.append(col)
        return category_cols
    
    def _detect_rating_columns(self) -> List[str]:
        """Auto-detect rating/score columns"""
        rating_cols = []
        for col in self.df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['rating', 'score', 'satisfaction', 'quality']):
                if pd.api.types.is_numeric_dtype(self.df[col]):
                    rating_cols.append(col)
        return rating_cols
    
    def _detect_id_columns(self) -> List[str]:
        """Auto-detect ID/identifier columns"""
        id_cols = []
        for col in self.df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['id', 'key', 'number', 'code']):
                unique_ratio = self.df[col].nunique() / len(self.df)
                if unique_ratio > 0.8:  # High uniqueness suggests ID column
                    id_cols.append(col)
        return id_cols
    
    def _calculate_financial_ratios(self, df: pd.DataFrame = None) -> Dict[str, Any]:
        """Calculate comprehensive financial ratios from RFP Bayanati specifications"""
        if df is None:
            df = self.df
            
        ratios = {
            'liquidity_ratios': {},
            'profitability_ratios': {},
            'efficiency_ratios': {},
            'market_ratios': {},
            'business_kpis': {},
            'banking_ratios': {},
            'insights': []
        }
        
        try:
            # Detect financial columns
            current_assets = self._detect_column_by_keywords(df, ['current_assets', 'assets', 'total_assets'])
            current_liabilities = self._detect_column_by_keywords(df, ['current_liabilities', 'liabilities', 'debt'])
            inventory = self._detect_column_by_keywords(df, ['inventory', 'stock'])
            cash = self._detect_column_by_keywords(df, ['cash', 'cash_flow'])
            revenue = self._detect_column_by_keywords(df, ['revenue', 'sales', 'income', 'amount'])
            cogs = self._detect_column_by_keywords(df, ['cogs', 'cost_of_goods', 'cost'])
            net_income = self._detect_column_by_keywords(df, ['net_income', 'profit', 'earnings'])
            equity = self._detect_column_by_keywords(df, ['equity', 'shareholders_equity'])
            total_assets = self._detect_column_by_keywords(df, ['total_assets', 'assets'])
            
            # LIQUIDITY RATIOS
            if current_assets and current_liabilities:
                # Current Ratio = Current Assets ÷ Current Liabilities
                current_ratio = df[current_assets].sum() / df[current_liabilities].sum() if df[current_liabilities].sum() != 0 else 0
                ratios['liquidity_ratios']['current_ratio'] = {
                    'value': round(current_ratio, 2),
                    'formula': 'Current Assets ÷ Current Liabilities',
                    'interpretation': 'Good' if current_ratio >= 1.5 else 'Needs Attention',
                    'benchmark': '≥ 1.5 is healthy'
                }
                
                # Quick Ratio = (Current Assets - Inventory) ÷ Current Liabilities
                if inventory:
                    quick_ratio = (df[current_assets].sum() - df[inventory].sum()) / df[current_liabilities].sum() if df[current_liabilities].sum() != 0 else 0
                    ratios['liquidity_ratios']['quick_ratio'] = {
                        'value': round(quick_ratio, 2),
                        'formula': '(Current Assets - Inventory) ÷ Current Liabilities',
                        'interpretation': 'Good' if quick_ratio >= 1.0 else 'Needs Attention'
                    }
                
                # Cash Ratio = Cash ÷ Current Liabilities
                if cash:
                    cash_ratio = df[cash].sum() / df[current_liabilities].sum() if df[current_liabilities].sum() != 0 else 0
                    ratios['liquidity_ratios']['cash_ratio'] = {
                        'value': round(cash_ratio, 2),
                        'formula': 'Cash ÷ Current Liabilities',
                        'interpretation': 'Good' if cash_ratio >= 0.2 else 'Low'
                    }
            
            # PROFITABILITY RATIOS
            if revenue and net_income:
                # Net Profit Margin = Net Income ÷ Revenue
                net_margin = net_income and revenue and df[net_income].sum() / df[revenue].sum() * 100 if df[revenue].sum() != 0 else 0
                if net_margin:
                    ratios['profitability_ratios']['net_profit_margin'] = {
                        'value': round(net_margin, 2),
                        'formula': 'Net Income ÷ Revenue × 100',
                        'unit': '%',
                        'interpretation': 'Excellent' if net_margin >= 20 else 'Good' if net_margin >= 10 else 'Needs Improvement'
                    }
                
                # Gross Profit Margin = (Revenue - COGS) ÷ Revenue
                if cogs:
                    gross_margin = (df[revenue].sum() - df[cogs].sum()) / df[revenue].sum() * 100 if df[revenue].sum() != 0 else 0
                    ratios['profitability_ratios']['gross_profit_margin'] = {
                        'value': round(gross_margin, 2),
                        'formula': '(Revenue - COGS) ÷ Revenue × 100',
                        'unit': '%',
                        'interpretation': 'Excellent' if gross_margin >= 40 else 'Good' if gross_margin >= 25 else 'Low'
                    }
                
                # ROE = Net Income ÷ Equity
                if equity:
                    roe = df[net_income].sum() / df[equity].sum() * 100 if df[equity].sum() != 0 else 0
                    ratios['profitability_ratios']['roe'] = {
                        'value': round(roe, 2),
                        'formula': 'Net Income ÷ Equity × 100',
                        'unit': '%',
                        'interpretation': 'Excellent' if roe >= 15 else 'Good' if roe >= 10 else 'Needs Improvement',
                        'benchmark': 'Target: 15-20%+'
                    }
                
                # ROA = Net Income ÷ Total Assets
                if total_assets:
                    roa = df[net_income].sum() / df[total_assets].sum() * 100 if df[total_assets].sum() != 0 else 0
                    ratios['profitability_ratios']['roa'] = {
                        'value': round(roa, 2),
                        'formula': 'Net Income ÷ Total Assets × 100',
                        'unit': '%',
                        'interpretation': 'Excellent' if roa >= 5 else 'Good' if roa >= 2 else 'Low'
                    }
            
            # BUSINESS KPIs from RFP requirements
            if len(df) > 1 and revenue:  # Need historical data for growth
                # Sales Growth % (if we have time series data)
                if self.date_columns and len(df) > 1:
                    df_sorted = df.sort_values(self.date_columns[0])
                    recent_revenue = df_sorted[revenue].tail(int(len(df)/2)).sum()
                    older_revenue = df_sorted[revenue].head(int(len(df)/2)).sum()
                    if older_revenue != 0:
                        sales_growth = (recent_revenue - older_revenue) / older_revenue * 100
                        ratios['business_kpis']['sales_growth'] = {
                            'value': round(sales_growth, 2),
                            'formula': '(Recent Sales - Previous Sales) ÷ Previous Sales × 100',
                            'unit': '%',
                            'interpretation': 'Excellent' if sales_growth >= 20 else 'Good' if sales_growth >= 10 else 'Declining' if sales_growth < 0 else 'Stable'
                        }
            
            # Customer metrics
            customer_cols = self._detect_customer_columns(df)
            if customer_cols and revenue:
                # Customer Lifetime Value estimation
                avg_transaction = df[revenue].mean()
                transaction_frequency = len(df) / df[customer_cols[0]].nunique() if df[customer_cols[0]].nunique() > 0 else 1
                estimated_ltv = avg_transaction * transaction_frequency * 12  # Annualized estimate
                
                ratios['business_kpis']['estimated_customer_ltv'] = {
                    'value': round(estimated_ltv, 2),
                    'formula': 'Avg Transaction × Frequency × 12 months',
                    'unit': '$',
                    'interpretation': 'High Value' if estimated_ltv >= 1000 else 'Medium Value' if estimated_ltv >= 500 else 'Low Value'
                }
            
            # Generate insights
            if ratios['profitability_ratios']:
                if 'roe' in ratios['profitability_ratios']:
                    roe_val = ratios['profitability_ratios']['roe']['value']
                    ratios['insights'].append(f"ROE of {roe_val}% indicates {'strong' if roe_val >= 15 else 'moderate' if roe_val >= 10 else 'weak'} profitability")
                
                if 'net_profit_margin' in ratios['profitability_ratios']:
                    margin_val = ratios['profitability_ratios']['net_profit_margin']['value']
                    ratios['insights'].append(f"Net margin of {margin_val}% shows {'excellent' if margin_val >= 20 else 'good' if margin_val >= 10 else 'concerning'} efficiency")
            
            if ratios['liquidity_ratios']:
                if 'current_ratio' in ratios['liquidity_ratios']:
                    current_val = ratios['liquidity_ratios']['current_ratio']['value']
                    ratios['insights'].append(f"Current ratio of {current_val} indicates {'strong' if current_val >= 2 else 'adequate' if current_val >= 1.5 else 'tight'} liquidity")
            
        except Exception as e:
            ratios['error'] = f"Financial ratio calculation error: {str(e)}"
            ratios['insights'].append("Financial data may need standardization for accurate ratio calculation")
        
        return ratios
    
    def _detect_column_by_keywords(self, df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
        """Detect column by keywords with fuzzy matching"""
        for col in df.columns:
            col_lower = str(col).lower().replace('_', ' ').replace('-', ' ')
            for keyword in keywords:
                if keyword.lower() in col_lower:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        return col
        return None
    
    def _detect_customer_columns(self, df: pd.DataFrame = None) -> List[str]:
        """Auto-detect customer-related columns"""
        if df is None:
            df = self.df
        customer_cols = []
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['customer', 'client', 'user', 'buyer', 'account']):
                customer_cols.append(col)
        return customer_cols
    
    def generate_business_overview(self) -> Dict[str, Any]:
        """Generate comprehensive business overview with financial ratios"""
        
        overview = {
            'dataset_info': {
                'total_records': self.total_rows,
                'total_columns': len(self.columns),
                'date_range': self._get_date_range(),
                'business_metrics_available': len(self.amount_columns) > 0
            },
            'column_classification': {
                'date_columns': self.date_columns,
                'amount_columns': self.amount_columns,
                'category_columns': self.category_columns,
                'rating_columns': self.rating_columns,
                'id_columns': self.id_columns
            },
            'financial_ratios': self._calculate_financial_ratios(),
            'key_business_insights': self._extract_key_insights()
        }
        
        return overview
    
    def _get_date_range(self) -> Dict[str, Any]:
        """Get date range information"""
        if not self.date_columns:
            return {'status': 'No date columns detected'}
        
        date_info = {}
        for col in self.date_columns:
            try:
                date_series = pd.to_datetime(self.df[col])
                date_info[col] = {
                    'start_date': date_series.min().strftime('%Y-%m-%d'),
                    'end_date': date_series.max().strftime('%Y-%m-%d'),
                    'duration_days': (date_series.max() - date_series.min()).days,
                    'data_points': date_series.notna().sum()
                }
            except:
                date_info[col] = {'status': 'Date parsing failed'}
        
        return date_info
    
    def _extract_key_insights(self) -> List[str]:
        """Extract key business insights"""
        insights = []
        
        # Revenue insights
        if self.amount_columns:
            for col in self.amount_columns:
                total = self.df[col].sum()
                avg = self.df[col].mean()
                insights.append(f"Total {col}: ${total:,.2f} | Average: ${avg:,.2f}")
        
        # Category insights
        if self.category_columns:
            for col in self.category_columns[:2]:  # Limit to top 2
                top_category = self.df[col].value_counts().head(1)
                if not top_category.empty:
                    insights.append(f"Top {col}: {top_category.index[0]} ({top_category.values[0]} records)")
        
        # Rating insights
        if self.rating_columns:
            for col in self.rating_columns:
                avg_rating = self.df[col].mean()
                insights.append(f"Average {col}: {avg_rating:.2f}")
        
        # Time-based insights
        if self.date_columns:
            insights.append(f"Data spans {len(self.date_columns)} time dimension(s)")
        
        return insights[:6]  # Limit to top 6 insights
    
    def analyze_sales_performance(self) -> Dict[str, Any]:
        """Comprehensive sales performance analysis"""
        
        if not self.amount_columns:
            return {'error': 'No sales/amount columns detected for analysis'}
        
        sales_analysis = {}
        
        for amount_col in self.amount_columns:
            analysis = {
                'total_revenue': float(self.df[amount_col].sum()),
                'average_transaction': float(self.df[amount_col].mean()),
                'median_transaction': float(self.df[amount_col].median()),
                'top_transactions': self._get_top_transactions(amount_col),
                'revenue_distribution': self._analyze_revenue_distribution(amount_col),
                'performance_metrics': self._calculate_performance_metrics(amount_col)
            }
            
            # Time-based analysis if date columns available
            if self.date_columns:
                analysis['time_trends'] = self._analyze_time_trends(amount_col)
            
            # Category-based analysis if category columns available
            if self.category_columns:
                analysis['category_performance'] = self._analyze_category_performance(amount_col)
            
            sales_analysis[amount_col] = analysis
        
        return sales_analysis
    
    def _get_top_transactions(self, amount_col: str, top_n: int = 10) -> List[Dict]:
        """Get top transactions by amount"""
        top_transactions = self.df.nlargest(top_n, amount_col)
        
        result = []
        for _, row in top_transactions.iterrows():
            transaction = {'amount': float(row[amount_col])}
            
            # Add relevant context columns
            for col in self.category_columns + self.id_columns + self.date_columns:
                if col in row and pd.notna(row[col]):
                    transaction[col] = str(row[col])
            
            result.append(transaction)
        
        return result
    
    def _analyze_revenue_distribution(self, amount_col: str) -> Dict[str, Any]:
        """Analyze revenue distribution patterns"""
        amounts = self.df[amount_col].dropna()
        
        return {
            'quartiles': {
                'Q1': float(amounts.quantile(0.25)),
                'Q2_median': float(amounts.quantile(0.5)),
                'Q3': float(amounts.quantile(0.75)),
                'Q4_max': float(amounts.max())
            },
            'percentiles': {
                'P90': float(amounts.quantile(0.9)),
                'P95': float(amounts.quantile(0.95)),
                'P99': float(amounts.quantile(0.99))
            },
            'distribution_stats': {
                'std_deviation': float(amounts.std()),
                'coefficient_of_variation': float(amounts.std() / amounts.mean()) if amounts.mean() != 0 else 0,
                'skewness': float(amounts.skew())
            }
        }
    
    def _calculate_performance_metrics(self, amount_col: str) -> Dict[str, Any]:
        """Calculate key performance metrics"""
        amounts = self.df[amount_col].dropna()
        
        # Basic metrics
        metrics = {
            'total_transactions': len(amounts),
            'revenue_per_transaction': float(amounts.mean()),
            'conversion_rate': len(amounts) / self.total_rows * 100  # Assuming non-null amounts are conversions
        }
        
        # Growth metrics (if multiple time periods)
        if self.date_columns and len(amounts) > 10:
            try:
                date_col = self.date_columns[0]
                df_with_dates = self.df[[date_col, amount_col]].dropna()
                df_with_dates[date_col] = pd.to_datetime(df_with_dates[date_col])
                
                # Monthly growth
                monthly_revenue = df_with_dates.groupby(df_with_dates[date_col].dt.to_period('M'))[amount_col].sum()
                if len(monthly_revenue) > 1:
                    growth_rate = ((monthly_revenue.iloc[-1] - monthly_revenue.iloc[0]) / monthly_revenue.iloc[0] * 100)
                    metrics['revenue_growth_rate'] = float(growth_rate)
                
            except Exception:
                metrics['revenue_growth_rate'] = 'Unable to calculate'
        
        return metrics
    
    def _analyze_time_trends(self, amount_col: str) -> Dict[str, Any]:
        """Analyze time-based trends"""
        if not self.date_columns:
            return {'error': 'No date columns available'}
        
        trends = {}
        
        for date_col in self.date_columns:
            try:
                df_time = self.df[[date_col, amount_col]].dropna()
                df_time[date_col] = pd.to_datetime(df_time[date_col])
                
                # Monthly trends
                monthly_data = df_time.groupby(df_time[date_col].dt.to_period('M')).agg({
                    amount_col: ['sum', 'mean', 'count']
                }).round(2)
                
                trends[date_col] = {
                    'monthly_revenue': monthly_data[amount_col]['sum'].to_dict(),
                    'monthly_average': monthly_data[amount_col]['mean'].to_dict(),
                    'monthly_transactions': monthly_data[amount_col]['count'].to_dict(),
                    'peak_month': str(monthly_data[amount_col]['sum'].idxmax()),
                    'lowest_month': str(monthly_data[amount_col]['sum'].idxmin())
                }
                
            except Exception as e:
                trends[date_col] = {'error': f'Time analysis failed: {str(e)}'}
        
        return trends
    
    def _analyze_category_performance(self, amount_col: str) -> Dict[str, Any]:
        """Analyze performance by categories"""
        category_analysis = {}
        
        for cat_col in self.category_columns:
            try:
                cat_performance = self.df.groupby(cat_col)[amount_col].agg([
                    'sum', 'mean', 'count', 'std'
                ]).round(2)
                
                category_analysis[cat_col] = {
                    'revenue_by_category': cat_performance['sum'].to_dict(),
                    'average_by_category': cat_performance['mean'].to_dict(),
                    'transaction_count': cat_performance['count'].to_dict(),
                    'top_performer': cat_performance['sum'].idxmax(),
                    'lowest_performer': cat_performance['sum'].idxmin(),
                    'performance_variance': cat_performance['std'].to_dict()
                }
                
            except Exception as e:
                category_analysis[cat_col] = {'error': f'Category analysis failed: {str(e)}'}
        
        return category_analysis
    
    def analyze_customer_insights(self) -> Dict[str, Any]:
        """Comprehensive customer analysis"""
        
        customer_analysis = {
            'customer_segmentation': self._segment_customers(),
            'customer_behavior': self._analyze_customer_behavior(),
            'satisfaction_analysis': self._analyze_customer_satisfaction(),
            'customer_value_analysis': self._analyze_customer_value()
        }
        
        return customer_analysis
    
    def _segment_customers(self) -> Dict[str, Any]:
        """Customer segmentation analysis"""
        
        if not self.amount_columns:
            return {'error': 'No transaction amounts available for segmentation'}
        
        amount_col = self.amount_columns[0]
        amounts = self.df[amount_col].dropna()
        
        # RFM-style segmentation based on available data
        segments = {}
        
        # Value-based segmentation
        q75 = amounts.quantile(0.75)
        q50 = amounts.quantile(0.50)
        q25 = amounts.quantile(0.25)
        
        segments['value_segments'] = {
            'high_value': len(amounts[amounts > q75]),
            'medium_value': len(amounts[(amounts > q50) & (amounts <= q75)]),
            'low_value': len(amounts[amounts <= q50]),
            'segment_thresholds': {
                'high_value_min': float(q75),
                'medium_value_min': float(q50),
                'low_value_max': float(q50)
            }
        }
        
        # Frequency analysis if customer identifiers available
        customer_cols = [col for col in self.columns if 'customer' in str(col).lower() or 'client' in str(col).lower()]
        if customer_cols:
            customer_col = customer_cols[0]
            customer_frequency = self.df[customer_col].value_counts()
            
            segments['frequency_segments'] = {
                'repeat_customers': len(customer_frequency[customer_frequency > 1]),
                'one_time_customers': len(customer_frequency[customer_frequency == 1]),
                'top_customers': customer_frequency.head(10).to_dict(),
                'average_transactions_per_customer': float(customer_frequency.mean())
            }
        
        return segments
    
    def _analyze_customer_behavior(self) -> Dict[str, Any]:
        """Customer behavior pattern analysis"""
        
        behavior_analysis = {}
        
        # Purchase patterns by categories
        if self.category_columns and self.amount_columns:
            cat_col = self.category_columns[0]
            amount_col = self.amount_columns[0]
            
            behavior_analysis['purchase_patterns'] = {
                'category_preferences': self.df[cat_col].value_counts().head(10).to_dict(),
                'spending_by_category': self.df.groupby(cat_col)[amount_col].mean().round(2).to_dict()
            }
        
        # Time-based behavior
        if self.date_columns:
            date_col = self.date_columns[0]
            try:
                df_dates = self.df.copy()
                df_dates[date_col] = pd.to_datetime(df_dates[date_col])
                
                behavior_analysis['temporal_patterns'] = {
                    'transactions_by_weekday': df_dates[date_col].dt.day_name().value_counts().to_dict(),
                    'transactions_by_month': df_dates[date_col].dt.month_name().value_counts().to_dict(),
                    'seasonal_trends': self._identify_seasonal_trends(df_dates, date_col)
                }
                
            except Exception as e:
                behavior_analysis['temporal_patterns'] = {'error': f'Date analysis failed: {str(e)}'}
        
        return behavior_analysis
    
    def _analyze_customer_satisfaction(self) -> Dict[str, Any]:
        """Customer satisfaction analysis"""
        
        if not self.rating_columns:
            return {'message': 'No rating/satisfaction columns detected'}
        
        satisfaction_analysis = {}
        
        for rating_col in self.rating_columns:
            ratings = self.df[rating_col].dropna()
            
            satisfaction_analysis[rating_col] = {
                'average_rating': float(ratings.mean()),
                'rating_distribution': ratings.value_counts().sort_index().to_dict(),
                'satisfaction_level': self._categorize_satisfaction(ratings.mean()),
                'total_responses': len(ratings),
                'rating_trends': self._analyze_rating_trends(rating_col) if self.date_columns else None
            }
        
        return satisfaction_analysis
    
    def _categorize_satisfaction(self, avg_rating: float) -> str:
        """Categorize satisfaction level"""
        if avg_rating >= 4.5:
            return 'Excellent'
        elif avg_rating >= 4.0:
            return 'Good'
        elif avg_rating >= 3.5:
            return 'Average'
        elif avg_rating >= 3.0:
            return 'Below Average'
        else:
            return 'Poor'
    
    def _analyze_rating_trends(self, rating_col: str) -> Dict[str, Any]:
        """Analyze rating trends over time"""
        if not self.date_columns:
            return None
        
        try:
            date_col = self.date_columns[0]
            df_ratings = self.df[[date_col, rating_col]].dropna()
            df_ratings[date_col] = pd.to_datetime(df_ratings[date_col])
            
            monthly_ratings = df_ratings.groupby(df_ratings[date_col].dt.to_period('M'))[rating_col].mean()
            
            return {
                'monthly_trends': monthly_ratings.round(2).to_dict(),
                'trend_direction': 'improving' if monthly_ratings.iloc[-1] > monthly_ratings.iloc[0] else 'declining',
                'best_month': str(monthly_ratings.idxmax()),
                'worst_month': str(monthly_ratings.idxmin())
            }
            
        except Exception:
            return {'error': 'Unable to analyze rating trends'}
    
    def _analyze_customer_value(self) -> Dict[str, Any]:
        """Customer lifetime value and value analysis"""
        
        if not self.amount_columns:
            return {'error': 'No transaction amounts available for value analysis'}
        
        amount_col = self.amount_columns[0]
        value_analysis = {
            'total_customer_value': float(self.df[amount_col].sum()),
            'average_customer_value': float(self.df[amount_col].mean()),
            'value_distribution': self._get_value_distribution(amount_col)
        }
        
        # Customer-specific analysis if customer identifiers exist
        customer_cols = [col for col in self.columns if 'customer' in str(col).lower()]
        if customer_cols:
            customer_col = customer_cols[0]
            customer_values = self.df.groupby(customer_col)[amount_col].agg(['sum', 'count', 'mean'])
            
            value_analysis['customer_lifetime_value'] = {
                'top_customers_by_value': customer_values['sum'].nlargest(10).to_dict(),
                'top_customers_by_frequency': customer_values['count'].nlargest(10).to_dict(),
                'average_clv': float(customer_values['sum'].mean()),
                'clv_distribution': self._get_clv_distribution(customer_values['sum'])
            }
        
        return value_analysis
    
    def _get_value_distribution(self, amount_col: str) -> Dict[str, Any]:
        """Get customer value distribution"""
        amounts = self.df[amount_col].dropna()
        
        return {
            'high_value_transactions': len(amounts[amounts > amounts.quantile(0.8)]),
            'medium_value_transactions': len(amounts[(amounts > amounts.quantile(0.4)) & (amounts <= amounts.quantile(0.8))]),
            'low_value_transactions': len(amounts[amounts <= amounts.quantile(0.4)]),
            'value_concentration': float(amounts.quantile(0.8)) / float(amounts.mean()) if amounts.mean() != 0 else 0
        }
    
    def _get_clv_distribution(self, clv_series: pd.Series) -> Dict[str, Any]:
        """Get customer lifetime value distribution"""
        return {
            'high_clv_customers': len(clv_series[clv_series > clv_series.quantile(0.8)]),
            'medium_clv_customers': len(clv_series[(clv_series > clv_series.quantile(0.4)) & (clv_series <= clv_series.quantile(0.8))]),
            'low_clv_customers': len(clv_series[clv_series <= clv_series.quantile(0.4)]),
            'clv_variance': float(clv_series.std())
        }
    
    def _identify_seasonal_trends(self, df_dates: pd.DataFrame, date_col: str) -> Dict[str, Any]:
        """Identify seasonal trends in the data"""
        try:
            # Quarterly analysis
            quarterly = df_dates.groupby(df_dates[date_col].dt.quarter).size()
            
            return {
                'quarterly_distribution': quarterly.to_dict(),
                'peak_quarter': f"Q{quarterly.idxmax()}",
                'seasonal_pattern': 'detected' if quarterly.std() > quarterly.mean() * 0.2 else 'minimal'
            }
        except Exception:
            return {'error': 'Unable to identify seasonal trends'}
    
    def analyze_product_performance(self) -> Dict[str, Any]:
        """Comprehensive product performance analysis"""
        
        product_analysis = {}
        
        # Find product-related columns
        product_cols = [col for col in self.columns if any(keyword in str(col).lower() for keyword in ['product', 'item', 'service'])]
        
        if not product_cols:
            return {'message': 'No product columns detected in the dataset'}
        
        for product_col in product_cols:
            analysis = {
                'product_overview': self._analyze_product_overview(product_col),
                'product_performance': self._analyze_product_sales_performance(product_col),
                'product_trends': self._analyze_product_trends(product_col) if self.date_columns else None
            }
            
            product_analysis[product_col] = analysis
        
        return product_analysis
    
    def _analyze_product_overview(self, product_col: str) -> Dict[str, Any]:
        """Product overview analysis"""
        product_counts = self.df[product_col].value_counts()
        
        return {
            'total_products': len(product_counts),
            'most_popular_product': product_counts.index[0] if len(product_counts) > 0 else None,
            'least_popular_product': product_counts.index[-1] if len(product_counts) > 0 else None,
            'product_distribution': product_counts.head(10).to_dict(),
            'product_diversity_index': len(product_counts) / self.total_rows  # Higher = more diverse
        }
    
    def _analyze_product_sales_performance(self, product_col: str) -> Dict[str, Any]:
        """Product sales performance analysis"""
        
        if not self.amount_columns:
            return {'error': 'No sales amount columns available'}
        
        performance = {}
        
        for amount_col in self.amount_columns:
            product_sales = self.df.groupby(product_col)[amount_col].agg(['sum', 'mean', 'count']).round(2)
            
            performance[amount_col] = {
                'revenue_by_product': product_sales['sum'].to_dict(),
                'average_price_by_product': product_sales['mean'].to_dict(),
                'sales_volume_by_product': product_sales['count'].to_dict(),
                'top_revenue_product': product_sales['sum'].idxmax(),
                'top_volume_product': product_sales['count'].idxmax(),
                'highest_priced_product': product_sales['mean'].idxmax()
            }
        
        return performance
    
    def _analyze_product_trends(self, product_col: str) -> Dict[str, Any]:
        """Product trends over time"""
        
        if not self.date_columns or not self.amount_columns:
            return {'error': 'Need both date and amount columns for trend analysis'}
        
        try:
            date_col = self.date_columns[0]
            amount_col = self.amount_columns[0]
            
            df_trends = self.df[[date_col, product_col, amount_col]].dropna()
            df_trends[date_col] = pd.to_datetime(df_trends[date_col])
            
            # Monthly trends by product
            monthly_trends = df_trends.groupby([df_trends[date_col].dt.to_period('M'), product_col])[amount_col].sum().unstack(fill_value=0)
            
            trends = {}
            for product in monthly_trends.columns:
                product_trend = monthly_trends[product]
                trends[product] = {
                    'monthly_sales': product_trend.to_dict(),
                    'trend_direction': 'growing' if product_trend.iloc[-1] > product_trend.iloc[0] else 'declining',
                    'peak_month': str(product_trend.idxmax()),
                    'growth_rate': ((product_trend.iloc[-1] - product_trend.iloc[0]) / product_trend.iloc[0] * 100) if product_trend.iloc[0] != 0 else 0
                }
            
            return trends
            
        except Exception as e:
            return {'error': f'Trend analysis failed: {str(e)}'}
    
    def generate_business_recommendations(self) -> List[str]:
        """Generate actionable business recommendations based on analysis"""
        
        recommendations = []
        
        # Revenue recommendations
        if self.amount_columns:
            amount_col = self.amount_columns[0]
            amounts = self.df[amount_col].dropna()
            
            if amounts.std() / amounts.mean() > 1.0:  # High variance
                recommendations.append("📊 High revenue variance detected - consider implementing tiered pricing strategy")
            
            if len(amounts[amounts < amounts.quantile(0.1)]) > len(amounts) * 0.1:
                recommendations.append("💰 Many low-value transactions - explore upselling opportunities")
        
        # Customer recommendations
        if self.rating_columns:
            rating_col = self.rating_columns[0]
            avg_rating = self.df[rating_col].mean()
            
            if avg_rating < 3.5:
                recommendations.append("⭐ Customer satisfaction below average - prioritize service improvement")
            elif avg_rating > 4.5:
                recommendations.append("🌟 Excellent customer satisfaction - leverage for referral programs")
        
        # Product recommendations
        product_cols = [col for col in self.category_columns if 'product' in str(col).lower()]
        if product_cols and self.amount_columns:
            product_col = product_cols[0]
            amount_col = self.amount_columns[0]
            
            product_performance = self.df.groupby(product_col)[amount_col].sum()
            if len(product_performance) > 1:
                top_product_share = product_performance.max() / product_performance.sum()
                if top_product_share > 0.5:
                    recommendations.append("🎯 Revenue concentrated in one product - diversify product portfolio")
        
        # Seasonal recommendations
        if self.date_columns:
            try:
                date_col = self.date_columns[0]
                df_seasonal = self.df.copy()
                df_seasonal[date_col] = pd.to_datetime(df_seasonal[date_col])
                
                monthly_volume = df_seasonal.groupby(df_seasonal[date_col].dt.month).size()
                if monthly_volume.std() / monthly_volume.mean() > 0.5:
                    recommendations.append("📅 Seasonal patterns detected - plan inventory and marketing accordingly")
            except:
                pass
        
        # Category recommendations
        if len(self.category_columns) > 1:
            recommendations.append("🔍 Multiple business dimensions available - consider cross-category analysis")
        
        # Data quality recommendations
        missing_pct = self.df.isnull().sum().sum() / (len(self.df) * len(self.df.columns)) * 100
        if missing_pct > 10:
            recommendations.append("📋 Significant missing data detected - improve data collection processes")
        
        return recommendations[:8]  # Limit to top 8 recommendations
    
    def analyze_profitability_metrics(self) -> Dict[str, Any]:
        """Advanced profitability and financial analysis"""
        
        profitability_analysis = {
            'discount_effectiveness': self._analyze_discount_effectiveness(),
            'revenue_optimization': self._analyze_revenue_optimization(),
            'financial_kpis': self._calculate_financial_kpis(),
            'margin_analysis': self._analyze_profit_margins()
        }
        
        return profitability_analysis
    
    def _analyze_discount_effectiveness(self) -> Dict[str, Any]:
        """Analyze discount effectiveness and impact"""
        
        discount_cols = [col for col in self.columns if 'discount' in str(col).lower() or 'promo' in str(col).lower()]
        
        if not discount_cols or not self.amount_columns:
            return {'message': 'No discount or amount columns available for analysis'}
        
        discount_col = discount_cols[0]
        amount_col = self.amount_columns[0]
        
        # Remove rows with missing discount or amount data
        df_discount = self.df[[discount_col, amount_col]].dropna()
        
        if len(df_discount) == 0:
            return {'message': 'No valid discount data available'}
        
        # Categorize discounts
        df_discount['discount_category'] = pd.cut(
            df_discount[discount_col], 
            bins=[-float('inf'), 0, 5, 15, 30, float('inf')],
            labels=['No Discount', 'Low (0-5%)', 'Medium (5-15%)', 'High (15-30%)', 'Very High (30%+)']
        )
        
        discount_analysis = df_discount.groupby('discount_category')[amount_col].agg([
            'count', 'sum', 'mean', 'std'
        ]).round(2)
        
        return {
            'discount_impact': {
                'transaction_count_by_discount': discount_analysis['count'].to_dict(),
                'total_revenue_by_discount': discount_analysis['sum'].to_dict(),
                'average_transaction_by_discount': discount_analysis['mean'].to_dict(),
                'revenue_variance_by_discount': discount_analysis['std'].to_dict()
            },
            'discount_effectiveness_score': self._calculate_discount_effectiveness_score(df_discount, discount_col, amount_col),
            'optimal_discount_range': self._find_optimal_discount_range(df_discount, discount_col, amount_col)
        }
    
    def _calculate_discount_effectiveness_score(self, df_discount: pd.DataFrame, 
                                              discount_col: str, amount_col: str) -> Dict[str, Any]:
        """Calculate discount effectiveness score"""
        
        # Correlation between discount and transaction amount
        correlation = df_discount[discount_col].corr(df_discount[amount_col])
        
        # Revenue per discount point
        total_discounts = df_discount[discount_col].sum()
        total_revenue = df_discount[amount_col].sum()
        revenue_per_discount_point = total_revenue / total_discounts if total_discounts > 0 else 0
        
        return {
            'discount_revenue_correlation': float(correlation) if not pd.isna(correlation) else 0,
            'revenue_per_discount_point': float(revenue_per_discount_point),
            'effectiveness_rating': self._rate_discount_effectiveness(correlation, revenue_per_discount_point)
        }
    
    def _rate_discount_effectiveness(self, correlation: float, revenue_per_point: float) -> str:
        """Rate discount effectiveness"""
        if pd.isna(correlation):
            return 'Insufficient data'
        
        if correlation < -0.3 and revenue_per_point < 50:
            return 'Poor - discounts may be reducing profitability'
        elif correlation < 0 and revenue_per_point > 50:
            return 'Moderate - discounts driving volume but reducing margins'
        elif correlation > 0.1:
            return 'Good - discounts correlated with higher transaction values'
        else:
            return 'Average - mixed discount effectiveness'
    
    def _find_optimal_discount_range(self, df_discount: pd.DataFrame, 
                                   discount_col: str, amount_col: str) -> Dict[str, Any]:
        """Find optimal discount range for revenue maximization"""
        
        # Group by discount ranges and calculate metrics
        df_discount['discount_range'] = pd.cut(df_discount[discount_col], bins=10)
        range_analysis = df_discount.groupby('discount_range')[amount_col].agg(['count', 'sum', 'mean'])
        
        # Find range with highest total revenue
        optimal_range_revenue = range_analysis['sum'].idxmax()
        
        # Find range with highest average transaction
        optimal_range_avg = range_analysis['mean'].idxmax()
        
        return {
            'optimal_for_total_revenue': str(optimal_range_revenue),
            'optimal_for_average_transaction': str(optimal_range_avg),
            'revenue_by_discount_range': range_analysis['sum'].to_dict()
        }
    
    def _analyze_revenue_optimization(self) -> Dict[str, Any]:
        """Analyze revenue optimization opportunities"""
        
        if not self.amount_columns:
            return {'message': 'No amount columns available for revenue optimization'}
        
        amount_col = self.amount_columns[0]
        optimization_analysis = {}
        
        # Price point analysis
        amounts = self.df[amount_col].dropna()
        optimization_analysis['price_point_analysis'] = {
            'current_average_price': float(amounts.mean()),
            'price_distribution': {
                'low_price_transactions': len(amounts[amounts < amounts.quantile(0.33)]),
                'medium_price_transactions': len(amounts[(amounts >= amounts.quantile(0.33)) & (amounts < amounts.quantile(0.67))]),
                'high_price_transactions': len(amounts[amounts >= amounts.quantile(0.67)])
            },
            'revenue_concentration': {
                'top_20_percent_contribution': float(amounts.nlargest(int(len(amounts) * 0.2)).sum() / amounts.sum() * 100),
                'bottom_20_percent_contribution': float(amounts.nsmallest(int(len(amounts) * 0.2)).sum() / amounts.sum() * 100)
            }
        }
        
        # Category optimization (if available)
        if self.category_columns:
            cat_col = self.category_columns[0]
            category_revenue = self.df.groupby(cat_col)[amount_col].agg(['sum', 'count', 'mean'])
            
            optimization_analysis['category_optimization'] = {
                'underperforming_categories': category_revenue[category_revenue['mean'] < amounts.mean()]['mean'].to_dict(),
                'high_potential_categories': category_revenue[category_revenue['count'] > category_revenue['count'].median()]['mean'].to_dict(),
                'revenue_per_category': category_revenue['sum'].to_dict()
            }
        
        return optimization_analysis
    
    def _calculate_financial_kpis(self) -> Dict[str, Any]:
        """Calculate key financial performance indicators"""
        
        if not self.amount_columns:
            return {'message': 'No financial data available for KPI calculation'}
        
        amount_col = self.amount_columns[0]
        amounts = self.df[amount_col].dropna()
        
        kpis = {
            'revenue_metrics': {
                'total_revenue': float(amounts.sum()),
                'average_revenue_per_transaction': float(amounts.mean()),
                'median_revenue_per_transaction': float(amounts.median()),
                'revenue_standard_deviation': float(amounts.std())
            },
            'volume_metrics': {
                'total_transactions': len(amounts),
                'transaction_frequency': len(amounts) / self.total_rows * 100  # Percentage of rows with transactions
            }
        }
        
        # Quantity-based metrics (if quantity column exists)
        quantity_cols = [col for col in self.columns if 'quantity' in str(col).lower() or 'qty' in str(col).lower() or 'volume' in str(col).lower()]
        if quantity_cols:
            qty_col = quantity_cols[0]
            quantities = self.df[qty_col].dropna()
            
            kpis['quantity_metrics'] = {
                'total_quantity_sold': float(quantities.sum()),
                'average_quantity_per_transaction': float(quantities.mean()),
                'revenue_per_unit': float(amounts.sum() / quantities.sum()) if quantities.sum() > 0 else 0
            }
        
        # Time-based KPIs (if date columns exist)
        if self.date_columns:
            date_col = self.date_columns[0]
            try:
                df_dates = self.df[[date_col, amount_col]].dropna()
                df_dates[date_col] = pd.to_datetime(df_dates[date_col])
                
                date_range = (df_dates[date_col].max() - df_dates[date_col].min()).days
                if date_range > 0:
                    kpis['time_based_metrics'] = {
                        'revenue_per_day': float(amounts.sum() / date_range),
                        'transactions_per_day': len(amounts) / date_range,
                        'data_period_days': date_range
                    }
            except:
                kpis['time_based_metrics'] = {'error': 'Unable to calculate time-based metrics'}
        
        return kpis
    
    def _analyze_profit_margins(self) -> Dict[str, Any]:
        """Analyze profit margins and cost structures"""
        
        # Look for cost/expense columns
        cost_cols = [col for col in self.columns if any(keyword in str(col).lower() for keyword in ['cost', 'expense', 'cogs'])]
        
        if not cost_cols or not self.amount_columns:
            return {'message': 'No cost or revenue columns available for margin analysis'}
        
        amount_col = self.amount_columns[0]
        cost_col = cost_cols[0]
        
        # Calculate margins
        df_margin = self.df[[amount_col, cost_col]].dropna()
        
        if len(df_margin) == 0:
            return {'message': 'No valid revenue-cost pairs for margin analysis'}
        
        df_margin['profit'] = df_margin[amount_col] - df_margin[cost_col]
        df_margin['margin_percent'] = (df_margin['profit'] / df_margin[amount_col] * 100).round(2)
        
        margin_analysis = {
            'overall_margins': {
                'total_revenue': float(df_margin[amount_col].sum()),
                'total_costs': float(df_margin[cost_col].sum()),
                'total_profit': float(df_margin['profit'].sum()),
                'overall_margin_percent': float(df_margin['profit'].sum() / df_margin[amount_col].sum() * 100)
            },
            'margin_distribution': {
                'average_margin_percent': float(df_margin['margin_percent'].mean()),
                'median_margin_percent': float(df_margin['margin_percent'].median()),
                'margin_volatility': float(df_margin['margin_percent'].std())
            },
            'profitability_segments': {
                'high_margin_transactions': len(df_margin[df_margin['margin_percent'] > 30]),
                'medium_margin_transactions': len(df_margin[(df_margin['margin_percent'] > 15) & (df_margin['margin_percent'] <= 30)]),
                'low_margin_transactions': len(df_margin[df_margin['margin_percent'] <= 15]),
                'negative_margin_transactions': len(df_margin[df_margin['margin_percent'] < 0])
            }
        }
        
        # Category-based margin analysis
        if self.category_columns:
            cat_col = self.category_columns[0]
            if cat_col in df_margin.columns:
                category_margins = df_margin.groupby(cat_col).agg({
                    'profit': 'sum',
                    'margin_percent': 'mean',
                    amount_col: 'sum'
                }).round(2)
                
                margin_analysis['margin_by_category'] = {
                    'profit_by_category': category_margins['profit'].to_dict(),
                    'margin_percent_by_category': category_margins['margin_percent'].to_dict(),
                    'most_profitable_category': category_margins['profit'].idxmax(),
                    'highest_margin_category': category_margins['margin_percent'].idxmax()
                }
        
        return margin_analysis
    
    def compare_business_performance(self, other_analyzer: 'BusinessIntelligenceAnalyzer') -> Dict[str, Any]:
        """Compare business performance between two datasets"""
        
        comparison = {
            'dataset_comparison': self._compare_dataset_basics(other_analyzer),
            'revenue_comparison': self._compare_revenue_performance(other_analyzer),
            'customer_comparison': self._compare_customer_metrics(other_analyzer),
            'product_comparison': self._compare_product_performance(other_analyzer),
            'performance_summary': self._generate_performance_summary(other_analyzer)
        }
        
        return comparison
    
    def _compare_dataset_basics(self, other_analyzer: 'BusinessIntelligenceAnalyzer') -> Dict[str, Any]:
        """Compare basic dataset characteristics"""
        
        return {
            'dataset_a': {
                'total_rows': self.total_rows,
                'total_columns': len(self.columns),
                'amount_columns': len(self.amount_columns),
                'category_columns': len(self.category_columns)
            },
            'dataset_b': {
                'total_rows': other_analyzer.total_rows,
                'total_columns': len(other_analyzer.columns),
                'amount_columns': len(other_analyzer.amount_columns),
                'category_columns': len(other_analyzer.category_columns)
            },
            'differences': {
                'row_difference': other_analyzer.total_rows - self.total_rows,
                'column_difference': len(other_analyzer.columns) - len(self.columns),
                'growth_rate': ((other_analyzer.total_rows - self.total_rows) / self.total_rows * 100) if self.total_rows > 0 else 0
            }
        }
    
    def _compare_revenue_performance(self, other_analyzer: 'BusinessIntelligenceAnalyzer') -> Dict[str, Any]:
        """Compare revenue performance between datasets"""
        
        if not self.amount_columns or not other_analyzer.amount_columns:
            return {'message': 'Both datasets need amount columns for revenue comparison'}
        
        amount_col_a = self.amount_columns[0]
        amount_col_b = other_analyzer.amount_columns[0]
        
        revenue_a = self.df[amount_col_a].sum()
        revenue_b = other_analyzer.df[amount_col_b].sum()
        
        avg_transaction_a = self.df[amount_col_a].mean()
        avg_transaction_b = other_analyzer.df[amount_col_b].mean()
        
        return {
            'total_revenue': {
                'dataset_a': float(revenue_a),
                'dataset_b': float(revenue_b),
                'difference': float(revenue_b - revenue_a),
                'growth_rate': float((revenue_b - revenue_a) / revenue_a * 100) if revenue_a > 0 else 0
            },
            'average_transaction': {
                'dataset_a': float(avg_transaction_a),
                'dataset_b': float(avg_transaction_b),
                'difference': float(avg_transaction_b - avg_transaction_a),
                'improvement': float((avg_transaction_b - avg_transaction_a) / avg_transaction_a * 100) if avg_transaction_a > 0 else 0
            }
        }
    
    def _compare_customer_metrics(self, other_analyzer: 'BusinessIntelligenceAnalyzer') -> Dict[str, Any]:
        """Compare customer-related metrics"""
        
        # Find customer columns in both datasets
        customer_cols_a = [col for col in self.columns if 'customer' in str(col).lower()]
        customer_cols_b = [col for col in other_analyzer.columns if 'customer' in str(col).lower()]
        
        if not customer_cols_a or not customer_cols_b:
            return {'message': 'Customer columns not available in both datasets'}
        
        customer_col_a = customer_cols_a[0]
        customer_col_b = customer_cols_b[0]
        
        unique_customers_a = self.df[customer_col_a].nunique()
        unique_customers_b = other_analyzer.df[customer_col_b].nunique()
        
        return {
            'customer_count': {
                'dataset_a': unique_customers_a,
                'dataset_b': unique_customers_b,
                'difference': unique_customers_b - unique_customers_a,
                'growth_rate': ((unique_customers_b - unique_customers_a) / unique_customers_a * 100) if unique_customers_a > 0 else 0
            }
        }
    
    def _compare_product_performance(self, other_analyzer: 'BusinessIntelligenceAnalyzer') -> Dict[str, Any]:
        """Compare product performance between datasets"""
        
        product_cols_a = [col for col in self.columns if 'product' in str(col).lower()]
        product_cols_b = [col for col in other_analyzer.columns if 'product' in col.lower()]
        
        if not product_cols_a or not product_cols_b:
            return {'message': 'Product columns not available in both datasets'}
        
        product_col_a = product_cols_a[0]
        product_col_b = product_cols_b[0]
        
        unique_products_a = self.df[product_col_a].nunique()
        unique_products_b = other_analyzer.df[product_col_b].nunique()
        
        return {
            'product_diversity': {
                'dataset_a': unique_products_a,
                'dataset_b': unique_products_b,
                'difference': unique_products_b - unique_products_a
            }
        }
    
    def _generate_performance_summary(self, other_analyzer: 'BusinessIntelligenceAnalyzer') -> List[str]:
        """Generate performance comparison summary"""
        
        summary = []
        
        # Dataset size comparison
        if other_analyzer.total_rows > self.total_rows:
            growth = ((other_analyzer.total_rows - self.total_rows) / self.total_rows * 100)
            summary.append(f"📈 Dataset B has {growth:.1f}% more records than Dataset A")
        elif other_analyzer.total_rows < self.total_rows:
            decline = ((self.total_rows - other_analyzer.total_rows) / self.total_rows * 100)
            summary.append(f"📉 Dataset B has {decline:.1f}% fewer records than Dataset A")
        
        # Revenue comparison
        if self.amount_columns and other_analyzer.amount_columns:
            revenue_a = self.df[self.amount_columns[0]].sum()
            revenue_b = other_analyzer.df[other_analyzer.amount_columns[0]].sum()
            
            if revenue_b > revenue_a:
                growth = ((revenue_b - revenue_a) / revenue_a * 100)
                summary.append(f"💰 Revenue increased by {growth:.1f}% in Dataset B")
            elif revenue_b < revenue_a:
                decline = ((revenue_a - revenue_b) / revenue_a * 100)
                summary.append(f"📉 Revenue decreased by {decline:.1f}% in Dataset B")
        
        # Business dimensions comparison
        biz_dims_a = len(self.category_columns) + len(self.amount_columns)
        biz_dims_b = len(other_analyzer.category_columns) + len(other_analyzer.amount_columns)
        
        if biz_dims_b > biz_dims_a:
            summary.append(f"🎯 Dataset B has more business dimensions ({biz_dims_b} vs {biz_dims_a})")
        
        return summary[:5]  # Limit to top 5 insights
    
    def analyze_advanced_business_kpis(self) -> Dict[str, Any]:
        """Calculate advanced business KPIs from RFP Bayanati specifications"""
        
        kpis = {
            'sales_marketing_kpis': self._calculate_sales_marketing_kpis(),
            'operational_kpis': self._calculate_operational_kpis(),
            'hr_kpis': self._calculate_hr_kpis(),
            'project_management_kpis': self._calculate_project_kpis(),
            'banking_cash_kpis': self._calculate_banking_kpis(),
            'forecasting_ai_metrics': self._calculate_forecasting_metrics(),
            'user_analytics': self._calculate_user_analytics(),
            'governance_metrics': self._calculate_governance_metrics()
        }
        
        return kpis
    
    def _calculate_sales_marketing_kpis(self) -> Dict[str, Any]:
        """Sales & Marketing KPIs from RFP specifications"""
        kpis = {}
        
        try:
            # Sales Growth %
            if self.date_columns and self.amount_columns:
                date_col = self.date_columns[0]
                amount_col = self.amount_columns[0]
                
                df_temporal = self.df[[date_col, amount_col]].dropna()
                df_temporal[date_col] = pd.to_datetime(df_temporal[date_col])
                df_temporal = df_temporal.sort_values(date_col)
                
                if len(df_temporal) > 1:
                    # Split into periods
                    mid_point = len(df_temporal) // 2
                    recent_sales = df_temporal.tail(mid_point)[amount_col].sum()
                    previous_sales = df_temporal.head(mid_point)[amount_col].sum()
                    
                    if previous_sales > 0:
                        sales_growth = (recent_sales - previous_sales) / previous_sales * 100
                        kpis['sales_growth_percent'] = {
                            'value': round(sales_growth, 2),
                            'formula': '(Recent Sales - Previous Sales) ÷ Previous Sales × 100',
                            'interpretation': 'Excellent' if sales_growth >= 20 else 'Good' if sales_growth >= 10 else 'Concerning' if sales_growth < 0 else 'Stable',
                            'benchmark': 'Target: 15-25% annual growth'
                        }
            
            # Customer Retention Rate (estimated)
            customer_cols = self._detect_customer_columns()
            if customer_cols and self.date_columns:
                customer_col = customer_cols[0]
                date_col = self.date_columns[0]
                
                df_customer = self.df[[customer_col, date_col]].dropna()
                df_customer[date_col] = pd.to_datetime(df_customer[date_col])
                
                # Estimate retention based on repeat customers
                total_customers = df_customer[customer_col].nunique()
                repeat_customers = df_customer[customer_col].value_counts()
                repeat_count = (repeat_customers > 1).sum()
                
                retention_rate = repeat_count / total_customers * 100 if total_customers > 0 else 0
                kpis['customer_retention_rate'] = {
                    'value': round(retention_rate, 2),
                    'formula': 'Repeat Customers ÷ Total Customers × 100',
                    'interpretation': 'Excellent' if retention_rate >= 80 else 'Good' if retention_rate >= 60 else 'Needs Improvement',
                    'benchmark': 'Target: 80%+'
                }
                
                # Customer Churn Rate
                churn_rate = 100 - retention_rate
                kpis['customer_churn_rate'] = {
                    'value': round(churn_rate, 2),
                    'formula': '100% - Retention Rate',
                    'interpretation': 'Good' if churn_rate <= 20 else 'Concerning' if churn_rate >= 40 else 'Moderate'
                }
            
            # Customer Acquisition Cost (CAC) - estimated
            marketing_cols = self._detect_column_by_keywords(self.df, ['marketing', 'advertising', 'promotion', 'campaign'])
            if marketing_cols and customer_cols:
                total_marketing_spend = self.df[marketing_cols].sum()
                new_customers = self.df[customer_cols[0]].nunique()
                
                cac = total_marketing_spend / new_customers if new_customers > 0 else 0
                kpis['customer_acquisition_cost'] = {
                    'value': round(cac, 2),
                    'formula': 'Total Marketing Spend ÷ New Customers',
                    'unit': '$',
                    'interpretation': 'Efficient' if cac <= 100 else 'Expensive' if cac >= 500 else 'Moderate'
                }
            
            # Lifetime Value (LTV) - enhanced estimation
            if self.amount_columns and customer_cols:
                amount_col = self.amount_columns[0]
                customer_col = customer_cols[0]
                
                customer_spending = self.df.groupby(customer_col)[amount_col].agg(['sum', 'count', 'mean'])
                avg_customer_value = customer_spending['sum'].mean()
                avg_purchase_frequency = customer_spending['count'].mean()
                avg_order_value = customer_spending['mean'].mean()
                
                # Estimate LTV (simplified)
                estimated_ltv = avg_customer_value * (avg_purchase_frequency / 12)  # Annualized
                
                kpis['customer_lifetime_value'] = {
                    'value': round(estimated_ltv, 2),
                    'formula': 'Avg Customer Value × Purchase Frequency',
                    'unit': '$',
                    'interpretation': 'High Value' if estimated_ltv >= 1000 else 'Medium Value' if estimated_ltv >= 500 else 'Low Value'
                }
                
        except Exception as e:
            kpis['error'] = f"Sales/Marketing KPI calculation error: {str(e)}"
        
        return kpis
    
    def _calculate_operational_kpis(self) -> Dict[str, Any]:
        """Operational KPIs"""
        kpis = {}
        
        try:
            # Order Fulfillment Cycle Time
            delivery_cols = self._detect_column_by_keywords(self.df, ['delivery', 'shipped', 'fulfilled', 'completed'])
            order_cols = self._detect_column_by_keywords(self.df, ['order', 'created', 'placed'])
            
            if delivery_cols and order_cols and len(self.date_columns) >= 2:
                try:
                    delivery_dates = pd.to_datetime(self.df[delivery_cols])
                    order_dates = pd.to_datetime(self.df[order_cols])
                    cycle_times = (delivery_dates - order_dates).dt.days
                    
                    avg_cycle_time = cycle_times.mean()
                    kpis['order_fulfillment_cycle_time'] = {
                        'value': round(avg_cycle_time, 1),
                        'formula': 'Avg(Delivery Date - Order Date)',
                        'unit': 'days',
                        'interpretation': 'Excellent' if avg_cycle_time <= 2 else 'Good' if avg_cycle_time <= 5 else 'Needs Improvement'
                    }
                except:
                    pass
            
            # On-Time Delivery % (if we have status columns)
            status_cols = self._detect_column_by_keywords(self.df, ['status', 'delivered', 'on_time'])
            if status_cols:
                status_col = status_cols[0]
                total_orders = len(self.df[status_col].dropna())
                on_time_orders = self.df[status_col].str.contains('on.time|delivered|success', case=False, na=False).sum()
                
                on_time_percentage = on_time_orders / total_orders * 100 if total_orders > 0 else 0
                kpis['on_time_delivery_rate'] = {
                    'value': round(on_time_percentage, 2),
                    'formula': 'On-time Deliveries ÷ Total Deliveries × 100',
                    'unit': '%',
                    'interpretation': 'Excellent' if on_time_percentage >= 95 else 'Good' if on_time_percentage >= 85 else 'Needs Improvement'
                }
            
        except Exception as e:
            kpis['error'] = f"Operational KPI calculation error: {str(e)}"
        
        return kpis
    
    def _calculate_hr_kpis(self) -> Dict[str, Any]:
        """HR KPIs"""
        kpis = {}
        
        try:
            # Employee Turnover % (if we have employee data)
            employee_cols = self._detect_column_by_keywords(self.df, ['employee', 'staff', 'worker', 'person'])
            if employee_cols:
                total_employees = self.df[employee_cols[0]].nunique()
                
                # Look for termination/departure indicators
                departure_cols = self._detect_column_by_keywords(self.df, ['terminated', 'left', 'departure', 'end_date'])
                if departure_cols:
                    departed_employees = self.df[departure_cols[0]].notna().sum()
                    turnover_rate = departed_employees / total_employees * 100 if total_employees > 0 else 0
                    
                    kpis['employee_turnover_rate'] = {
                        'value': round(turnover_rate, 2),
                        'formula': 'Departed Employees ÷ Total Employees × 100',
                        'unit': '%',
                        'interpretation': 'Good' if turnover_rate <= 10 else 'Concerning' if turnover_rate >= 20 else 'Moderate'
                    }
            
            # Training ROI (if we have training cost and performance data)
            training_cols = self._detect_column_by_keywords(self.df, ['training', 'development', 'course'])
            performance_cols = self._detect_column_by_keywords(self.df, ['performance', 'rating', 'score'])
            
            if training_cols and performance_cols:
                training_cost = self.df[training_cols[0]].sum() if pd.api.types.is_numeric_dtype(self.df[training_cols[0]]) else 0
                avg_performance = self.df[performance_cols[0]].mean() if pd.api.types.is_numeric_dtype(self.df[performance_cols[0]]) else 0
                
                if training_cost > 0:
                    training_roi = (avg_performance * 1000 - training_cost) / training_cost * 100  # Simplified ROI calculation
                    kpis['training_roi'] = {
                        'value': round(training_roi, 2),
                        'formula': '(Performance Gain - Training Cost) ÷ Training Cost × 100',
                        'unit': '%',
                        'interpretation': 'Excellent' if training_roi >= 200 else 'Good' if training_roi >= 100 else 'Poor'
                    }
            
        except Exception as e:
            kpis['error'] = f"HR KPI calculation error: {str(e)}"
        
        return kpis
    
    def _calculate_project_kpis(self) -> Dict[str, Any]:
        """Project Management KPIs"""
        kpis = {}
        
        try:
            # Cost Performance Index (CPI)
            actual_cost_cols = self._detect_column_by_keywords(self.df, ['actual_cost', 'spent', 'cost'])
            planned_cost_cols = self._detect_column_by_keywords(self.df, ['planned_cost', 'budget', 'estimated'])
            
            if actual_cost_cols and planned_cost_cols:
                actual_cost = self.df[actual_cost_cols[0]].sum()
                planned_cost = self.df[planned_cost_cols[0]].sum()
                
                cpi = planned_cost / actual_cost if actual_cost > 0 else 0
                kpis['cost_performance_index'] = {
                    'value': round(cpi, 2),
                    'formula': 'Planned Cost ÷ Actual Cost',
                    'interpretation': 'Under Budget' if cpi > 1 else 'Over Budget' if cpi < 1 else 'On Budget',
                    'benchmark': 'Target: ≥ 1.0'
                }
            
            # Schedule Performance Index (SPI) - if we have timeline data
            if self.date_columns and len(self.date_columns) >= 2:
                planned_date_col = None
                actual_date_col = None
                
                for col in self.date_columns:
                    if 'planned' in col.lower() or 'scheduled' in col.lower():
                        planned_date_col = col
                    elif 'actual' in col.lower() or 'completed' in col.lower():
                        actual_date_col = col
                
                if planned_date_col and actual_date_col:
                    planned_dates = pd.to_datetime(self.df[planned_date_col])
                    actual_dates = pd.to_datetime(self.df[actual_date_col])
                    
                    avg_planned_duration = (planned_dates.max() - planned_dates.min()).days
                    avg_actual_duration = (actual_dates.max() - actual_dates.min()).days
                    
                    spi = avg_planned_duration / avg_actual_duration if avg_actual_duration > 0 else 0
                    kpis['schedule_performance_index'] = {
                        'value': round(spi, 2),
                        'formula': 'Planned Duration ÷ Actual Duration',
                        'interpretation': 'Ahead of Schedule' if spi > 1 else 'Behind Schedule' if spi < 1 else 'On Schedule'
                    }
            
        except Exception as e:
            kpis['error'] = f"Project KPI calculation error: {str(e)}"
        
        return kpis
    
    def _calculate_banking_kpis(self) -> Dict[str, Any]:
        """Banking & Cash Flow KPIs"""
        kpis = {}
        
        try:
            # Cash Conversion Cycle (DSO + DIO - DPO)
            receivables_cols = self._detect_column_by_keywords(self.df, ['receivables', 'ar', 'outstanding'])
            inventory_cols = self._detect_column_by_keywords(self.df, ['inventory', 'stock'])
            payables_cols = self._detect_column_by_keywords(self.df, ['payables', 'ap', 'owed'])
            sales_cols = self.amount_columns
            
            if receivables_cols and sales_cols:
                # Days Sales Outstanding (DSO)
                avg_receivables = self.df[receivables_cols[0]].mean()
                daily_sales = self.df[sales_cols[0]].sum() / 365
                dso = avg_receivables / daily_sales if daily_sales > 0 else 0
                
                kpis['days_sales_outstanding'] = {
                    'value': round(dso, 1),
                    'formula': 'Average Receivables ÷ Daily Sales',
                    'unit': 'days',
                    'interpretation': 'Excellent' if dso <= 30 else 'Good' if dso <= 45 else 'Concerning'
                }
            
            # Debt-to-Equity Ratio
            debt_cols = self._detect_column_by_keywords(self.df, ['debt', 'liabilities', 'loan'])
            equity_cols = self._detect_column_by_keywords(self.df, ['equity', 'capital'])
            
            if debt_cols and equity_cols:
                total_debt = self.df[debt_cols[0]].sum()
                total_equity = self.df[equity_cols[0]].sum()
                
                debt_to_equity = total_debt / total_equity if total_equity > 0 else 0
                kpis['debt_to_equity_ratio'] = {
                    'value': round(debt_to_equity, 2),
                    'formula': 'Total Debt ÷ Total Equity',
                    'interpretation': 'Conservative' if debt_to_equity <= 0.5 else 'Moderate' if debt_to_equity <= 1.0 else 'Aggressive',
                    'benchmark': 'Industry dependent, typically < 1.0'
                }
            
        except Exception as e:
            kpis['error'] = f"Banking KPI calculation error: {str(e)}"
        
        return kpis
    
    def _calculate_forecasting_metrics(self) -> Dict[str, Any]:
        """Forecasting & AI Metrics"""
        metrics = {}
        
        try:
            # Trend Prediction using simple linear regression
            if self.date_columns and self.amount_columns:
                date_col = self.date_columns[0]
                amount_col = self.amount_columns[0]
                
                df_trend = self.df[[date_col, amount_col]].dropna()
                df_trend[date_col] = pd.to_datetime(df_trend[date_col])
                df_trend = df_trend.sort_values(date_col)
                
                if len(df_trend) >= 3:
                    # Simple trend calculation
                    df_trend['days'] = (df_trend[date_col] - df_trend[date_col].min()).dt.days
                    
                    # Calculate trend slope
                    x = df_trend['days'].values
                    y = df_trend[amount_col].values
                    
                    if len(x) > 1:
                        slope = np.polyfit(x, y, 1)[0]
                        trend_percentage = slope / df_trend[amount_col].mean() * 100 if df_trend[amount_col].mean() != 0 else 0
                        
                        metrics['trend_prediction'] = {
                            'daily_trend': round(slope, 2),
                            'trend_percentage': round(trend_percentage, 2),
                            'interpretation': 'Growing' if slope > 0 else 'Declining' if slope < 0 else 'Stable',
                            'confidence': 'Basic' if len(df_trend) < 10 else 'Moderate' if len(df_trend) < 30 else 'Good'
                        }
            
            # Anomaly Detection
            if self.amount_columns:
                amount_col = self.amount_columns[0]
                amounts = self.df[amount_col].dropna()
                
                if len(amounts) > 5:
                    mean_val = amounts.mean()
                    std_val = amounts.std()
                    
                    # Identify outliers (values beyond 2 standard deviations)
                    outliers = amounts[(amounts > mean_val + 2*std_val) | (amounts < mean_val - 2*std_val)]
                    anomaly_rate = len(outliers) / len(amounts) * 100
                    
                    metrics['anomaly_detection'] = {
                        'anomaly_rate': round(anomaly_rate, 2),
                        'anomaly_count': len(outliers),
                        'threshold_upper': round(mean_val + 2*std_val, 2),
                        'threshold_lower': round(mean_val - 2*std_val, 2),
                        'interpretation': 'High Variance' if anomaly_rate > 10 else 'Normal Variance' if anomaly_rate > 5 else 'Low Variance'
                    }
            
        except Exception as e:
            metrics['error'] = f"Forecasting metrics calculation error: {str(e)}"
        
        return metrics
    
    def _calculate_user_analytics(self) -> Dict[str, Any]:
        """User Analytics & Engagement Metrics"""
        analytics = {}
        
        try:
            # Active Users (if we have user/customer data)
            user_cols = self._detect_column_by_keywords(self.df, ['user', 'customer', 'client', 'account'])
            if user_cols:
                active_users = self.df[user_cols[0]].nunique()
                analytics['active_users'] = {
                    'count': active_users,
                    'interpretation': 'High Engagement' if active_users >= 1000 else 'Medium Engagement' if active_users >= 100 else 'Low Engagement'
                }
            
            # Engagement Metrics (clicks, sessions, etc.)
            engagement_cols = self._detect_column_by_keywords(self.df, ['clicks', 'views', 'sessions', 'interactions'])
            if engagement_cols:
                total_engagement = self.df[engagement_cols[0]].sum()
                avg_engagement = self.df[engagement_cols[0]].mean()
                
                analytics['engagement_metrics'] = {
                    'total_interactions': int(total_engagement),
                    'avg_interactions_per_record': round(avg_engagement, 2),
                    'interpretation': 'High Activity' if avg_engagement >= 10 else 'Moderate Activity' if avg_engagement >= 5 else 'Low Activity'
                }
            
            # Session Duration (if available)
            duration_cols = self._detect_column_by_keywords(self.df, ['duration', 'time_spent', 'session_time'])
            if duration_cols:
                avg_duration = self.df[duration_cols[0]].mean()
                analytics['session_duration'] = {
                    'average_minutes': round(avg_duration, 2),
                    'interpretation': 'Highly Engaged' if avg_duration >= 15 else 'Engaged' if avg_duration >= 5 else 'Brief Sessions'
                }
            
        except Exception as e:
            analytics['error'] = f"User analytics calculation error: {str(e)}"
        
        return analytics
    
    def _calculate_governance_metrics(self) -> Dict[str, Any]:
        """Governance & Data Quality Metrics"""
        governance = {}
        
        try:
            # Data Quality Score
            total_cells = self.total_rows * len(self.columns)
            valid_cells = total_cells - self.df.isnull().sum().sum()
            data_quality_score = valid_cells / total_cells * 100 if total_cells > 0 else 0
            
            governance['data_quality_score'] = {
                'score': round(data_quality_score, 2),
                'formula': 'Valid Cells ÷ Total Cells × 100',
                'interpretation': 'Excellent' if data_quality_score >= 95 else 'Good' if data_quality_score >= 85 else 'Needs Improvement',
                'benchmark': 'Target: ≥ 95%'
            }
            
            # Audit Trail Metrics (if we have tracking columns)
            audit_cols = self._detect_column_by_keywords(self.df, ['created_by', 'updated_by', 'modified', 'audit'])
            if audit_cols:
                tracked_records = self.df[audit_cols[0]].notna().sum()
                audit_coverage = tracked_records / self.total_rows * 100
                
                governance['audit_coverage'] = {
                    'coverage_percentage': round(audit_coverage, 2),
                    'tracked_records': int(tracked_records),
                    'interpretation': 'Full Compliance' if audit_coverage >= 95 else 'Partial Compliance' if audit_coverage >= 80 else 'Poor Compliance'
                }
            
            # Subscription Revenue (MRR) - if available
            subscription_cols = self._detect_column_by_keywords(self.df, ['subscription', 'recurring', 'monthly'])
            if subscription_cols and self.amount_columns:
                monthly_revenue = self.df[subscription_cols[0]].sum() if pd.api.types.is_numeric_dtype(self.df[subscription_cols[0]]) else 0
                
                governance['monthly_recurring_revenue'] = {
                    'mrr': round(monthly_revenue, 2),
                    'unit': '$',
                    'interpretation': 'Strong MRR' if monthly_revenue >= 50000 else 'Growing MRR' if monthly_revenue >= 10000 else 'Early Stage'
                }
            
        except Exception as e:
            governance['error'] = f"Governance metrics calculation error: {str(e)}"
        
        return governance
    
    def analyze_benchmarking_alerts(self) -> Dict[str, Any]:
        """Cross-Institution Benchmarking and Early Warning System"""
        
        benchmarking = {
            'early_warning_indicators': self._generate_early_warnings(),
            'performance_benchmarks': self._calculate_performance_benchmarks(),
            'sector_analysis': self._analyze_sector_performance(),
            'risk_assessment': self._assess_business_risks(),
            'recommendations': []
        }
        
        return benchmarking
    
    def _generate_early_warnings(self) -> Dict[str, Any]:
        """Generate early warning indicators based on thresholds"""
        warnings = {
            'critical_alerts': [],
            'warning_alerts': [],
            'informational_alerts': [],
            'threshold_checks': {}
        }
        
        try:
            # Liquidity Warning (Threshold: < 1.0)
            financial_ratios = self._calculate_financial_ratios()
            if 'liquidity_ratios' in financial_ratios:
                if 'current_ratio' in financial_ratios['liquidity_ratios']:
                    current_ratio = financial_ratios['liquidity_ratios']['current_ratio']['value']
                    if current_ratio < 1.0:
                        warnings['critical_alerts'].append({
                            'type': 'Liquidity Crisis',
                            'message': f'Current ratio of {current_ratio} is below safe threshold of 1.0',
                            'recommendation': 'Immediate cash flow review required',
                            'severity': 'CRITICAL'
                        })
                    elif current_ratio < 1.5:
                        warnings['warning_alerts'].append({
                            'type': 'Liquidity Warning',
                            'message': f'Current ratio of {current_ratio} approaching low threshold',
                            'recommendation': 'Monitor cash flow closely',
                            'severity': 'WARNING'
                        })
            
            # Revenue Trend Warning
            if self.amount_columns and self.date_columns:
                amount_col = self.amount_columns[0]
                date_col = self.date_columns[0]
                
                df_trend = self.df[[date_col, amount_col]].dropna()
                if len(df_trend) > 2:
                    df_trend[date_col] = pd.to_datetime(df_trend[date_col])
                    df_trend = df_trend.sort_values(date_col)
                    
                    recent_avg = df_trend.tail(int(len(df_trend)/3))[amount_col].mean()
                    earlier_avg = df_trend.head(int(len(df_trend)/3))[amount_col].mean()
                    
                    if earlier_avg > 0:
                        revenue_change = (recent_avg - earlier_avg) / earlier_avg * 100
                        
                        if revenue_change < -20:
                            warnings['critical_alerts'].append({
                                'type': 'Revenue Decline',
                                'message': f'Revenue declined by {abs(revenue_change):.1f}%',
                                'recommendation': 'Urgent business review required',
                                'severity': 'CRITICAL'
                            })
                        elif revenue_change < -10:
                            warnings['warning_alerts'].append({
                                'type': 'Revenue Softening',
                                'message': f'Revenue declined by {abs(revenue_change):.1f}%',
                                'recommendation': 'Investigate market conditions',
                                'severity': 'WARNING'
                            })
            
            # Anomaly Detection Warnings
            if self.amount_columns:
                amount_col = self.amount_columns[0]
                amounts = self.df[amount_col].dropna()
                
                if len(amounts) > 10:
                    mean_val = amounts.mean()
                    std_val = amounts.std()
                    outliers = amounts[(amounts > mean_val + 3*std_val) | (amounts < mean_val - 3*std_val)]
                    
                    if len(outliers) > 0:
                        warnings['warning_alerts'].append({
                            'type': 'Data Anomalies Detected',
                            'message': f'{len(outliers)} extreme outliers found in transaction data',
                            'recommendation': 'Investigate unusual transactions',
                            'severity': 'WARNING'
                        })
            
            # Customer Concentration Risk
            customer_cols = self._detect_customer_columns()
            if customer_cols and self.amount_columns:
                customer_col = customer_cols[0]
                amount_col = self.amount_columns[0]
                
                customer_revenue = self.df.groupby(customer_col)[amount_col].sum()
                total_revenue = customer_revenue.sum()
                
                if total_revenue > 0:
                    top_customer_pct = customer_revenue.max() / total_revenue * 100
                    
                    if top_customer_pct > 50:
                        warnings['critical_alerts'].append({
                            'type': 'Customer Concentration Risk',
                            'message': f'Top customer represents {top_customer_pct:.1f}% of revenue',
                            'recommendation': 'Diversify customer base urgently',
                            'severity': 'CRITICAL'
                        })
                    elif top_customer_pct > 30:
                        warnings['warning_alerts'].append({
                            'type': 'High Customer Dependency',
                            'message': f'Top customer represents {top_customer_pct:.1f}% of revenue',
                            'recommendation': 'Consider customer diversification',
                            'severity': 'WARNING'
                        })
            
        except Exception as e:
            warnings['error'] = f"Early warning generation error: {str(e)}"
        
        return warnings
    
    def _calculate_performance_benchmarks(self) -> Dict[str, Any]:
        """Calculate performance benchmarks for comparison"""
        benchmarks = {}
        
        try:
            # Revenue per record benchmark
            if self.amount_columns:
                amount_col = self.amount_columns[0]
                revenue_per_record = self.df[amount_col].mean()
                
                benchmarks['revenue_efficiency'] = {
                    'revenue_per_record': round(revenue_per_record, 2),
                    'benchmark_tier': 'High Performer' if revenue_per_record >= 1000 else 'Average Performer' if revenue_per_record >= 500 else 'Below Average',
                    'industry_comparison': 'Above median' if revenue_per_record >= 750 else 'Below median'
                }
            
            # Data completeness benchmark
            completeness_score = (1 - self.df.isnull().sum().sum() / (len(self.df) * len(self.df.columns))) * 100
            benchmarks['data_quality'] = {
                'completeness_score': round(completeness_score, 2),
                'benchmark_tier': 'Excellent' if completeness_score >= 95 else 'Good' if completeness_score >= 85 else 'Needs Improvement',
                'compliance_level': 'Enterprise Grade' if completeness_score >= 95 else 'Business Grade' if completeness_score >= 85 else 'Basic Grade'
            }
            
            # Business complexity benchmark
            business_dimensions = len(self.category_columns) + len(self.amount_columns) + len(self.date_columns)
            benchmarks['business_complexity'] = {
                'dimension_count': business_dimensions,
                'complexity_tier': 'Enterprise' if business_dimensions >= 10 else 'Mid-Market' if business_dimensions >= 5 else 'Small Business',
                'analytics_readiness': 'Advanced' if business_dimensions >= 8 else 'Intermediate' if business_dimensions >= 4 else 'Basic'
            }
            
        except Exception as e:
            benchmarks['error'] = f"Benchmark calculation error: {str(e)}"
        
        return benchmarks
    
    def generate_executive_summary(self) -> Dict[str, Any]:
        """Generate comprehensive executive summary with key insights and recommendations"""
        summary = {
            'overview': {},
            'key_findings': [],
            'financial_highlights': {},
            'risk_assessment': {},
            'recommendations': [],
            'performance_scorecard': {}
        }
        
        try:
            # Business Overview
            summary['overview'] = {
                'dataset_size': f"{len(self.df):,} records",
                'data_completeness': f"{(1 - self.df.isnull().sum().sum() / (len(self.df) * len(self.df.columns))) * 100:.1f}%",
                'business_dimensions': len(self.category_columns) + len(self.amount_columns) + len(self.date_columns),
                'analysis_date': datetime.now().strftime('%Y-%m-%d %H:%M'),
                'data_quality_grade': 'A' if (1 - self.df.isnull().sum().sum() / (len(self.df) * len(self.df.columns))) >= 0.95 else 'B' if (1 - self.df.isnull().sum().sum() / (len(self.df) * len(self.df.columns))) >= 0.85 else 'C'
            }
            
            # Key Financial Findings
            if self.amount_columns:
                amount_col = self.amount_columns[0]
                total_value = self.df[amount_col].sum()
                avg_value = self.df[amount_col].mean()
                
                summary['financial_highlights'] = {
                    'total_value': f"${total_value:,.2f}",
                    'average_transaction': f"${avg_value:,.2f}",
                    'value_concentration': self._assess_value_concentration(),
                    'growth_indicators': self._calculate_growth_indicators()
                }
                
                # Key findings based on financial analysis
                if total_value > 1000000:
                    summary['key_findings'].append("💰 High-value dataset with significant financial impact")
                if avg_value > 1000:
                    summary['key_findings'].append("📈 Premium transaction profile indicates quality customer base")
            
            # Customer Insights
            customer_insights = self._analyze_customer_patterns()
            if customer_insights.get('insights'):
                summary['key_findings'].extend(customer_insights['insights'][:3])  # Top 3 customer insights
            
            # Risk Assessment Summary
            early_warnings = self.analyze_benchmarking_alerts()
            risk_level = 'LOW'
            critical_issues = 0
            
            if 'critical_alerts' in early_warnings:
                critical_issues = len(early_warnings['critical_alerts'])
                if critical_issues > 0:
                    risk_level = 'HIGH'
                elif len(early_warnings.get('warning_alerts', [])) > 2:
                    risk_level = 'MEDIUM'
            
            summary['risk_assessment'] = {
                'overall_risk_level': risk_level,
                'critical_issues': critical_issues,
                'monitoring_required': len(early_warnings.get('warning_alerts', [])),
                'risk_score': self._calculate_risk_score(early_warnings)
            }
            
            # Performance Scorecard
            financial_ratios = self._calculate_financial_ratios()
            performance_score = 0
            total_metrics = 0
            
            # Calculate overall performance score
            if 'liquidity_ratios' in financial_ratios:
                for ratio_name, ratio_data in financial_ratios['liquidity_ratios'].items():
                    if isinstance(ratio_data, dict) and 'value' in ratio_data:
                        score = self._score_financial_metric(ratio_name, ratio_data['value'])
                        performance_score += score
                        total_metrics += 1
            
            if 'profitability_ratios' in financial_ratios:
                for ratio_name, ratio_data in financial_ratios['profitability_ratios'].items():
                    if isinstance(ratio_data, dict) and 'value' in ratio_data:
                        score = self._score_financial_metric(ratio_name, ratio_data['value'])
                        performance_score += score
                        total_metrics += 1
            
            overall_score = (performance_score / max(total_metrics, 1)) * 100 if total_metrics > 0 else 0
            
            summary['performance_scorecard'] = {
                'overall_score': round(overall_score, 1),
                'performance_grade': self._get_performance_grade(overall_score),
                'metrics_analyzed': total_metrics,
                'benchmark_comparison': 'Above Industry Average' if overall_score >= 75 else 'Industry Average' if overall_score >= 60 else 'Below Industry Average'
            }
            
            # Strategic Recommendations
            recommendations = self._generate_strategic_recommendations(summary)
            summary['recommendations'] = recommendations
            
        except Exception as e:
            summary['error'] = f"Executive summary generation error: {str(e)}"
        
        return summary
    
    def _assess_value_concentration(self) -> str:
        """Assess concentration of values in amount columns"""
        if not self.amount_columns:
            return "No financial data available"
        
        amount_col = self.amount_columns[0]
        top_20_pct = self.df[amount_col].quantile(0.8)
        top_20_sum = self.df[self.df[amount_col] >= top_20_pct][amount_col].sum()
        total_sum = self.df[amount_col].sum()
        
        concentration = (top_20_sum / total_sum) * 100 if total_sum > 0 else 0
        
        if concentration >= 80:
            return f"High concentration - Top 20% accounts for {concentration:.1f}% of value"
        elif concentration >= 60:
            return f"Moderate concentration - Top 20% accounts for {concentration:.1f}% of value"
        else:
            return f"Well distributed - Top 20% accounts for {concentration:.1f}% of value"
    
    def _calculate_growth_indicators(self) -> Dict[str, Any]:
        """Calculate growth and trend indicators"""
        indicators = {}
        
        try:
            if self.date_columns and self.amount_columns:
                date_col = self.date_columns[0]
                amount_col = self.amount_columns[0]
                
                # Sort by date and calculate trend
                df_sorted = self.df.sort_values(date_col)
                df_sorted['month'] = pd.to_datetime(df_sorted[date_col]).dt.to_period('M')
                monthly_totals = df_sorted.groupby('month')[amount_col].sum()
                
                if len(monthly_totals) >= 2:
                    recent_growth = ((monthly_totals.iloc[-1] - monthly_totals.iloc[-2]) / monthly_totals.iloc[-2]) * 100
                    indicators['recent_growth_rate'] = f"{recent_growth:+.1f}%"
                    indicators['trend_direction'] = 'Increasing' if recent_growth > 5 else 'Stable' if recent_growth > -5 else 'Decreasing'
                else:
                    indicators['trend_direction'] = 'Insufficient data for trend analysis'
            else:
                indicators['trend_direction'] = 'No date/amount columns for trend analysis'
                
        except Exception as e:
            indicators['error'] = f"Growth calculation error: {str(e)}"
        
        return indicators
    
    def _calculate_risk_score(self, early_warnings: Dict) -> int:
        """Calculate overall risk score from 0-100"""
        risk_score = 0
        
        # Critical alerts contribute heavily to risk score
        critical_count = len(early_warnings.get('critical_alerts', []))
        risk_score += critical_count * 30
        
        # Warning alerts contribute moderately
        warning_count = len(early_warnings.get('warning_alerts', []))
        risk_score += warning_count * 15
        
        # Cap at 100
        return min(risk_score, 100)
    
    def _score_financial_metric(self, metric_name: str, value: float) -> int:
        """Score individual financial metrics from 0-100"""
        if metric_name.lower() == 'current_ratio':
            if value >= 2.0:
                return 100
            elif value >= 1.5:
                return 80
            elif value >= 1.0:
                return 60
            else:
                return 30
        elif 'margin' in metric_name.lower():
            if value >= 20:
                return 100
            elif value >= 15:
                return 80
            elif value >= 10:
                return 60
            elif value >= 5:
                return 40
            else:
                return 20
        elif metric_name.lower() in ['roe', 'roa']:
            if value >= 15:
                return 100
            elif value >= 10:
                return 80
            elif value >= 5:
                return 60
            else:
                return 40
        else:
            # Default scoring for unknown metrics
            return 70
    
    def _get_performance_grade(self, score: float) -> str:
        """Convert performance score to letter grade"""
        if score >= 90:
            return 'A+'
        elif score >= 85:
            return 'A'
        elif score >= 80:
            return 'A-'
        elif score >= 75:
            return 'B+'
        elif score >= 70:
            return 'B'
        elif score >= 65:
            return 'B-'
        elif score >= 60:
            return 'C+'
        elif score >= 55:
            return 'C'
        else:
            return 'D'
    
    def _generate_strategic_recommendations(self, summary: Dict) -> List[str]:
        """Generate strategic recommendations based on analysis results"""
        recommendations = []
        
        # Risk-based recommendations
        risk_level = summary.get('risk_assessment', {}).get('overall_risk_level', 'LOW')
        if risk_level == 'HIGH':
            recommendations.append("🚨 Immediate attention required: Address critical risk factors identified in analysis")
        elif risk_level == 'MEDIUM':
            recommendations.append("⚠️ Monitor key risk indicators and implement preventive measures")
        
        # Performance-based recommendations
        performance_score = summary.get('performance_scorecard', {}).get('overall_score', 0)
        if performance_score < 60:
            recommendations.append("📈 Focus on improving key financial ratios and operational efficiency")
        elif performance_score >= 85:
            recommendations.append("🎯 Excellent performance - Consider expansion or investment opportunities")
        
        # Data quality recommendations
        data_grade = summary.get('overview', {}).get('data_quality_grade', 'C')
        if data_grade == 'C':
            recommendations.append("🔍 Improve data quality processes to enhance analysis accuracy")
        
        # General strategic recommendations
        recommendations.extend([
            "📊 Implement regular monitoring dashboards for key performance indicators",
            "🎯 Establish benchmarking processes against industry standards",
            "📋 Create automated reporting for stakeholder communication"
        ])
        
        return recommendations[:5]  # Return top 5 recommendations
    
    def _analyze_sector_performance(self) -> Dict[str, Any]:
        """Analyze performance by sectors/categories"""
        sector_analysis = {}
        
        try:
            if self.category_columns and self.amount_columns:
                category_col = self.category_columns[0]
                amount_col = self.amount_columns[0]
                
                sector_performance = self.df.groupby(category_col)[amount_col].agg(['sum', 'mean', 'count'])
                sector_performance['total_revenue'] = sector_performance['sum']
                sector_performance['avg_transaction'] = sector_performance['mean']
                sector_performance['transaction_count'] = sector_performance['count']
                
                # Calculate market share for each sector
                total_revenue = sector_performance['total_revenue'].sum()
                sector_performance['market_share'] = sector_performance['total_revenue'] / total_revenue * 100
                
                # Rank sectors
                sector_performance = sector_performance.sort_values('total_revenue', ascending=False)
                
                sector_analysis['sector_rankings'] = []
                for idx, (sector, data) in enumerate(sector_performance.head(5).iterrows()):
                    sector_analysis['sector_rankings'].append({
                        'rank': idx + 1,
                        'sector': str(sector),
                        'revenue': round(data['total_revenue'], 2),
                        'market_share': round(data['market_share'], 2),
                        'avg_transaction': round(data['avg_transaction'], 2),
                        'performance_tier': 'Leader' if idx == 0 else 'Strong' if idx < 3 else 'Developing'
                    })
                
                # Identify growth opportunities
                sector_analysis['growth_opportunities'] = []
                bottom_sectors = sector_performance.tail(3)
                for sector, data in bottom_sectors.iterrows():
                    if data['market_share'] < 5:  # Less than 5% market share
                        sector_analysis['growth_opportunities'].append({
                            'sector': str(sector),
                            'current_share': round(data['market_share'], 2),
                            'potential': 'High' if data['avg_transaction'] > sector_performance['avg_transaction'].median() else 'Medium'
                        })
            
        except Exception as e:
            sector_analysis['error'] = f"Sector analysis error: {str(e)}"
        
        return sector_analysis
    
    def _assess_business_risks(self) -> Dict[str, Any]:
        """Assess various business risks"""
        risk_assessment = {
            'risk_factors': [],
            'risk_score': 0,
            'risk_level': 'Low',
            'mitigation_strategies': []
        }
        
        try:
            risk_score = 0
            
            # Revenue concentration risk
            if self.category_columns and self.amount_columns:
                category_col = self.category_columns[0]
                amount_col = self.amount_columns[0]
                
                category_revenue = self.df.groupby(category_col)[amount_col].sum()
                total_revenue = category_revenue.sum()
                
                if total_revenue > 0:
                    top_category_pct = category_revenue.max() / total_revenue * 100
                    
                    if top_category_pct > 70:
                        risk_score += 30
                        risk_assessment['risk_factors'].append({
                            'type': 'Revenue Concentration',
                            'severity': 'High',
                            'description': f'Over-dependence on single category ({top_category_pct:.1f}%)',
                            'impact': 'Business vulnerability to category-specific downturns'
                        })
                        risk_assessment['mitigation_strategies'].append('Diversify product/service portfolio across multiple categories')
            
            # Data quality risk
            missing_data_pct = self.df.isnull().sum().sum() / (len(self.df) * len(self.df.columns)) * 100
            if missing_data_pct > 20:
                risk_score += 20
                risk_assessment['risk_factors'].append({
                    'type': 'Data Quality',
                    'severity': 'Medium' if missing_data_pct < 40 else 'High',
                    'description': f'High missing data rate ({missing_data_pct:.1f}%)',
                    'impact': 'Unreliable analytics and decision-making'
                })
                risk_assessment['mitigation_strategies'].append('Implement data quality controls and validation processes')
            
            # Operational complexity risk
            if len(self.columns) > 50:
                risk_score += 10
                risk_assessment['risk_factors'].append({
                    'type': 'Operational Complexity',
                    'severity': 'Medium',
                    'description': f'High number of data dimensions ({len(self.columns)} columns)',
                    'impact': 'Increased operational overhead and maintenance costs'
                })
                risk_assessment['mitigation_strategies'].append('Simplify data model and focus on key business metrics')
            
            # Determine overall risk level
            if risk_score >= 50:
                risk_assessment['risk_level'] = 'High'
            elif risk_score >= 25:
                risk_assessment['risk_level'] = 'Medium'
            else:
                risk_assessment['risk_level'] = 'Low'
            
            risk_assessment['risk_score'] = risk_score
            
        except Exception as e:
            risk_assessment['error'] = f"Risk assessment error: {str(e)}"
        
        return risk_assessment