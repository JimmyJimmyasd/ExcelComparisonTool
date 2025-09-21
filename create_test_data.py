import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

# Set random seed for reproducible data
np.random.seed(42)
random.seed(42)

def create_comprehensive_test_data():
    """Create comprehensive test data for RFP Bayanati Business Intelligence Analysis"""
    
    # Generate 1000 rows of test data
    n_rows = 1000
    
    # Date range for the past 2 years
    start_date = datetime.now() - timedelta(days=730)
    dates = [start_date + timedelta(days=i) for i in range(n_rows)]
    
    # Customer data
    customers = [f"Customer_{i:04d}" for i in range(1, 201)]  # 200 unique customers
    
    # Product data
    products = ['Product_A', 'Product_B', 'Product_C', 'Product_D', 'Product_E', 
               'Service_Alpha', 'Service_Beta', 'Service_Gamma']
    
    # Regions and departments
    regions = ['North', 'South', 'East', 'West', 'Central']
    departments = ['Sales', 'Marketing', 'Operations', 'IT', 'HR', 'Finance']
    
    # Sales representatives
    sales_reps = [f"Rep_{i:02d}" for i in range(1, 21)]  # 20 sales reps
    
    # Generate comprehensive dataset
    data = []
    
    for i in range(n_rows):
        # Basic transaction info
        transaction_date = dates[i]
        customer_id = random.choice(customers)
        product_name = random.choice(products)
        region = random.choice(regions)
        department = random.choice(departments)
        sales_rep = random.choice(sales_reps)
        
        # Financial metrics (for financial ratio analysis)
        current_assets = random.uniform(50000, 500000)
        current_liabilities = random.uniform(20000, 300000)
        inventory = random.uniform(5000, 50000)
        cash = random.uniform(10000, 100000)
        total_assets = current_assets + random.uniform(100000, 1000000)
        equity = random.uniform(50000, 500000)
        
        # Sales and revenue data
        quantity = random.randint(1, 100)
        unit_price = random.uniform(10, 1000)
        sales_amount = quantity * unit_price
        discount = random.uniform(0, 0.2) * sales_amount  # 0-20% discount
        net_sales = sales_amount - discount
        
        # Cost data
        cost_of_goods = sales_amount * random.uniform(0.3, 0.7)  # 30-70% COGS
        operating_expenses = sales_amount * random.uniform(0.1, 0.3)  # 10-30% OpEx
        net_income = net_sales - cost_of_goods - operating_expenses
        
        # Customer metrics
        customer_satisfaction = random.uniform(1, 5)  # 1-5 rating
        customer_retention_days = random.randint(30, 365)
        
        # Employee metrics
        employee_count = random.randint(50, 500)
        employee_satisfaction = random.uniform(3, 5)
        training_hours = random.randint(10, 100)
        
        # Marketing metrics
        marketing_spend = random.uniform(1000, 50000)
        leads_generated = random.randint(10, 200)
        conversion_rate = random.uniform(0.05, 0.3)
        
        # Operational metrics
        production_volume = random.randint(100, 5000)
        quality_score = random.uniform(85, 99)
        delivery_time = random.randint(1, 30)  # days
        
        # Banking/Financial ratios (additional columns)
        loan_amount = random.uniform(0, 1000000)
        interest_rate = random.uniform(0.02, 0.15)
        debt_to_equity = random.uniform(0.1, 2.0)
        
        # Project management metrics
        project_budget = random.uniform(10000, 500000)
        project_completion = random.uniform(0.1, 1.0)  # 10-100% completion
        
        # Create row data
        row = {
            # Basic Information
            'Date': transaction_date,
            'Customer_ID': customer_id,
            'Customer_Name': f"Company {customer_id.split('_')[1]}",
            'Product_Name': product_name,
            'Region': region,
            'Department': department,
            'Sales_Rep': sales_rep,
            
            # Financial Statement Data (for ratio analysis)
            'Current_Assets': round(current_assets, 2),
            'Current_Liabilities': round(current_liabilities, 2),
            'Inventory': round(inventory, 2),
            'Cash': round(cash, 2),
            'Total_Assets': round(total_assets, 2),
            'Equity': round(equity, 2),
            'Total_Liabilities': round(current_liabilities + random.uniform(10000, 200000), 2),
            
            # Sales Data
            'Quantity': quantity,
            'Unit_Price': round(unit_price, 2),
            'Sales_Amount': round(sales_amount, 2),
            'Revenue': round(net_sales, 2),
            'Discount': round(discount, 2),
            'Net_Sales': round(net_sales, 2),
            
            # Cost and Profitability
            'Cost_of_Goods_Sold': round(cost_of_goods, 2),
            'COGS': round(cost_of_goods, 2),  # Alternative name
            'Operating_Expenses': round(operating_expenses, 2),
            'Net_Income': round(net_income, 2),
            'Profit': round(net_income, 2),  # Alternative name
            'Gross_Profit': round(net_sales - cost_of_goods, 2),
            
            # Customer Analytics
            'Customer_Satisfaction': round(customer_satisfaction, 1),
            'Customer_Rating': round(customer_satisfaction, 1),
            'Customer_Retention_Days': customer_retention_days,
            'Customer_Segment': random.choice(['Premium', 'Standard', 'Basic']),
            'Customer_Type': random.choice(['B2B', 'B2C']),
            
            # Employee/HR Metrics
            'Employee_Count': employee_count,
            'Employee_Satisfaction': round(employee_satisfaction, 1),
            'Training_Hours': training_hours,
            'Turnover_Rate': round(random.uniform(0.05, 0.25), 3),
            
            # Marketing Metrics
            'Marketing_Spend': round(marketing_spend, 2),
            'Marketing_Cost': round(marketing_spend, 2),  # Alternative name
            'Leads_Generated': leads_generated,
            'Conversion_Rate': round(conversion_rate, 3),
            'Customer_Acquisition_Cost': round(marketing_spend / max(leads_generated * conversion_rate, 1), 2),
            
            # Operational Metrics
            'Production_Volume': production_volume,
            'Volume': production_volume,  # Alternative name
            'Quality_Score': round(quality_score, 1),
            'Delivery_Time_Days': delivery_time,
            'On_Time_Delivery': random.choice([0, 1]),  # Binary
            
            # Banking/Financial Metrics
            'Loan_Amount': round(loan_amount, 2),
            'Interest_Rate': round(interest_rate, 4),
            'Debt_to_Equity_Ratio': round(debt_to_equity, 2),
            'ROE': round((net_income / equity) * 100, 2) if equity > 0 else 0,
            'ROA': round((net_income / total_assets) * 100, 2) if total_assets > 0 else 0,
            
            # Project Management
            'Project_Budget': round(project_budget, 2),
            'Project_Completion_Percent': round(project_completion * 100, 1),
            'Budget_Utilization': round(random.uniform(0.7, 1.2), 2),
            
            # Additional Business KPIs
            'Market_Share': round(random.uniform(0.05, 0.3), 3),
            'Customer_Lifetime_Value': round(random.uniform(1000, 50000), 2),
            'Churn_Rate': round(random.uniform(0.02, 0.15), 3),
            'Upsell_Revenue': round(random.uniform(0, sales_amount * 0.3), 2),
            
            # Status and Category Fields
            'Order_Status': random.choice(['Completed', 'Pending', 'Cancelled', 'Processing']),
            'Payment_Status': random.choice(['Paid', 'Pending', 'Overdue']),
            'Priority': random.choice(['High', 'Medium', 'Low']),
            'Risk_Level': random.choice(['Low', 'Medium', 'High']),
            
            # Time-based metrics
            'Quarter': f"Q{((transaction_date.month - 1) // 3) + 1}",
            'Month': transaction_date.strftime('%B'),
            'Year': transaction_date.year,
            'Week_Number': transaction_date.isocalendar()[1],
            'Day_of_Week': transaction_date.strftime('%A'),
            
            # Additional Financial Ratios Components
            'Accounts_Receivable': round(random.uniform(10000, 100000), 2),
            'Accounts_Payable': round(random.uniform(5000, 80000), 2),
            'Working_Capital': round(current_assets - current_liabilities, 2),
            'Quick_Assets': round(current_assets - inventory, 2),
            
            # Performance Indicators
            'Performance_Score': round(random.uniform(70, 100), 1),
            'Efficiency_Rating': round(random.uniform(0.6, 1.0), 2),
            'Growth_Rate': round(random.uniform(-0.1, 0.3), 3),
            'Benchmark_Score': round(random.uniform(80, 120), 1),
        }
        
        data.append(row)
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Add some calculated fields for advanced analysis
    df['Profit_Margin'] = (df['Net_Income'] / df['Revenue'] * 100).round(2)
    df['Current_Ratio'] = (df['Current_Assets'] / df['Current_Liabilities']).round(2)
    df['Quick_Ratio'] = (df['Quick_Assets'] / df['Current_Liabilities']).round(2)
    df['Inventory_Turnover'] = (df['COGS'] / df['Inventory']).round(2)
    df['Asset_Turnover'] = (df['Revenue'] / df['Total_Assets']).round(2)
    
    return df

def main():
    """Create and save comprehensive test data"""
    print("Creating comprehensive test data for RFP Bayanati Business Intelligence Analysis...")
    
    # Generate the test data
    df = create_comprehensive_test_data()
    
    # Save to Excel file
    excel_filename = 'comprehensive_bi_test_data.xlsx'
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        # Main dataset
        df.to_excel(writer, sheet_name='Business_Data', index=False)
        
        # Create summary sheet with data dictionary
        summary_data = {
            'Sheet_Name': ['Business_Data'],
            'Description': ['Comprehensive business intelligence test data with financial ratios, KPIs, and business metrics'],
            'Rows': [len(df)],
            'Columns': [len(df.columns)],
            'Date_Range': [f"{df['Date'].min().date()} to {df['Date'].max().date()}"],
            'Key_Features': ['Financial ratios, Customer analytics, Employee metrics, Marketing KPIs, Operational data, Project management, Banking ratios']
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Data_Summary', index=False)
        
        # Create column dictionary
        columns_info = []
        for col in df.columns:
            if 'Amount' in col or 'Revenue' in col or 'Cost' in col or 'Assets' in col or 'Income' in col:
                category = 'Financial'
            elif 'Customer' in col or 'Satisfaction' in col or 'Retention' in col:
                category = 'Customer Analytics'
            elif 'Employee' in col or 'Training' in col or 'Turnover' in col:
                category = 'HR Metrics'
            elif 'Marketing' in col or 'Leads' in col or 'Conversion' in col:
                category = 'Marketing KPIs'
            elif 'Production' in col or 'Quality' in col or 'Delivery' in col:
                category = 'Operational'
            elif 'Project' in col or 'Budget' in col:
                category = 'Project Management'
            elif 'Ratio' in col or 'ROE' in col or 'ROA' in col:
                category = 'Financial Ratios'
            else:
                category = 'General'
                
            columns_info.append({
                'Column_Name': col,
                'Category': category,
                'Data_Type': str(df[col].dtype),
                'Sample_Value': str(df[col].iloc[0]) if len(df) > 0 else 'N/A'
            })
        
        columns_df = pd.DataFrame(columns_info)
        columns_df.to_excel(writer, sheet_name='Column_Dictionary', index=False)
    
    print(f"\n✅ Test data created successfully!")
    print(f"📁 File: {excel_filename}")
    print(f"📊 Rows: {len(df):,}")
    print(f"📈 Columns: {len(df.columns)}")
    print(f"� Date Range: {df['Date'].min().date()} to {df['Date'].max().date()}")
    
    print(f"\n🎯 Key Features for BI Analysis:")
    print(f"• Financial Ratios: ROE, ROA, Current Ratio, Quick Ratio, Profit Margins")
    print(f"• Customer Analytics: Satisfaction, Retention, Segmentation, Lifetime Value")
    print(f"• Sales Metrics: Revenue, Growth, Conversion Rates, Upselling")
    print(f"• Employee KPIs: Satisfaction, Training, Turnover, Performance")
    print(f"• Marketing Data: Spend, Leads, CAC, Market Share")
    print(f"• Operational Metrics: Quality, Delivery, Efficiency, Production")
    print(f"• Banking Ratios: Debt-to-Equity, Interest Rates, Loan Data")
    print(f"• Project Management: Budgets, Completion, Utilization")
    
    print(f"\n🚀 Ready to test all RFP Bayanati Business Intelligence features!")
    print(f"📋 Upload this file to your application at http://localhost:8505")

if __name__ == "__main__":
    main()