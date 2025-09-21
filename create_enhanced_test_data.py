import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

# Set random seed for reproducible data
np.random.seed(42)
random.seed(42)

def create_enhanced_test_excel():
    """Create an enhanced test Excel file with both diverse sheets and sheets with common columns"""
    
    # Sheet 1: Sales Data Q1
    print("Creating Sales Data Q1 sheet...")
    dates_q1 = pd.date_range(start='2024-01-01', end='2024-03-31', freq='D')
    sales_q1_data = {
        'Date': np.random.choice(dates_q1, 300),
        'Product': np.random.choice(['Laptop', 'Desktop', 'Tablet', 'Phone', 'Accessories'], 300),
        'Region': np.random.choice(['North', 'South', 'East', 'West', 'Central'], 300),
        'Sales_Amount': np.random.normal(1000, 300, 300).round(2),
        'Quantity': np.random.poisson(5, 300),
        'Customer_Rating': np.random.choice([1, 2, 3, 4, 5], 300, p=[0.05, 0.1, 0.2, 0.35, 0.3]),
        'Discount_Percent': np.random.exponential(5, 300).round(1),
        'Sales_Rep': np.random.choice(['Alice', 'Bob', 'Charlie', 'Diana', 'Eve'], 300)
    }
    
    # Add some missing values intentionally
    sales_q1_df = pd.DataFrame(sales_q1_data)
    sales_q1_df.loc[np.random.choice(300, 15), 'Customer_Rating'] = np.nan
    sales_q1_df.loc[np.random.choice(300, 10), 'Discount_Percent'] = np.nan
    
    # Sheet 2: Sales Data Q2 (Similar structure to Q1 for comparison)
    print("Creating Sales Data Q2 sheet...")
    dates_q2 = pd.date_range(start='2024-04-01', end='2024-06-30', freq='D')
    sales_q2_data = {
        'Date': np.random.choice(dates_q2, 350),
        'Product': np.random.choice(['Laptop', 'Desktop', 'Tablet', 'Phone', 'Accessories', 'Monitor'], 350),
        'Region': np.random.choice(['North', 'South', 'East', 'West', 'Central'], 350),
        'Sales_Amount': np.random.normal(1200, 350, 350).round(2),  # Slightly higher average
        'Quantity': np.random.poisson(6, 350),  # Slightly higher average
        'Customer_Rating': np.random.choice([1, 2, 3, 4, 5], 350, p=[0.03, 0.07, 0.15, 0.4, 0.35]),  # Better ratings
        'Discount_Percent': np.random.exponential(4, 350).round(1),  # Lower discounts
        'Sales_Rep': np.random.choice(['Alice', 'Bob', 'Charlie', 'Diana', 'Eve', 'Frank'], 350)
    }
    
    # Add some missing values
    sales_q2_df = pd.DataFrame(sales_q2_data)
    sales_q2_df.loc[np.random.choice(350, 20), 'Customer_Rating'] = np.nan
    sales_q2_df.loc[np.random.choice(350, 8), 'Discount_Percent'] = np.nan
    
    # Sheet 3: Product Performance A
    print("Creating Product Performance A sheet...")
    products_a = ['Product_A', 'Product_B', 'Product_C', 'Product_D', 'Product_E']
    perf_a_data = {
        'Product_Name': np.random.choice(products_a, 200),
        'Revenue': np.random.normal(50000, 15000, 200).round(2),
        'Units_Sold': np.random.poisson(100, 200),
        'Customer_Score': np.random.normal(4.2, 0.8, 200).round(1),
        'Market_Share': np.random.uniform(5, 25, 200).round(2),
        'Launch_Year': np.random.choice([2020, 2021, 2022, 2023, 2024], 200),
        'Category': np.random.choice(['Electronics', 'Software', 'Hardware'], 200)
    }
    
    perf_a_df = pd.DataFrame(perf_a_data)
    perf_a_df.loc[np.random.choice(200, 12), 'Customer_Score'] = np.nan
    
    # Sheet 4: Product Performance B (Similar structure for comparison)
    print("Creating Product Performance B sheet...")
    products_b = ['Product_F', 'Product_G', 'Product_H', 'Product_I', 'Product_J']
    perf_b_data = {
        'Product_Name': np.random.choice(products_b, 180),
        'Revenue': np.random.normal(45000, 12000, 180).round(2),  # Slightly lower
        'Units_Sold': np.random.poisson(85, 180),  # Lower volume
        'Customer_Score': np.random.normal(3.8, 0.9, 180).round(1),  # Lower satisfaction
        'Market_Share': np.random.uniform(3, 20, 180).round(2),  # Lower market share
        'Launch_Year': np.random.choice([2019, 2020, 2021, 2022, 2023], 180),
        'Category': np.random.choice(['Electronics', 'Software', 'Hardware'], 180)
    }
    
    perf_b_df = pd.DataFrame(perf_b_data)
    perf_b_df.loc[np.random.choice(180, 15), 'Customer_Score'] = np.nan
    
    # Sheet 5: Employee Data (Unique structure)
    print("Creating Employee Data sheet...")
    employee_data = {
        'Employee_ID': range(1, 201),
        'Name': [f'Employee_{i}' for i in range(1, 201)],
        'Department': np.random.choice(['IT', 'Sales', 'Marketing', 'HR', 'Finance'], 200),
        'Position': np.random.choice(['Junior', 'Senior', 'Manager', 'Director'], 200, p=[0.4, 0.35, 0.2, 0.05]),
        'Salary': np.random.normal(65000, 20000, 200).round(0),
        'Years_Experience': np.random.exponential(3, 200).round(1),
        'Performance_Rating': np.random.choice([1, 2, 3, 4, 5], 200, p=[0.05, 0.15, 0.4, 0.3, 0.1]),
        'Remote_Work': np.random.choice(['Yes', 'No', 'Hybrid'], 200, p=[0.3, 0.4, 0.3]),
        'Training_Hours': np.random.poisson(20, 200),
        'Last_Promotion': pd.to_datetime(np.random.choice(
            pd.date_range('2020-01-01', '2024-12-31'), 200
        ))
    }
    
    employee_df = pd.DataFrame(employee_data)
    employee_df.loc[np.random.choice(200, 10), 'Years_Experience'] = np.nan
    employee_df.loc[np.random.choice(200, 8), 'Training_Hours'] = np.nan
    
    # Sheet 6: Survey Data (Unique structure)
    print("Creating Survey Data sheet...")
    survey_data = {
        'Response_ID': range(1, 401),
        'Age_Group': np.random.choice(['18-25', '26-35', '36-45', '46-55', '55+'], 400),
        'Gender': np.random.choice(['Male', 'Female', 'Other'], 400, p=[0.45, 0.45, 0.1]),
        'Education': np.random.choice(['High School', 'Bachelor', 'Master', 'PhD'], 400, p=[0.2, 0.5, 0.25, 0.05]),
        'Income_Range': np.random.choice(['<30k', '30-50k', '50-80k', '80-120k', '>120k'], 400),
        'Satisfaction_Score': np.random.choice([1, 2, 3, 4, 5], 400, p=[0.05, 0.1, 0.25, 0.4, 0.2]),
        'Recommendation_Score': np.random.choice(range(0, 11), 400),
        'Usage_Frequency': np.random.choice(['Daily', 'Weekly', 'Monthly', 'Rarely'], 400),
        'Purchase_Amount': np.random.exponential(200, 400).round(2),
        'Years_Customer': np.random.exponential(2, 400).round(1)
    }
    
    survey_df = pd.DataFrame(survey_data)
    survey_df.loc[np.random.choice(400, 25), 'Purchase_Amount'] = np.nan
    survey_df.loc[np.random.choice(400, 15), 'Years_Customer'] = np.nan
    
    # Create Excel file with all sheets
    print("Writing to Excel file...")
    with pd.ExcelWriter('enhanced_test_data.xlsx', engine='openpyxl') as writer:
        sales_q1_df.to_excel(writer, sheet_name='Sales_Q1', index=False)
        sales_q2_df.to_excel(writer, sheet_name='Sales_Q2', index=False)
        perf_a_df.to_excel(writer, sheet_name='Product_Performance_A', index=False)
        perf_b_df.to_excel(writer, sheet_name='Product_Performance_B', index=False)
        employee_df.to_excel(writer, sheet_name='Employee_Data', index=False)
        survey_df.to_excel(writer, sheet_name='Survey_Data', index=False)
    
    print("✅ Enhanced test Excel file created successfully!")
    print("\n📊 **Sheets with Common Columns for Comparison Testing:**")
    print("• Sales_Q1 vs Sales_Q2: Sales_Amount, Quantity, Customer_Rating, Discount_Percent")
    print("• Product_Performance_A vs Product_Performance_B: Revenue, Units_Sold, Customer_Score, Market_Share")
    print("\n📋 **Unique Structure Sheets:**")
    print("• Employee_Data: Salary, Years_Experience, Performance_Rating, Training_Hours")
    print("• Survey_Data: Satisfaction_Score, Recommendation_Score, Purchase_Amount, Years_Customer")
    
    # Show summary statistics
    print(f"\n📈 **Data Summary:**")
    sheets = {
        'Sales_Q1': sales_q1_df,
        'Sales_Q2': sales_q2_df, 
        'Product_Performance_A': perf_a_df,
        'Product_Performance_B': perf_b_df,
        'Employee_Data': employee_df,
        'Survey_Data': survey_df
    }
    
    for name, df in sheets.items():
        print(f"• {name}: {len(df)} rows, {len(df.columns)} columns")

if __name__ == "__main__":
    create_enhanced_test_excel()