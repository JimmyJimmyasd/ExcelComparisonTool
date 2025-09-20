# Sample Excel files for testing the Excel Comparison Tool

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

def create_sample_data():
    """Create sample Excel files for testing"""
    
    # Sample data for Sheet A (Customers)
    customers_a = {
        'Customer_ID': ['CUST001', 'CUST002', 'CUST003', 'CUST004', 'CUST005', 'CUST006', 'CUST007', 'CUST008'],
        'Customer_Name': ['John Smith', 'Jane Doe', 'Bob Johnson', 'Alice Brown', 'Charlie Wilson', 'Diana Ross', 'Eva Green', 'Frank Miller'],
        'Email': ['john.smith@email.com', 'jane.doe@email.com', 'bob.johnson@email.com', 'alice.brown@email.com', 
                 'charlie.wilson@email.com', 'diana.ross@email.com', 'eva.green@email.com', 'frank.miller@email.com'],
        'Phone': ['555-0101', '555-0102', '555-0103', '555-0104', '555-0105', '555-0106', '555-0107', '555-0108'],
        'City': ['New York', 'Los Angeles', 'Chicago', 'Houston', 'Phoenix', 'Philadelphia', 'San Antonio', 'San Diego']
    }
    
    # Sample data for Sheet B (Orders) - some names slightly different for fuzzy matching
    orders_b = {
        'Customer_Name': ['John Smith', 'Jane Doe', 'Robert Johnson', 'Alice Brown', 'Charles Wilson', 'Diana Ross', 'Eva Green', 'Franklin Miller', 'New Customer'],
        'Order_ID': ['ORD001', 'ORD002', 'ORD003', 'ORD004', 'ORD005', 'ORD006', 'ORD007', 'ORD008', 'ORD009'],
        'Order_Amount': [1250.50, 875.25, 2100.00, 450.75, 1800.00, 925.50, 675.25, 1350.00, 500.00],
        'Order_Date': ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19', '2024-01-20', '2024-01-21', '2024-01-22', '2024-01-23'],
        'Product_Category': ['Electronics', 'Clothing', 'Books', 'Home & Garden', 'Sports', 'Electronics', 'Clothing', 'Books', 'Electronics'],
        'Status': ['Completed', 'Pending', 'Completed', 'Shipped', 'Completed', 'Pending', 'Shipped', 'Completed', 'Pending']
    }
    
    # Create DataFrames
    df_customers = pd.DataFrame(customers_a)
    df_orders = pd.DataFrame(orders_b)
    
    # Save to Excel files
    with pd.ExcelWriter('sample_customers.xlsx', engine='xlsxwriter') as writer:
        df_customers.to_excel(writer, sheet_name='Customers', index=False)
        
        # Add a second sheet for testing
        df_customers_copy = df_customers.copy()
        df_customers_copy['Region'] = ['North', 'West', 'Central', 'South', 'West', 'East', 'South', 'West']
        df_customers_copy.to_excel(writer, sheet_name='Customers_With_Region', index=False)
    
    with pd.ExcelWriter('sample_orders.xlsx', engine='xlsxwriter') as writer:
        df_orders.to_excel(writer, sheet_name='Orders', index=False)
        
        # Add summary sheet
        summary_data = {
            'Total_Orders': [len(df_orders)],
            'Total_Amount': [df_orders['Order_Amount'].sum()],
            'Avg_Order_Amount': [df_orders['Order_Amount'].mean()],
            'Date_Generated': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
        }
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='Summary', index=False)
    
    print("Sample Excel files created:")
    print("- sample_customers.xlsx (with 'Customers' and 'Customers_With_Region' sheets)")
    print("- sample_orders.xlsx (with 'Orders' and 'Summary' sheets)")
    print("\nTest scenarios:")
    print("1. Exact matches: John Smith, Jane Doe, Alice Brown, Diana Ross, Eva Green")
    print("2. Fuzzy matches: Robert Johnson -> Bob Johnson, Charles Wilson -> Charlie Wilson, Franklin Miller -> Frank Miller")
    print("3. Unmatched: New Customer (only in orders)")
    print("4. Missing: Charlie Wilson (customer exists but no order)")

if __name__ == "__main__":
    create_sample_data()