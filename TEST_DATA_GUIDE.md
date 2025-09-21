# Test Data Guide 📊

## Overview
I've created a comprehensive test Excel file (`test_data.xlsx`) with 4 different sheets containing realistic sample data to thoroughly test all the analysis features of your Excel Comparison Tool.

## Test File Structure

### 📊 Sales_Data (500 rows, 9 columns)
**Perfect for testing:** Statistical analysis, correlations, missing data handling
- **Date**: Sales transaction dates throughout 2024
- **Product**: Categories (Laptop, Desktop, Tablet, Phone, Accessories)  
- **Region**: Geographic regions (North, South, East, West, Central)
- **Sales_Amount**: Revenue amounts with normal distribution ($400-$1600 range)
- **Quantity**: Items sold per transaction
- **Customer_Age**: Customer demographics (missing data intentionally added)
- **Customer_Satisfaction**: 1-5 rating scale
- **Discount_Applied**: Boolean discount flag
- **Sales_Rep**: Assigned sales representative

### 👥 Employee_Data (200 rows, 10 columns)
**Perfect for testing:** HR analytics, performance correlations, categorical analysis
- **Employee_ID**: Unique identifiers
- **Name**: Employee names
- **Department**: IT, Sales, Marketing, HR, Finance
- **Position**: Junior, Senior, Manager, Director hierarchy
- **Salary**: Compensation data with realistic distribution
- **Years_Experience**: Professional experience
- **Performance_Rating**: 1-5 performance scores
- **Remote_Work**: Work arrangement (Yes/No/Hybrid)
- **Training_Hours**: Professional development hours
- **Last_Promotion**: Date of last career advancement

### 💰 Financial_Data (300 rows, 11 columns)
**Perfect for testing:** Financial correlations, ratio analysis, business metrics
- **Company_ID**: Unique company identifiers
- **Industry**: Tech, Healthcare, Finance, Retail, Manufacturing
- **Revenue_Million**: Company revenue (highly correlated with other metrics)
- **Profit_Million**: Profit amounts (correlated with revenue)
- **Profit_Margin**: Profitability percentages
- **Employees**: Company size
- **Market_Cap_Million**: Market valuation
- **Debt_Ratio**: Financial leverage ratios
- **R&D_Spending**: Research & development investment
- **Founded_Year**: Company age
- **Public_Company**: Public/private company status

### 📋 Survey_Data (1000 rows, 11 columns)
**Perfect for testing:** Large dataset analysis, categorical distributions, customer insights
- **Response_ID**: Survey response identifiers
- **Age_Group**: Demographic segments (18-25, 26-35, etc.)
- **Gender**: Gender distribution
- **Education**: Education levels (High School to PhD)
- **Income_Range**: Income brackets (<30k to >100k)
- **Satisfaction_Score**: 1-10 satisfaction ratings
- **Recommendation_Score**: 0-10 NPS scores
- **Usage_Frequency**: Product usage patterns
- **Product_Category**: Product preferences
- **Purchase_Amount**: Transaction amounts
- **Years_Customer**: Customer tenure

## Testing Recommendations

### 🔍 Comparison Testing
1. **Two Files Mode**: Create a copy of the file and modify some values to test comparison accuracy
2. **Same File Mode**: Compare different sheets (e.g., Sales_Data vs Employee_Data) to test cross-dataset analysis
3. **Multi-Sheet Batch**: Select multiple sheets for comprehensive batch comparisons

### 📈 Statistical Analysis Testing
1. **Financial_Data** - Best for correlation analysis (Revenue vs Profit vs Market Cap)
2. **Sales_Data** - Great for missing data analysis and distributions  
3. **Survey_Data** - Perfect for large dataset performance testing
4. **Employee_Data** - Excellent for categorical analysis and HR metrics

### 🎯 Feature Testing Scenarios
- **Missing Data Analysis**: All sheets have intentionally missing values
- **Correlation Detection**: Financial_Data has strong correlations between revenue, profit, and market cap
- **Distribution Analysis**: Sales_Amount and Salary have normal distributions
- **Categorical Analysis**: All sheets have rich categorical variables
- **Outlier Detection**: Large value ranges in financial metrics

## Quick Start Guide
1. Open your Excel Comparison Tool
2. Upload `test_data.xlsx` in the file selector
3. Try different comparison modes with various sheet combinations
4. Explore the **📊 Statistical Analysis** tab to see all the new analytical features
5. Test the **🎨 Theme** toggle in the sidebar to switch between Light and Dark modes

## Data Quality Notes  
- **Realistic Distributions**: All numerical data uses appropriate statistical distributions
- **Intentional Missing Values**: 5-15% missing data per relevant column for testing
- **Correlated Variables**: Financial metrics show realistic business relationships
- **Diverse Data Types**: Mix of numerical, categorical, date, and boolean data
- **Scalable Testing**: Different dataset sizes (200-1000 rows) for performance testing

This test data will help you thoroughly validate all features of your Excel Comparison Tool, from basic comparisons to advanced statistical analysis! 🚀