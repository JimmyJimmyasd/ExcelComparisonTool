import streamlit as st
import pandas as pd
from typing import Optional, List

def validate_excel_file(uploaded_file) -> bool:
    """Validate if uploaded file is a valid Excel file"""
    if uploaded_file is None:
        return False
    
    try:
        # Check file extension
        if not uploaded_file.name.lower().endswith(('.xlsx', '.xls')):
            st.error("Please upload a valid Excel file (.xlsx or .xls)")
            return False
        
        # Try to read the file
        pd.ExcelFile(uploaded_file)
        return True
    except Exception as e:
        st.error(f"Invalid Excel file: {str(e)}")
        return False

def validate_dataframe(df: pd.DataFrame, file_name: str) -> bool:
    """Validate DataFrame after loading"""
    if df is None:
        st.error(f"Failed to load data from {file_name}")
        return False
    
    if df.empty:
        st.warning(f"{file_name} is empty")
        return False
    
    if df.shape[0] == 0:
        st.warning(f"No data rows found in {file_name}")
        return False
    
    return True

def validate_columns(df: pd.DataFrame, column_name: str, file_name: str) -> bool:
    """Validate if column exists and has data"""
    if column_name not in df.columns:
        st.error(f"Column '{column_name}' not found in {file_name}")
        return False
    
    # Check if column has any non-null values
    if df[column_name].isnull().all():
        st.warning(f"Column '{column_name}' in {file_name} contains only null values")
        return False
    
    return True

def clean_column_values(series: pd.Series, ignore_case: bool = True) -> pd.Series:
    """Clean and normalize column values for better matching"""
    # Convert to string and handle nulls
    cleaned = series.astype(str).fillna('')
    
    if ignore_case:
        # Convert to lowercase and strip whitespace
        cleaned = cleaned.str.lower().str.strip()
        # Remove extra spaces
        cleaned = cleaned.str.replace(r'\s+', ' ', regex=True)
    
    return cleaned

def get_memory_usage(df: pd.DataFrame) -> str:
    """Get memory usage of DataFrame in human readable format"""
    memory_usage = df.memory_usage(deep=True).sum()
    
    # Convert bytes to MB
    memory_mb = memory_usage / (1024 * 1024)
    
    if memory_mb < 1:
        return f"{memory_usage / 1024:.1f} KB"
    else:
        return f"{memory_mb:.1f} MB"

def safe_excel_read(uploaded_file, sheet_name: str) -> Optional[pd.DataFrame]:
    """Safely read Excel file with error handling"""
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        
        # Basic cleaning
        if df.empty:
            return None
        
        # Remove completely empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        return df
    except Exception as e:
        st.error(f"Error reading sheet '{sheet_name}': {str(e)}")
        return None

def format_similarity_score(score: float) -> str:
    """Format similarity score for display"""
    if score >= 95:
        return f"🟢 {score:.1f}%"
    elif score >= 80:
        return f"🟡 {score:.1f}%"
    else:
        return f"🔴 {score:.1f}%"

def get_column_info(df: pd.DataFrame) -> dict:
    """Get detailed information about DataFrame columns"""
    info = {}
    for col in df.columns:
        info[col] = {
            'dtype': str(df[col].dtype),
            'null_count': df[col].isnull().sum(),
            'unique_count': df[col].nunique(),
            'sample_values': df[col].dropna().head(3).tolist()
        }
    return info