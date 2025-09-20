import streamlit as st
import pandas as pd
from typing import Optional, List, Dict, Tuple
from rapidfuzz import fuzz
import re

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

def analyze_column_patterns(series: pd.Series) -> Dict[str, float]:
    """Analyze data patterns in a column to help with intelligent matching"""
    patterns = {
        'email_pattern': 0.0,
        'phone_pattern': 0.0,
        'date_pattern': 0.0,
        'number_pattern': 0.0,
        'id_pattern': 0.0,
        'name_pattern': 0.0,
        'address_pattern': 0.0
    }
    
    # Convert to string and get non-null sample
    sample_data = series.dropna().astype(str).head(100)
    total_samples = len(sample_data)
    
    if total_samples == 0:
        return patterns
    
    for value in sample_data:
        value_lower = value.lower().strip()
        
        # Email pattern
        if re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', value):
            patterns['email_pattern'] += 1
        
        # Phone pattern (various formats)
        if re.search(r'(\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', value):
            patterns['phone_pattern'] += 1
        
        # Date pattern
        if re.search(r'\b\d{1,4}[-/]\d{1,2}[-/]\d{1,4}\b', value):
            patterns['date_pattern'] += 1
        
        # Number pattern (pure numbers)
        if re.match(r'^\d+\.?\d*$', value):
            patterns['number_pattern'] += 1
        
        # ID pattern (alphanumeric codes)
        if re.match(r'^[A-Za-z0-9]{4,20}$', value) and any(c.isalpha() for c in value) and any(c.isdigit() for c in value):
            patterns['id_pattern'] += 1
        
        # Name pattern (common name words)
        name_indicators = ['john', 'jane', 'smith', 'johnson', 'brown', 'davis', 'miller', 'wilson', 'moore', 'taylor']
        if any(indicator in value_lower for indicator in name_indicators) or len(value.split()) >= 2:
            patterns['name_pattern'] += 1
        
        # Address pattern
        address_indicators = ['street', 'st', 'avenue', 'ave', 'road', 'rd', 'drive', 'dr', 'lane', 'ln', 'blvd']
        if any(indicator in value_lower for indicator in address_indicators):
            patterns['address_pattern'] += 1
    
    # Convert to percentages
    for pattern in patterns:
        patterns[pattern] = (patterns[pattern] / total_samples) * 100
    
    return patterns

def calculate_column_similarity(col_a: str, col_b: str, data_a: pd.Series, data_b: pd.Series) -> Dict[str, float]:
    """Calculate comprehensive similarity between two columns"""
    
    # 1. Name similarity (column names)
    name_similarity = fuzz.ratio(col_a.lower(), col_b.lower())
    
    # 2. Semantic similarity (check for common column name patterns)
    semantic_score = calculate_semantic_similarity(col_a, col_b)
    
    # 3. Data pattern similarity
    patterns_a = analyze_column_patterns(data_a)
    patterns_b = analyze_column_patterns(data_b)
    
    pattern_similarity = 0.0
    for pattern in patterns_a:
        if patterns_a[pattern] > 10 and patterns_b[pattern] > 10:  # Both have significant pattern
            pattern_similarity += min(patterns_a[pattern], patterns_b[pattern])
    
    # 4. Data type compatibility
    type_similarity = 100.0 if str(data_a.dtype) == str(data_b.dtype) else 0.0
    
    # 5. Value overlap (sample-based)
    value_similarity = calculate_value_overlap(data_a, data_b)
    
    # Weighted final score
    final_score = (
        name_similarity * 0.3 +
        semantic_score * 0.2 +
        pattern_similarity * 0.2 +
        type_similarity * 0.1 +
        value_similarity * 0.2
    )
    
    return {
        'overall_score': final_score,
        'name_similarity': name_similarity,
        'semantic_similarity': semantic_score,
        'pattern_similarity': pattern_similarity,
        'type_similarity': type_similarity,
        'value_similarity': value_similarity
    }

def calculate_semantic_similarity(col_a: str, col_b: str) -> float:
    """Calculate semantic similarity between column names using predefined mappings"""
    
    # Common column name synonyms and variations
    synonyms = {
        'name': ['name', 'full_name', 'customer_name', 'client_name', 'person_name', 'user_name'],
        'email': ['email', 'e_mail', 'email_address', 'mail', 'contact_email'],
        'phone': ['phone', 'telephone', 'mobile', 'contact_number', 'phone_number'],
        'id': ['id', 'identifier', 'customer_id', 'client_id', 'user_id', 'account_id'],
        'address': ['address', 'street_address', 'home_address', 'mailing_address'],
        'date': ['date', 'created_date', 'modified_date', 'birth_date', 'registration_date'],
        'amount': ['amount', 'price', 'cost', 'value', 'total', 'sum'],
        'status': ['status', 'state', 'condition', 'stage']
    }
    
    col_a_clean = re.sub(r'[^a-zA-Z]', '', col_a.lower())
    col_b_clean = re.sub(r'[^a-zA-Z]', '', col_b.lower())
    
    # Check if both columns belong to the same semantic category
    for category, variants in synonyms.items():
        if any(variant in col_a_clean for variant in variants) and any(variant in col_b_clean for variant in variants):
            return 90.0  # High semantic similarity
    
    # Check for partial matches
    for category, variants in synonyms.items():
        a_matches = [variant for variant in variants if variant in col_a_clean]
        b_matches = [variant for variant in variants if variant in col_b_clean]
        if a_matches and b_matches:
            return 70.0  # Moderate semantic similarity
    
    return 0.0

def calculate_value_overlap(data_a: pd.Series, data_b: pd.Series) -> float:
    """Calculate percentage of overlapping values between two columns"""
    
    # Get sample of unique values from both columns
    sample_a = set(data_a.dropna().astype(str).head(100).str.lower().str.strip())
    sample_b = set(data_b.dropna().astype(str).head(100).str.lower().str.strip())
    
    if not sample_a or not sample_b:
        return 0.0
    
    # Calculate overlap
    overlap = len(sample_a.intersection(sample_b))
    total_unique = len(sample_a.union(sample_b))
    
    if total_unique == 0:
        return 0.0
    
    return (overlap / total_unique) * 100

def suggest_column_mappings(df_a: pd.DataFrame, df_b: pd.DataFrame, top_n: int = 3) -> List[Dict]:
    """Generate intelligent column mapping suggestions"""
    
    suggestions = []
    
    for col_a in df_a.columns:
        column_suggestions = []
        
        for col_b in df_b.columns:
            similarity_metrics = calculate_column_similarity(
                col_a, col_b, df_a[col_a], df_b[col_b]
            )
            
            if similarity_metrics['overall_score'] > 20:  # Only suggest if reasonable similarity
                column_suggestions.append({
                    'column_a': col_a,
                    'column_b': col_b,
                    'confidence': similarity_metrics['overall_score'],
                    'reasons': generate_suggestion_reasons(similarity_metrics),
                    'metrics': similarity_metrics
                })
        
        # Sort by confidence and take top suggestions
        column_suggestions.sort(key=lambda x: x['confidence'], reverse=True)
        suggestions.extend(column_suggestions[:top_n])
    
    # Remove duplicates and sort by overall confidence
    seen_pairs = set()
    unique_suggestions = []
    
    for suggestion in sorted(suggestions, key=lambda x: x['confidence'], reverse=True):
        pair = (suggestion['column_a'], suggestion['column_b'])
        if pair not in seen_pairs:
            seen_pairs.add(pair)
            unique_suggestions.append(suggestion)
    
    return unique_suggestions[:10]  # Return top 10 suggestions

def generate_suggestion_reasons(metrics: Dict[str, float]) -> List[str]:
    """Generate human-readable reasons for column mapping suggestions"""
    
    reasons = []
    
    if metrics['name_similarity'] > 70:
        reasons.append(f"Column names are very similar ({metrics['name_similarity']:.0f}% match)")
    elif metrics['name_similarity'] > 40:
        reasons.append(f"Column names are somewhat similar ({metrics['name_similarity']:.0f}% match)")
    
    if metrics['semantic_similarity'] > 80:
        reasons.append("Columns appear to contain the same type of data")
    
    if metrics['pattern_similarity'] > 30:
        reasons.append(f"Data patterns are similar ({metrics['pattern_similarity']:.0f}% match)")
    
    if metrics['value_similarity'] > 20:
        reasons.append(f"Contains overlapping values ({metrics['value_similarity']:.0f}% overlap)")
    
    if metrics['type_similarity'] > 90:
        reasons.append("Compatible data types")
    
    if not reasons:
        reasons.append("Basic compatibility detected")
    
    return reasons