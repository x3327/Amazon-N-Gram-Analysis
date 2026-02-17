"""
CSV Parser Utility for N-gram Automation
Handles parsing raw search term data, filtering ASINs, and grouping by campaign.
"""

import pandas as pd
import re
from typing import Dict, List, Tuple


# Column mapping for Amazon Search Term Report
COLUMN_MAPPING = {
    'date': ['Date', 'date', 'Start Date', 'End Date'],
    'portfolio': ['Portfolio name', 'Portfolio Name', 'portfolio name', 'Portfolio'],
    'currency': ['Currency', 'currency'],
    'campaign': ['Campaign Name', 'Campaign name', 'campaign name', 'Campaign', 'CampaignName', 'campaign_name'],
    'ad_group': ['Ad Group Name', 'Ad Group name', 'ad group name', 'Ad Group', 'AdGroup', 'ad_group'],
    'targeting': ['Targeting', 'targeting', 'Targeting Expression'],
    'match_type': ['Match Type', 'Match type', 'match type'],
    'search_term': ['Customer Search Term', 'Search Term', 'search term', 'Search term', 'Keyword', 'keyword', 'Query', 'query', 'SearchTerm', 'search_term'],
    'impressions': ['Impressions', 'impressions', 'Impr.', 'impr'],
    'clicks': ['Clicks', 'clicks'],
    'ctr': ['Click-Thru Rate (CTR)', 'CTR', 'ctr', 'Click-Through Rate'],
    'cpc': ['Cost Per Click (CPC)', 'CPC', 'cpc'],
    'spend': ['Spend', 'spend', 'Cost', 'cost', 'Total Spend'],
    'sales': ['7 Day Total Sales', 'Sales', 'sales', '14 Day Total Sales', 'Total Sales', '7 Day Sales', '14 Day Sales', 'Total Sales (14d)'],
    'acos': ['Total Advertising Cost of Sales (ACOS)', 'ACOS', 'acos', 'Acos'],
    'roas': ['Total Return on Advertising Spend (ROAS)', 'ROAS', 'roas'],
    'orders': ['7 Day Total Orders (#)', 'Orders', 'orders', '14 Day Total Orders (#)', 'Total Orders', '7 Day Orders', '14 Day Orders', 'Total Orders (14d)'],
    'units': ['7 Day Total Units (#)', 'Units', 'units', '14 Day Total Units (#)', 'Total Units'],
    'conversion_rate': ['7 Day Conversion Rate', 'Conversion Rate', 'conversion rate']
}


def find_column(df: pd.DataFrame, possible_names: List[str]) -> str:
    """Find the actual column name from a list of possible names (case-insensitive)."""
    # First try exact match
    for name in possible_names:
        if name in df.columns:
            return name
    
    # Try case-insensitive match
    df_columns_lower = {col.lower(): col for col in df.columns}
    for name in possible_names:
        if name.lower() in df_columns_lower:
            return df_columns_lower[name.lower()]
    
    return None


def find_column_fuzzy(df: pd.DataFrame, keywords: List[str]) -> str:
    """Find column by checking if any keyword is in the column name."""
    for col in df.columns:
        col_lower = col.lower()
        for keyword in keywords:
            if keyword.lower() in col_lower:
                return col
    return None


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Standardize column names to our internal naming convention."""
    column_renames = {}
    
    for standard_name, possible_names in COLUMN_MAPPING.items():
        # Try exact match first
        actual_name = find_column(df, possible_names)
        if actual_name:
            column_renames[actual_name] = standard_name
        else:
            # Try fuzzy match as fallback
            fuzzy_keywords = {
                'search_term': ['search', 'term', 'customer'],
                'impressions': ['impression'],
                'clicks': ['click'],
                'spend': ['spend', 'cost'],
                'campaign': ['campaign'],
                'sales': ['sales'],
                'orders': ['order'],
                'acos': ['acos'],
                'ctr': ['ctr', 'click-thru'],
            }
            if standard_name in fuzzy_keywords:
                actual_name = find_column_fuzzy(df, fuzzy_keywords[standard_name])
                if actual_name:
                    column_renames[actual_name] = standard_name
    
    df = df.rename(columns=column_renames)
    return df


def detect_file_type(file_path: str) -> str:
    """Detect if file is CSV or Excel based on extension and content."""
    import os
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext in ['.xlsx', '.xls']:
        return 'excel'
    elif ext == '.csv':
        return 'csv'
    else:
        # Try to detect by reading first few bytes
        with open(file_path, 'rb') as f:
            header = f.read(4)
            if header == b'PK\x03\x04':  # ZIP/XLSX signature
                return 'excel'
        return 'csv'


def parse_csv(file_path: str) -> pd.DataFrame:
    """
    Parse a CSV or Excel file and return a standardized DataFrame.
    Handles various file formats and encoding issues.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Standardized DataFrame with consistent column names
    """
    import sys
    
    df = None
    file_type = detect_file_type(file_path)
    
    # Try reading as Excel first if it looks like Excel
    if file_type == 'excel':
        try:
            df = pd.read_excel(file_path)
            print(f"DEBUG: Successfully read as Excel file", file=sys.stderr)
        except Exception as e:
            print(f"DEBUG: Excel read failed: {e}", file=sys.stderr)
    
    # If Excel failed or it's CSV, try CSV parsing
    if df is None:
        encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1', 'windows-1252']
        
        for encoding in encodings:
            try:
                df = pd.read_csv(file_path, encoding=encoding, on_bad_lines='skip', low_memory=False)
                print(f"DEBUG: Successfully read CSV with {encoding} encoding", file=sys.stderr)
                break
            except Exception:
                try:
                    df = pd.read_csv(file_path, encoding=encoding, engine='python', on_bad_lines='skip')
                    print(f"DEBUG: Successfully read CSV with {encoding} encoding (python engine)", file=sys.stderr)
                    break
                except Exception:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding, sep=';', on_bad_lines='skip')
                        print(f"DEBUG: Successfully read CSV with {encoding} encoding (semicolon)", file=sys.stderr)
                        break
                    except Exception:
                        continue
    
    # Last resort: manual parsing
    if df is None or len(df.columns) == 0 or all(len(str(col)) > 50 for col in df.columns):
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            
            # Find header line
            lines = content.split('\n')
            header_idx = 0
            for i, line in enumerate(lines[:30]):
                if any(keyword in line for keyword in ['Campaign', 'Search Term', 'Customer', 'Impressions', 'Clicks']):
                    header_idx = i
                    break
            
            # Parse from header
            from io import StringIO
            clean_content = '\n'.join(lines[header_idx:])
            df = pd.read_csv(StringIO(clean_content), on_bad_lines='skip', engine='python')
            print(f"DEBUG: Successfully read with manual parsing", file=sys.stderr)
        except Exception as e:
            raise Exception(f"Unable to parse file. Please ensure it's a valid CSV or Excel file. Error: {str(e)}")
    
    # Validate we got meaningful data
    if df is None or len(df.columns) == 0:
        raise Exception("Could not read any data from the file. Please check the file format.")
    
    # Check if columns look garbled (very long column names suggest binary/encrypted data)
    if all(len(str(col)) > 30 for col in df.columns):
        raise Exception("File appears to be corrupted, encrypted, or in an unsupported format. Please upload a valid CSV or Excel file.")
    
    print(f"DEBUG: Original columns: {list(df.columns)}", file=sys.stderr)
    
    # Standardize column names
    df = standardize_columns(df)
    
    print(f"DEBUG: Standardized columns: {list(df.columns)}", file=sys.stderr)
    
    # Clean up numeric columns
    numeric_columns = ['impressions', 'clicks', 'spend', 'sales', 'orders', 'units']
    for col in numeric_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'[$,]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    pct_columns = ['ctr', 'acos', 'conversion_rate']
    for col in pct_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('%', '')
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    return df


def filter_asins(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, int]:
    """
    Filter out rows where search term starts with 'B0' (ASIN pattern).
    
    Args:
        df: DataFrame with search term data
        
    Returns:
        Tuple of (filtered DataFrame without ASINs, DataFrame with only ASINs, count of ASINs)
    """
    if 'search_term' not in df.columns:
        return df, pd.DataFrame(), 0
    
    # Count rows before filtering
    original_count = len(df)
    
    # Filter out search terms starting with B0 (case insensitive)
    # ASIN pattern: starts with B0 followed by alphanumeric characters
    asin_pattern = r'^[Bb]0[A-Za-z0-9]+'
    
    # Mask for non-ASIN rows
    mask = ~df['search_term'].astype(str).str.match(asin_pattern, na=False)
    df_filtered = df[mask].copy()
    
    # Get ASIN rows
    df_asins = df[~mask].copy()
    
    asin_count = len(df_asins)
    
    return df_filtered, df_asins, asin_count


def group_by_campaign(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """
    Group the data by campaign name.
    
    Args:
        df: DataFrame with search term data
        
    Returns:
        Dictionary mapping campaign names to their DataFrames
    """
    if 'campaign' not in df.columns:
        # If no campaign column, return all data under a single key
        return {'All Campaigns': df}
    
    campaigns = {}
    
    for campaign_name, group_df in df.groupby('campaign', dropna=False):
        # Handle NaN campaign names
        if pd.isna(campaign_name):
            campaign_name = 'Unknown Campaign'
        
        campaigns[campaign_name] = group_df.reset_index(drop=True)
    
    return campaigns


def get_data_summary(df: pd.DataFrame) -> dict:
    """
    Get a summary of the data.
    
    Args:
        df: DataFrame with search term data
        
    Returns:
        Dictionary with summary statistics
    """
    summary = {
        'total_rows': len(df),
        'total_campaigns': df['campaign'].nunique() if 'campaign' in df.columns else 0,
        'total_search_terms': df['search_term'].nunique() if 'search_term' in df.columns else 0,
        'total_impressions': df['impressions'].sum() if 'impressions' in df.columns else 0,
        'total_clicks': df['clicks'].sum() if 'clicks' in df.columns else 0,
        'total_spend': df['spend'].sum() if 'spend' in df.columns else 0,
        'total_sales': df['sales'].sum() if 'sales' in df.columns else 0,
        'total_orders': df['orders'].sum() if 'orders' in df.columns else 0,
    }
    
    return summary


def validate_csv(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    """
    Validate that the CSV has the required columns.
    
    Args:
        df: DataFrame to validate
        
    Returns:
        Tuple of (is_valid, list of missing required columns)
    """
    required_columns = ['search_term', 'impressions', 'clicks', 'spend']
    missing = []
    
    for col in required_columns:
        if col not in df.columns:
            missing.append(col)
    
    return len(missing) == 0, missing
