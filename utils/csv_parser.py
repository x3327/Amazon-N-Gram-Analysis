"""
CSV Parser Utility for N-gram Automation
Handles parsing raw search term data, filtering ASINs, and grouping by campaign.
"""

import pandas as pd
import re
from typing import Dict, List, Tuple


# Column mapping for Amazon Search Term Report
COLUMN_MAPPING = {
    'date': ['Date', 'date'],
    'portfolio': ['Portfolio name', 'Portfolio Name', 'portfolio name'],
    'currency': ['Currency', 'currency'],
    'campaign': ['Campaign Name', 'Campaign name', 'campaign name', 'Campaign'],
    'ad_group': ['Ad Group Name', 'Ad Group name', 'ad group name', 'Ad Group'],
    'targeting': ['Targeting', 'targeting'],
    'match_type': ['Match Type', 'Match type', 'match type'],
    'search_term': ['Customer Search Term', 'Search Term', 'search term', 'Search term'],
    'impressions': ['Impressions', 'impressions'],
    'clicks': ['Clicks', 'clicks'],
    'ctr': ['Click-Thru Rate (CTR)', 'CTR', 'ctr'],
    'cpc': ['Cost Per Click (CPC)', 'CPC', 'cpc'],
    'spend': ['Spend', 'spend', 'Cost'],
    'sales': ['7 Day Total Sales', 'Sales', 'sales', '14 Day Total Sales', 'Total Sales'],
    'acos': ['Total Advertising Cost of Sales (ACOS)', 'ACOS', 'acos'],
    'roas': ['Total Return on Advertising Spend (ROAS)', 'ROAS', 'roas'],
    'orders': ['7 Day Total Orders (#)', 'Orders', 'orders', '14 Day Total Orders (#)', 'Total Orders'],
    'units': ['7 Day Total Units (#)', 'Units', 'units', '14 Day Total Units (#)'],
    'conversion_rate': ['7 Day Conversion Rate', 'Conversion Rate', 'conversion rate']
}


def find_column(df: pd.DataFrame, possible_names: List[str]) -> str:
    """Find the actual column name from a list of possible names."""
    for name in possible_names:
        if name in df.columns:
            return name
    return None


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Standardize column names to our internal naming convention."""
    column_renames = {}
    
    for standard_name, possible_names in COLUMN_MAPPING.items():
        actual_name = find_column(df, possible_names)
        if actual_name:
            column_renames[actual_name] = standard_name
    
    df = df.rename(columns=column_renames)
    return df


def parse_csv(file_path: str) -> pd.DataFrame:
    """
    Parse a CSV file and return a standardized DataFrame.
    Handles various CSV formats and encoding issues.
    
    Args:
        file_path: Path to the CSV file
        
    Returns:
        Standardized DataFrame with consistent column names
    """
    df = None
    encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    
    # Try different encodings and parsing options
    for encoding in encodings:
        try:
            # Try with default settings first
            df = pd.read_csv(file_path, encoding=encoding, on_bad_lines='skip')
            break
        except Exception:
            try:
                # Try with python engine (more forgiving)
                df = pd.read_csv(file_path, encoding=encoding, engine='python', on_bad_lines='skip')
                break
            except Exception:
                try:
                    # Try with semicolon separator (European format)
                    df = pd.read_csv(file_path, encoding=encoding, sep=';', on_bad_lines='skip')
                    break
                except Exception:
                    continue
    
    if df is None:
        # Last resort: try reading line by line and skip problematic lines
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
            
            # Find the header line (usually line with 'Campaign' or 'Search Term')
            header_line = 0
            for i, line in enumerate(lines[:20]):  # Check first 20 lines
                if 'Campaign' in line or 'Search Term' in line or 'search_term' in line.lower():
                    header_line = i
                    break
            
            # Read from header line, skipping bad lines
            from io import StringIO
            csv_content = '\n'.join(lines[header_line:])
            df = pd.read_csv(StringIO(csv_content), on_bad_lines='skip')
        except Exception as e:
            raise Exception(f"Unable to parse CSV file. Error: {str(e)}")
    
    # Standardize column names
    df = standardize_columns(df)
    
    # Clean up numeric columns
    numeric_columns = ['impressions', 'clicks', 'spend', 'sales', 'orders', 'units']
    for col in numeric_columns:
        if col in df.columns:
            # Remove currency symbols and commas, convert to numeric
            df[col] = df[col].astype(str).str.replace(r'[$,]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Clean up percentage columns
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
