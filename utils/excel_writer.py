"""
Excel Writer Utility for N-gram Automation
Generates multi-sheet Excel output with N-gram analysis per campaign.
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from typing import Dict, List
import os
from datetime import datetime


# Column configurations for output (reduced for horizontal layout)
NGRAM_COLUMNS = [
    ('ngram', 'N-gram', 20),
    ('impressions', 'Impression', 12),
    ('clicks', 'Click', 10),
    ('spend', 'Spend', 12),
    ('orders', 'Order 14d', 12),
    ('sales', 'Sales 14d', 12),
    ('ctr', 'CTR', 10),
    ('cvr', 'CVR', 10),
    ('cpc', 'CPC', 10),
    ('acos', 'ACOS', 10),
    ('suggestion', 'NE/NP', 8)
]

SEARCH_TERM_COLUMNS = [
    ('search_term', 'Search Term', 35),
    ('impressions', 'Impression', 12),
    ('clicks', 'Click', 10),
    ('spend', 'Spend', 12),
    ('orders', 'Order 14d', 12),
    ('sales', 'Sales 14d', 12),
    ('suggestion', 'NE/NP', 8)
]

# Colors
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
MONO_HEADER_FILL = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
BI_HEADER_FILL = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
TRI_HEADER_FILL = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
SEARCH_HEADER_FILL = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
NE_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
NP_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# Column width for each N-gram section (number of columns per section + 1 gap)
SECTION_WIDTH = len(NGRAM_COLUMNS) + 1  # 13 columns per section (12 + 1 gap)


def sanitize_sheet_name(name: str, max_length: int = 31) -> str:
    """
    Sanitize sheet name for Excel (max 31 chars, no special chars).
    
    Args:
        name: Original sheet name
        max_length: Maximum allowed length
        
    Returns:
        Sanitized sheet name
    """
    # Remove invalid characters
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, ' ')
    
    # Truncate if too long
    if len(name) > max_length:
        name = name[:max_length-3] + '...'
    
    return name.strip()


def format_percentage(value) -> str:
    """Format a value as percentage."""
    if pd.isna(value) or value is None:
        return ""
    return f"{value:.2f}%"


def format_currency(value) -> str:
    """Format a value as currency."""
    if pd.isna(value) or value is None:
        return "$0.00"
    return f"${value:.2f}"


def write_ngram_section_horizontal(ws, df: pd.DataFrame, start_row: int, start_col: int,
                                    section_name: str, header_fill: PatternFill, 
                                    columns: list, ne_col: str, np_col: str) -> int:
    """
    Write an N-gram section to a worksheet at a specific column position (horizontal layout).
    
    Args:
        ws: Worksheet object
        df: DataFrame with N-gram data
        start_row: Starting row number
        start_col: Starting column number
        section_name: Name of the section (e.g., "Monogram")
        header_fill: Fill color for headers
        columns: Column configuration list
        ne_col: Column letter for NE reference list
        np_col: Column letter for NP reference list
        
    Returns:
        Number of rows written (for tracking max height)
    """
    current_row = start_row
    
    # Write section header
    ws.cell(row=current_row, column=start_col, value=section_name)
    ws.cell(row=current_row, column=start_col).font = Font(bold=True, size=12)
    current_row += 2
    
    # Write column headers
    for col_idx, (col_key, col_name, col_width) in enumerate(columns):
        actual_col = start_col + col_idx
        cell = ws.cell(row=current_row, column=actual_col, value=col_name)
        cell.fill = header_fill
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center')
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(actual_col)].width = col_width
    
    header_row = current_row
    current_row += 1
    
    # Find the index of the 'suggestion' column for formula
    suggestion_col_idx = None
    ngram_col_idx = None
    for idx, (col_key, col_name, col_width) in enumerate(columns):
        if col_key == 'suggestion':
            suggestion_col_idx = idx
        if col_key == 'ngram' or col_key == 'search_term':
            ngram_col_idx = idx
    
    # Write data rows
    rows_written = 0
    if not df.empty:
        for row_idx, row in df.iterrows():
            for col_idx, (col_key, col_name, col_width) in enumerate(columns):
                actual_col = start_col + col_idx
                value = row.get(col_key, '')
                
                # Format specific columns
                if col_key in ['ctr', 'cvr', 'acos']:
                    value = format_percentage(value) if pd.notna(value) else ''
                elif col_key in ['spend', 'sales', 'cpc']:
                    value = format_currency(value) if pd.notna(value) else ''
                elif col_key in ['impressions', 'clicks', 'orders']:
                    value = int(value) if pd.notna(value) else 0
                elif col_key == 'suggestion':
                    # Use Excel formula for NE/NP based on reference columns
                    ngram_cell = get_column_letter(start_col + ngram_col_idx) + str(current_row)
                    # Formula: Check NE list first, then NP list
                    # =IF(ISNUMBER(MATCH(A5,$NE$:$NE$,0)),"NE",IF(ISNUMBER(MATCH(A5,$NP$:$NP$,0)),"NP",""))
                    formula = f'=IF(ISNUMBER(MATCH({ngram_cell},{ne_col}:{ne_col},0)),"NE",IF(ISNUMBER(MATCH({ngram_cell},{np_col}:{np_col},0)),"NP",""))'
                    cell = ws.cell(row=current_row, column=actual_col, value=formula)
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal='center')
                    continue
                
                cell = ws.cell(row=current_row, column=actual_col, value=value)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center' if col_key != 'ngram' and col_key != 'search_term' else 'left')
            
            current_row += 1
            rows_written += 1
    
    return rows_written + 3  # Include header rows


def write_search_term_section(ws, df: pd.DataFrame, start_row: int, start_col: int,
                               header_fill: PatternFill, columns: list, 
                               ne_col: str, np_col: str) -> int:
    """
    Write search terms section below the N-gram sections.
    
    Args:
        ws: Worksheet object
        df: DataFrame with search term data
        start_row: Starting row number
        start_col: Starting column number
        header_fill: Fill color for headers
        columns: Column configuration list
        ne_col: Column letter for NE reference list
        np_col: Column letter for NP reference list
        
    Returns:
        Next available row
    """
    current_row = start_row
    
    # Write section header
    ws.cell(row=current_row, column=start_col, value="Search Term")
    ws.cell(row=current_row, column=start_col).font = Font(bold=True, size=12)
    current_row += 2
    
    # Write column headers
    for col_idx, (col_key, col_name, col_width) in enumerate(columns):
        actual_col = start_col + col_idx
        cell = ws.cell(row=current_row, column=actual_col, value=col_name)
        cell.fill = header_fill
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center')
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(actual_col)].width = col_width
    
    current_row += 1
    
    # Find column indices
    suggestion_col_idx = None
    search_term_col_idx = None
    for idx, (col_key, col_name, col_width) in enumerate(columns):
        if col_key == 'suggestion':
            suggestion_col_idx = idx
        if col_key == 'search_term':
            search_term_col_idx = idx
    
    # Write data rows
    if not df.empty:
        for row_idx, row in df.iterrows():
            for col_idx, (col_key, col_name, col_width) in enumerate(columns):
                actual_col = start_col + col_idx
                value = row.get(col_key, '')
                
                # Format specific columns
                if col_key in ['ctr', 'cvr', 'acos']:
                    value = format_percentage(value) if pd.notna(value) else ''
                elif col_key in ['spend', 'sales', 'cpc']:
                    value = format_currency(value) if pd.notna(value) else ''
                elif col_key in ['impressions', 'clicks', 'orders']:
                    value = int(value) if pd.notna(value) else 0
                elif col_key == 'suggestion':
                    # Use Excel formula for NE/NP
                    search_cell = get_column_letter(start_col + search_term_col_idx) + str(current_row)
                    formula = f'=IF(ISNUMBER(MATCH({search_cell},{ne_col}:{ne_col},0)),"NE",IF(ISNUMBER(MATCH({search_cell},{np_col}:{np_col},0)),"NP",""))'
                    cell = ws.cell(row=current_row, column=actual_col, value=formula)
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal='center')
                    continue
                
                cell = ws.cell(row=current_row, column=actual_col, value=value)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center' if col_key != 'search_term' else 'left')
            
            current_row += 1
    
    return current_row


def create_campaign_sheet(wb: Workbook, campaign_name: str, ngrams: Dict[str, pd.DataFrame]) -> None:
    """
    Create a worksheet for a single campaign with N-gram sections side by side (horizontal layout).
    
    Args:
        wb: Workbook object
        campaign_name: Name of the campaign
        ngrams: Dictionary with monograms, bigrams, trigrams, search_terms DataFrames
    """
    sheet_name = sanitize_sheet_name(campaign_name)
    ws = wb.create_sheet(title=sheet_name)
    
    current_row = 1
    
    # Write campaign header
    ws.cell(row=current_row, column=1, value=f"Campaign: {campaign_name}")
    ws.cell(row=current_row, column=1).font = Font(bold=True, size=14)
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=36)
    current_row += 1
    
    # Write back to summary link with hyperlink
    cell = ws.cell(row=current_row, column=1, value="← Back to Summary")
    cell.font = Font(color="0563C1", underline="single")
    cell.hyperlink = "#'Summary'!A1"
    current_row += 2
    
    # Define column positions for horizontal layout
    # Monogram starts at column 1
    # Bigram starts at column 14 (after 12 columns + 1 gap)
    # Trigram starts at column 27 (after another 12 columns + 1 gap)
    # Search Term starts at column 40
    
    mono_start_col = 1
    bi_start_col = mono_start_col + SECTION_WIDTH  # 14
    tri_start_col = bi_start_col + SECTION_WIDTH   # 27
    search_start_col = tri_start_col + SECTION_WIDTH  # 40
    
    # NE and NP reference columns (far right)
    ne_col_num = search_start_col + len(SEARCH_TERM_COLUMNS) + 2  # Column for NE list
    np_col_num = ne_col_num + 1  # Column for NP list
    ne_col = get_column_letter(ne_col_num)
    np_col = get_column_letter(np_col_num)
    
    ngram_start_row = current_row
    
    # Write all three N-gram sections side by side
    mono_df = ngrams.get('monograms', pd.DataFrame())
    bi_df = ngrams.get('bigrams', pd.DataFrame())
    tri_df = ngrams.get('trigrams', pd.DataFrame())
    st_df = ngrams.get('search_terms', pd.DataFrame())
    
    # Write Monogram section
    mono_rows = write_ngram_section_horizontal(ws, mono_df, ngram_start_row, mono_start_col,
                                                "Monogram", MONO_HEADER_FILL, NGRAM_COLUMNS,
                                                ne_col, np_col)
    
    # Write Bigram section
    bi_rows = write_ngram_section_horizontal(ws, bi_df, ngram_start_row, bi_start_col,
                                              "Bigram", BI_HEADER_FILL, NGRAM_COLUMNS,
                                              ne_col, np_col)
    
    # Write Trigram section
    tri_rows = write_ngram_section_horizontal(ws, tri_df, ngram_start_row, tri_start_col,
                                               "Trigram", TRI_HEADER_FILL, NGRAM_COLUMNS,
                                               ne_col, np_col)
    
    # Write Search Term section (4th column, side by side with others)
    search_rows = 0
    if not st_df.empty:
        search_rows = write_ngram_section_horizontal(ws, st_df, ngram_start_row, search_start_col,
                                                      "Search Term", SEARCH_HEADER_FILL, SEARCH_TERM_COLUMNS,
                                                      ne_col, np_col)
    
    # Find the maximum rows written
    max_rows = max(mono_rows, bi_rows, tri_rows, search_rows) if search_rows > 0 else max(mono_rows, bi_rows, tri_rows)
    
    # Create NE/NP reference lists (far right columns)
    # Write headers for NE and NP columns
    ws.cell(row=ngram_start_row, column=ne_col_num, value="NE Keywords")
    ws.cell(row=ngram_start_row, column=ne_col_num).font = Font(bold=True, color="C00000")
    ws.cell(row=ngram_start_row, column=ne_col_num).fill = NE_FILL
    ws.column_dimensions[ne_col].width = 25
    
    ws.cell(row=ngram_start_row, column=np_col_num, value="NP Keywords")
    ws.cell(row=ngram_start_row, column=np_col_num).font = Font(bold=True, color="9C5700")
    ws.cell(row=ngram_start_row, column=np_col_num).fill = NP_FILL
    ws.column_dimensions[np_col].width = 25
    
    # Add instruction for users
    ws.cell(row=ngram_start_row + 1, column=ne_col_num, value="(Add keywords to negate)")
    ws.cell(row=ngram_start_row + 1, column=ne_col_num).font = Font(italic=True, size=9, color="666666")
    ws.cell(row=ngram_start_row + 1, column=np_col_num, value="(Add keywords to negate)")
    ws.cell(row=ngram_start_row + 1, column=np_col_num).font = Font(italic=True, size=9, color="666666")
    
    # Pre-populate NE/NP lists with auto-suggested keywords based on rules
    ne_row = ngram_start_row + 2
    np_row = ngram_start_row + 2
    
    # Collect auto-suggested keywords from all N-gram types
    for ngram_type in ['monograms', 'bigrams', 'trigrams']:
        df = ngrams.get(ngram_type, pd.DataFrame())
        if not df.empty and 'suggestion' in df.columns:
            for _, row in df.iterrows():
                suggestion = row.get('suggestion', '')
                ngram_value = row.get('ngram', '')
                if suggestion == 'NE' and ngram_value:
                    ws.cell(row=ne_row, column=ne_col_num, value=ngram_value)
                    ne_row += 1
                elif suggestion == 'NP' and ngram_value:
                    ws.cell(row=np_row, column=np_col_num, value=ngram_value)
                    np_row += 1


def create_summary_sheet(wb: Workbook, campaigns: Dict[str, dict]) -> None:
    """
    Create the summary sheet with list of all campaigns and their metrics.
    
    Args:
        wb: Workbook object
        campaigns: Dictionary mapping campaign names to their summary metrics
    """
    ws = wb.active
    ws.title = "Summary"
    
    current_row = 1
    
    # Write title
    ws.cell(row=current_row, column=1, value="N-Gram Analysis Summary")
    ws.cell(row=current_row, column=1).font = Font(bold=True, size=16)
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
    current_row += 1
    
    # Write generation date
    ws.cell(row=current_row, column=1, value=f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    ws.cell(row=current_row, column=1).font = Font(italic=True)
    current_row += 2
    
    # Write summary headers
    summary_columns = [
        ('Campaign', 40),
        ('Search Terms', 15),
        ('Monograms', 12),
        ('Bigrams', 12),
        ('Trigrams', 12),
        ('Total Spend', 15),
        ('Total Sales', 15),
        ('Flagged', 10)
    ]
    
    for col_idx, (col_name, col_width) in enumerate(summary_columns, 1):
        cell = ws.cell(row=current_row, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center')
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = col_width
    
    current_row += 1
    
    # Write campaign rows with hyperlinks to their sheets
    for campaign_name, summary in campaigns.items():
        sheet_name = sanitize_sheet_name(campaign_name)
        
        # Campaign name cell with hyperlink to the campaign sheet
        cell = ws.cell(row=current_row, column=1, value=campaign_name)
        cell.border = THIN_BORDER
        cell.hyperlink = f"#'{sheet_name}'!A1"
        cell.font = Font(color="0563C1", underline="single")
        
        ws.cell(row=current_row, column=2, value=summary.get('search_term_count', 0)).border = THIN_BORDER
        ws.cell(row=current_row, column=3, value=summary.get('monogram_count', 0)).border = THIN_BORDER
        ws.cell(row=current_row, column=4, value=summary.get('bigram_count', 0)).border = THIN_BORDER
        ws.cell(row=current_row, column=5, value=summary.get('trigram_count', 0)).border = THIN_BORDER
        ws.cell(row=current_row, column=6, value=format_currency(summary.get('total_spend', 0))).border = THIN_BORDER
        ws.cell(row=current_row, column=7, value=format_currency(summary.get('total_sales', 0))).border = THIN_BORDER
        ws.cell(row=current_row, column=8, value=summary.get('total_flagged', 0)).border = THIN_BORDER
        
        # Center align numeric columns
        for col_idx in range(2, 9):
            ws.cell(row=current_row, column=col_idx).alignment = Alignment(horizontal='center')
        
        current_row += 1
    
    # Add totals row
    current_row += 1
    ws.cell(row=current_row, column=1, value="TOTAL").font = Font(bold=True)


def create_excel_output(processed_data: Dict[str, Dict], output_path: str) -> str:
    """
    Create the final Excel output file with all campaigns.
    
    Args:
        processed_data: Dictionary mapping campaign names to their processed data
            Each campaign has: 'ngrams' (mono/bi/tri/search_terms), 'summary'
        output_path: Path where the Excel file should be saved
        
    Returns:
        Path to the created file
    """
    wb = Workbook()
    
    # Collect campaign summaries for the summary sheet
    campaign_summaries = {}
    
    for campaign_name, data in processed_data.items():
        ngrams = data.get('ngrams', {})
        summary = data.get('summary', {})
        
        # Create campaign sheet
        create_campaign_sheet(wb, campaign_name, ngrams)
        
        # Store summary for summary sheet
        campaign_summaries[campaign_name] = summary
    
    # Create summary sheet (move it to first position)
    create_summary_sheet(wb, campaign_summaries)
    
    # Save the workbook
    wb.save(output_path)
    
    return output_path


def generate_output_filename(prefix: str = "NGram_Analysis") -> str:
    """
    Generate a filename with timestamp.
    
    Args:
        prefix: Prefix for the filename
        
    Returns:
        Filename string
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return f"{prefix}_{timestamp}.xlsx"
