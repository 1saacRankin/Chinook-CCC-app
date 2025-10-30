"""
Cash Conversion Cycle & 3-Statement Financial Model Calculator


Features:
- Automatic detection of time period frequency (monthly/yearly)
- Robust date parsing and normalization
- Multi-sheet Excel support with automatic combination
- Cash Conversion Cycle (CCC) calculation
- Complete financial ratios: AR turnover, AP turnover, Debt-to-Equity, DSC, etc
- Growth rate analysis
- Automated forecasting with user-adjustable parameters
- Working capital analysis
- 3-statement modeling capabilities

Author: Financial Analysis Tool
Version: 2.2
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
from datetime import datetime
from io import BytesIO
from dateutil import parser
import calendar

# ============================================================================
# CONFIGURATION
# ============================================================================

st.set_page_config(
    page_title="Cash Conversion Cycle & 3-Statement Model Calculator",
    layout="wide",
    page_icon="📊"
)

# ============================================================================
# DATE & TIME PERIOD HANDLING
# ============================================================================

def parse_date_flexible(date_str):
    """
    Flexibly parse various date formats including:
    - Jan 25, January 2025, Jan-25, 01/2025, Jan 26, January 2026
    - Sep-21, September 2021, 9/2021
    - 2025, 2026, 2015.0 (years only, including floats)
    - 2025-01-15 (ISO format)
    - Q1 2025, etc.
    
    Returns standardized datetime object or None
    """
    if pd.isna(date_str):
        return None
    
    # Convert to string and strip whitespace
    date_str = str(date_str).strip()
    
    # Handle float years (like 2015.0 from Excel)
    try:
        float_val = float(date_str)
        if 1900 <= float_val <= 2100 and float_val == int(float_val):
            # It's a year
            return datetime(int(float_val), 12, 31)
    except:
        pass
    
    # Try pandas to_datetime first (handles ISO format like 2025-01-15)
    try:
        dt = pd.to_datetime(date_str, errors='coerce')
        if not pd.isna(dt):
            return dt
    except:
        pass
    
    # Handle common patterns manually
    patterns = [
        # ISO format: 2025-01-15, 2025-1-15
        (r'(\d{4})-(\d{1,2})-(\d{1,2})', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))),
        
        # ISO month format: 2025-01, 2025-1
        (r'(\d{4})-(\d{1,2})$', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), 1)),
        
        # Month name/abbrev + 2-4 digit year: Jan 25, Jan-25, Jan 2025, January 2026, jan-26
        (r'\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b[\s\-,]*(\d{2,4})\b', 
         lambda m: parse_month_year(m.group(1), m.group(2))),
        
        # Numeric month + year: 01/2025, 1-2025, 01-2025, 12/2026
        (r'\b(\d{1,2})[\-/](\d{4})\b', 
         lambda m: datetime(int(m.group(2)), int(m.group(1)), 1)),
        
        # Just a 4-digit year: 2015, 2016, 2025, 2026 (end of year)
        (r'^\s*(\d{4})(?:\.0)?\s*$',  # Also matches 2015.0
         lambda m: datetime(int(float(m.group(1))), 12, 31)),
        
        # Just a 2-digit year: 15, 16, 25, 26 (end of year)
        (r'^\s*(\d{2})(?:\.0)?\s*$', 
         lambda m: parse_two_digit_year(m.group(1))),
    ]
    
    for pattern, handler in patterns:
        match = re.search(pattern, date_str.lower())
        if match:
            try:
                return handler(match)
            except:
                continue
    
    return None


def parse_month_year(month_str, year_str):
    """Convert month name/abbrev and year to datetime."""
    # Comprehensive month mapping
    month_map = {
        'jan': 1, 'january': 1,
        'feb': 2, 'february': 2,
        'mar': 3, 'march': 3,
        'apr': 4, 'april': 4,
        'may': 5,
        'jun': 6, 'june': 6,
        'jul': 7, 'july': 7,
        'aug': 8, 'august': 8,
        'sep': 9, 'sept': 9, 'september': 9,
        'oct': 10, 'october': 10,
        'nov': 11, 'november': 11,
        'dec': 12, 'december': 12
    }
    
    month_str_clean = month_str.lower().strip()
    
    # Try full match first
    month_num = month_map.get(month_str_clean)
    
    # If no match, try partial match (for abbreviated versions)
    if month_num is None:
        for key, val in month_map.items():
            if month_str_clean.startswith(key[:3]):
                month_num = val
                break
    
    if month_num is None:
        return None
    
    # Handle 2-digit vs 4-digit years
    year = int(year_str)
    if year < 100:
        # Assume 2000s for years 0-49, 1900s for 50-99
        year += 2000 if year < 50 else 1900
    
    return datetime(year, month_num, 1)


def parse_two_digit_year(year_str):
    """Parse a 2-digit year string to a full datetime (end of year)."""
    year = int(year_str)
    if year < 100:
        year += 2000 if year < 50 else 1900
    return datetime(year, 12, 31)


def normalize_column_dates(columns):
    """
    Parse and normalize column headers to standardized format.
    Returns list of (original_name, parsed_date, display_name) tuples
    """
    normalized = []
    
    for col in columns:
        parsed = parse_date_flexible(col)
        if parsed:
            # Check if it looks like a year-only date
            col_str = str(col).strip()
            # If original is just 4 digits or parsed to Dec 31, show just year
            if (re.match(r'^\d{4}(?:\.0)?$', col_str) or 
                (parsed.month == 12 and parsed.day == 31)):
                display = str(parsed.year)
            # If it's an ISO format with day, show month-year
            elif re.match(r'\d{4}-\d{1,2}-\d{1,2}', col_str):
                display = parsed.strftime('%b %Y')
            else:
                display = parsed.strftime('%b %Y')  # e.g., "Jan 2025"
        else:
            display = str(col)
        
        normalized.append((col, parsed, display))
    
    return normalized


def detect_time_frequency(dates):
    """
    Detect if data is monthly, quarterly, or yearly.
    Returns: ('monthly', 'quarterly', 'yearly', or 'unknown') and average days between periods
    """
    valid_dates = [d for d in dates if d is not None]
    
    if len(valid_dates) < 2:
        return 'unknown', None
    
    # Sort dates
    valid_dates = sorted(valid_dates)
    
    # Calculate differences in days
    diffs = [(valid_dates[i+1] - valid_dates[i]).days for i in range(len(valid_dates)-1)]
    
    # Remove outliers (differences more than 3x median)
    if len(diffs) > 3:
        median_diff = np.median(diffs)
        diffs = [d for d in diffs if d < median_diff * 3]
    
    if not diffs:
        return 'unknown', None
    
    avg_diff = np.mean(diffs)
    median_diff = np.median(diffs)
    
    # Classify based on average difference with tolerance
    # Use median for more robust classification
    if 20 <= median_diff <= 45:  # ~30 days ± tolerance
        return 'monthly', avg_diff
    elif 70 <= median_diff <= 120:  # ~90 days ± tolerance
        return 'quarterly', avg_diff
    elif 300 <= median_diff <= 400:  # ~365 days ± tolerance
        return 'yearly', avg_diff
    else:
        # Check if it's yearly with wide spacing
        if avg_diff > 250:
            return 'yearly', avg_diff
        return 'unknown', avg_diff

# ============================================================================
# EXCEL LOADING & SHEET COMBINATION
# ============================================================================

def load_excel_with_sheets(uploaded_file):
    """
    Load Excel file and return all sheets.
    Returns: dict of {sheet_name: dataframe}
    """
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheets = {}
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            sheets[sheet_name] = df
        
        return sheets
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None


def find_header_row(df):
    """
    Intelligently find the header row in a dataframe.
    Looks for row with date-like patterns or common financial terms.
    """
    for idx in range(min(10, len(df))):
        row = df.iloc[idx]
        
        # Skip if first cell is empty
        if pd.isna(row.iloc[0]):
            continue
        
        # Check if row contains date patterns (skip first column which is labels)
        date_count = 0
        non_empty_count = 0
        
        for cell in row.iloc[1:]:  # Skip first column
            if pd.notna(cell):
                non_empty_count += 1
                if parse_date_flexible(cell) is not None:
                    date_count += 1
        
        # If more than 50% of non-empty cells are dates, this is likely header
        if non_empty_count > 0 and date_count / non_empty_count > 0.5:
            return idx
        
        # Check for common header indicators in first cell
        first_cell = str(row.iloc[0]).lower()
        header_keywords = ['line item', 'account', 'description', 'period', 'metric', 'line_item']
        if any(keyword in first_cell for keyword in header_keywords):
            # Also check if subsequent cells look like dates/years
            year_like = sum(1 for cell in row.iloc[1:] if str(cell).strip().replace('.0', '').isdigit() and len(str(cell).strip().replace('.0', '')) == 4)
            if year_like >= 3:  # At least 3 year-like values
                return idx
    
    return 0  # Default to first row


def process_sheet(df):
    """
    Process a single sheet: find header, set index, normalize columns.
    Returns processed dataframe or None if invalid.
    """
    # Find header row
    header_row = find_header_row(df)
    
    # Set header and remove rows above it
    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    
    # First column becomes index (row labels)
    first_col_name = df.columns[0]
    df = df.set_index(first_col_name)
    df.index.name = 'Line_Item'
    
    # Remove completely empty rows and columns
    df = df.dropna(how='all', axis=0)
    df = df.dropna(how='all', axis=1)
    
    # Get remaining columns (excluding the index)
    data_columns = df.columns
    
    # Normalize column names (dates)
    normalized_cols = normalize_column_dates(data_columns)
    
    # Keep only columns that parsed as dates
    valid_col_indices = [i for i, (orig, parsed, disp) in enumerate(normalized_cols) if parsed is not None]
    
    if len(valid_col_indices) == 0:
        return None
    
    # Filter to valid columns
    df = df.iloc[:, valid_col_indices]
    
    # Rename columns to display names
    new_col_names = [normalized_cols[i][2] for i in valid_col_indices]
    df.columns = new_col_names
    
    # Store parsed dates as metadata
    df.attrs['parsed_dates'] = [normalized_cols[i][1] for i in valid_col_indices]
    
    # Convert all data columns to numeric, coercing errors
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    return df


def combine_sheets(sheets_dict):
    """
    Combine multiple sheets into one dataframe.
    Handles overlapping columns and row labels intelligently.
    """
    if not sheets_dict:
        return None
    
    # Process each sheet
    processed_sheets = {}
    for name, df in sheets_dict.items():
        processed = process_sheet(df)
        if processed is not None:
            processed_sheets[name] = processed
    
    if not processed_sheets:
        return None
    
    # If only one valid sheet, return it
    if len(processed_sheets) == 1:
        return list(processed_sheets.values())[0]
    
    # Combine multiple sheets
    # Strategy: concatenate all unique columns, merge on index
    combined = None
    
    for sheet_name, df in processed_sheets.items():
        if combined is None:
            combined = df.copy()
        else:
            # Merge on index, keeping all rows and columns
            combined = combined.join(df, how='outer', rsuffix=f'_{sheet_name}')
    
    # Sort columns by date if possible
    if hasattr(combined, 'attrs') and 'parsed_dates' in combined.attrs:
        try:
            col_dates = combined.attrs['parsed_dates']
            sorted_indices = sorted(range(len(col_dates)), key=lambda i: col_dates[i] if col_dates[i] else datetime.max)
            combined = combined.iloc[:, sorted_indices]
        except:
            pass
    
    return combined


# ============================================================================
# DATA CLEANING & FUZZY MATCHING
# ============================================================================

def clean_label(label):
    """Clean row labels for matching."""
    if pd.isna(label):
        return ""
    
    label = str(label).lower()
    # Remove special characters, numbers, extra whitespace
    label = re.sub(r'[^a-z\s]', ' ', label)
    label = ' '.join(label.split())
    return label


def fuzzy_match_row(df, search_terms, threshold=0.6):
    """
    Find best matching row in dataframe using fuzzy matching.
    Returns row label or None.
    """
    if df is None or len(df) == 0:
        return None
    
    cleaned_index = {clean_label(idx): idx for idx in df.index}
    
    best_match = None
    best_score = 0
    
    for term in search_terms:
        cleaned_term = clean_label(term)
        
        # Exact match
        if cleaned_term in cleaned_index:
            return cleaned_index[cleaned_term]
        
        # Fuzzy match
        for cleaned_idx, original_idx in cleaned_index.items():
            # Calculate simple similarity score
            if cleaned_term in cleaned_idx:
                score = len(cleaned_term) / len(cleaned_idx)
            elif cleaned_idx in cleaned_term:
                score = len(cleaned_idx) / len(cleaned_term)
            else:
                # Word overlap
                term_words = set(cleaned_term.split())
                idx_words = set(cleaned_idx.split())
                if term_words and idx_words:
                    overlap = len(term_words & idx_words)
                    score = overlap / max(len(term_words), len(idx_words))
                else:
                    score = 0
            
            if score > best_score:
                best_score = score
                best_match = original_idx
    
    return best_match if best_score >= threshold else None


def auto_detect_variables(df):
    """
    Automatically detect common financial statement line items.
    Returns dict mapping variable names to row labels.
    """
    search_map = {
        # Classics
        'Sales': ['sales', 'revenue', 'total revenue', 'net revenue', 'total sales'],
        'COGS': ['cogs', 'cost of goods sold', 'cost of sales', 'cost of revenue'],
        'Gross Profit': ['gross profit', 'gross margin', 'gross income'],
        'Operating Expenses': ['operating expenses', 'opex', 'operating costs'],
        'EBITDA': ['ebitda', 'earnings before interest'],
        'Depreciation': ['depreciation', 'depr', 'depreciation and amortization', 'da', 'd&a'],
        'Interest Expense': ['interest expense', 'interest', 'interest paid'],
        'Net Income': ['net income', 'net profit', 'net earnings', 'profit'],
        
        # Balance Sheet - Assets
        'Cash': ['cash', 'cash and cash equivalents', 'cash and equivalents'],
        'Accounts Receivable': ['accounts receivable', 'receivables', 'ar', 'trade receivables'],
        'Inventory': ['inventory', 'inventories', 'finished goods'],
        'WIP': ['wip', 'work in progress', 'work in process'],
        'Prepaid Expenses': ['prepaid expenses', 'prepaid', 'prepaids'],
        'Current Assets': ['current assets', 'total current assets'],
        'Fixed Assets': ['fixed assets', 'ppe', 'property plant equipment', 'net fixed assets', 'fixed assets net'],
        'Total Assets': ['total assets', 'assets'],
        
        # Balance Sheet - Liabilities
        'Accounts Payable': ['accounts payable', 'payables', 'ap', 'trade payables'],
        'Deferred Income': ['deferred income', 'deferred revenue', 'unearned revenue'],
        'Accrued Expenses': ['accrued expenses', 'accrued liabilities', 'accruals'],
        'Current Liabilities': ['current liabilities', 'total current liabilities'],
        'Long Term Debt': ['long term debt', 'long term liabilities', 'lt debt', 'long term borrowings'],
        'Short Term Debt': ['short term debt', 'short term borrowings', 'st debt', 'current debt'],
        'Total Liabilities': ['total liabilities', 'liabilities'],
        
        # Equity
        'Equity': ['equity', 'shareholders equity', 'stockholders equity', 'total equity'],
        'Retained Earnings': ['retained earnings', 'retained', 'accumulated earnings'],
    }
    
    mapping = {}
    for var_name, search_terms in search_map.items():
        found = fuzzy_match_row(df, search_terms)
        mapping[var_name] = found
    
    return mapping


def get_row_data(df, row_label, default=0.0):
    """Safely extract row data with fallback to default."""
    if row_label and row_label in df.index:
        return pd.to_numeric(df.loc[row_label], errors='coerce').fillna(default)
    return pd.Series(default, index=df.columns, dtype=float)


# ============================================================================
# FINANCIAL CALCULATIONS
# ============================================================================

def calculate_ccc_metrics(sales, cogs, ar, ap, inventory, wip, deferred_income, period_days):
    """
    Calculate Cash Conversion Cycle and components.
    
    CCC = DSO + DIO - DPO
    where:
    - DSO (Days Sales Outstanding) = (AR / Sales) * 365
    - DIO (Days Inventory Outstanding) = ((Inventory + WIP) / COGS) * 365
    - DPO (Days Payable Outstanding) = (AP / COGS) * 365
    
    Deferred income reduces effective AR (cash collected but not yet earned)
    """
    with np.errstate(divide='ignore', invalid='ignore'):
        # Adjust AR for deferred income (cash already collected)
        effective_ar = ar - deferred_income
        effective_ar = effective_ar.clip(lower=0)  # Can't be negative
        
        # Calculate components
        dso = (effective_ar / sales) * 365
        dso = dso.replace([np.inf, -np.inf], np.nan)
        
        total_inventory = inventory + wip
        dio = (total_inventory / cogs) * 365
        dio = dio.replace([np.inf, -np.inf], np.nan)
        
        dpo = (ap / cogs) * 365
        dpo = dpo.replace([np.inf, -np.inf], np.nan)
        
        ccc = dso + dio - dpo
    
    return {
        'DSO': dso,
        'DIO': dio,
        'DPO': dpo,
        'CCC': ccc
    }


def calculate_financial_ratios(df_data):
    """
    Calculate comprehensive financial ratios.
    
    Input: dict with keys like 'Sales', 'AR', 'AP', etc.
    Returns: dict of ratio series
    """
    ratios = {}
    
    # AR Turnover = Sales / AR
    if 'Sales' in df_data and 'AR' in df_data:
        with np.errstate(divide='ignore', invalid='ignore'):
            ratios['AR_Turnover'] = df_data['Sales'] / df_data['AR']
            ratios['AR_Turnover'] = ratios['AR_Turnover'].replace([np.inf, -np.inf], np.nan)
    
    # AP Turnover = COGS / AP
    if 'COGS' in df_data and 'AP' in df_data:
        with np.errstate(divide='ignore', invalid='ignore'):
            ratios['AP_Turnover'] = df_data['COGS'] / df_data['AP']
            ratios['AP_Turnover'] = ratios['AP_Turnover'].replace([np.inf, -np.inf], np.nan)
    
    # Inventory Turnover = COGS / (Inventory + WIP)
    if 'COGS' in df_data and 'Total_Inventory' in df_data:
        with np.errstate(divide='ignore', invalid='ignore'):
            ratios['Inventory_Turnover'] = df_data['COGS'] / df_data['Total_Inventory']
            ratios['Inventory_Turnover'] = ratios['Inventory_Turnover'].replace([np.inf, -np.inf], np.nan)
    
    # Debt to Equity = Total Debt / Equity
    if 'Total_Debt' in df_data and 'Equity' in df_data:
        with np.errstate(divide='ignore', invalid='ignore'):
            ratios['Debt_to_Equity'] = df_data['Total_Debt'] / df_data['Equity']
            ratios['Debt_to_Equity'] = ratios['Debt_to_Equity'].replace([np.inf, -np.inf], np.nan)
    
    # Debt Service Coverage Ratio = EBITDA / Interest Expense
    if 'EBITDA' in df_data and 'Interest_Expense' in df_data:
        with np.errstate(divide='ignore', invalid='ignore'):
            ratios['Debt_Service_Coverage'] = df_data['EBITDA'] / df_data['Interest_Expense']
            ratios['Debt_Service_Coverage'] = ratios['Debt_Service_Coverage'].replace([np.inf, -np.inf], np.nan)
    
    # Current Ratio = Current Assets / Current Liabilities
    if 'Current_Assets' in df_data and 'Current_Liabilities' in df_data:
        with np.errstate(divide='ignore', invalid='ignore'):
            ratios['Current_Ratio'] = df_data['Current_Assets'] / df_data['Current_Liabilities']
            ratios['Current_Ratio'] = ratios['Current_Ratio'].replace([np.inf, -np.inf], np.nan)
    
    # Working Capital = Current Assets - Current Liabilities
    if 'Current_Assets' in df_data and 'Current_Liabilities' in df_data:
        ratios['Working_Capital'] = df_data['Current_Assets'] - df_data['Current_Liabilities']
    
    # Quick Ratio = (Current Assets - Inventory) / Current Liabilities
    if 'Current_Assets' in df_data and 'Total_Inventory' in df_data and 'Current_Liabilities' in df_data:
        with np.errstate(divide='ignore', invalid='ignore'):
            ratios['Quick_Ratio'] = (df_data['Current_Assets'] - df_data['Total_Inventory']) / df_data['Current_Liabilities']
            ratios['Quick_Ratio'] = ratios['Quick_Ratio'].replace([np.inf, -np.inf], np.nan)
    
    return ratios


def calculate_growth_rates(series):
    """
    Calculate period-over-period growth rates.
    Returns series of growth rates and statistics dict.
    """
    growth_rates = series.pct_change()
    
    # Remove infinite and NaN values for statistics
    valid_rates = growth_rates.replace([np.inf, -np.inf], np.nan).dropna()
    
    if len(valid_rates) == 0:
        stats = {
            'mean': 0,
            'median': 0,
            'min': 0,
            'max': 0,
            'std': 0
        }
    else:
        stats = {
            'mean': valid_rates.mean(),
            'median': valid_rates.median(),
            'min': valid_rates.min(),
            'max': valid_rates.max(),
            'std': valid_rates.std()
        }
    
    return growth_rates, stats


def convert_annual_to_monthly_rate(annual_rate):
    """Convert annual growth rate to equivalent monthly rate."""
    return (1 + annual_rate) ** (1/12) - 1


def forecast_series(historical_series, num_periods, growth_rate):
    """
    Forecast future values using compound growth rate.
    
    Args:
        historical_series: pandas Series of historical data
        num_periods: number of periods to forecast
        growth_rate: growth rate per period (e.g., 0.10 for 10%)
    
    Returns:
        pandas Series of forecasted values
    """
    last_value = historical_series.iloc[-1]
    
    if pd.isna(last_value) or last_value == 0:
        # Use second-to-last or mean
        valid_values = historical_series.dropna()
        if len(valid_values) > 0:
            last_value = valid_values.iloc[-1]
        else:
            last_value = 0
    
    forecast_values = []
    for i in range(num_periods):
        last_value = last_value * (1 + growth_rate)
        forecast_values.append(last_value)
    
    return pd.Series(forecast_values)


# ============================================================================
# VISUALIZATION
# ============================================================================

def create_time_series_chart(data_dict, title, ylabel="Value", height=400):
    """Create line chart with multiple series."""
    fig = go.Figure()
    
    colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899']
    
    for idx, (name, series) in enumerate(data_dict.items()):
        fig.add_trace(go.Scatter(
            x=series.index,
            y=series.values,
            mode='lines+markers',
            name=name,
            line=dict(color=colors[idx % len(colors)], width=2),
            marker=dict(size=6)
        ))
    
    fig.update_layout(
        title=title,
        xaxis_title='Period',
        yaxis_title=ylabel,
        hovermode='x unified',
        height=height,
        template='plotly_white',
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    return fig


def create_forecast_comparison_chart(historical, forecast, metric_name, ylabel="Value"):
    """Create chart comparing historical and forecast data."""
    fig = go.Figure()
    
    # Historical
    fig.add_trace(go.Scatter(
        x=historical.index,
        y=historical.values,
        mode='lines+markers',
        name='Historical',
        line=dict(color='#3b82f6', width=3),
        marker=dict(size=8)
    ))
    
    # Forecast
    fig.add_trace(go.Scatter(
        x=forecast.index,
        y=forecast.values,
        mode='lines+markers',
        name='Forecast',
        line=dict(color='#8b5cf6', width=3, dash='dash'),
        marker=dict(size=8, symbol='diamond')
    ))
    
    # Connection line between last historical and first forecast
    if len(historical) > 0 and len(forecast) > 0:
        fig.add_trace(go.Scatter(
            x=[historical.index[-1], forecast.index[0]],
            y=[historical.values[-1], forecast.values[0]],
            mode='lines',
            line=dict(color='gray', width=1, dash='dot'),
            showlegend=False,
            hoverinfo='skip'
        ))
    
    fig.update_layout(
        title=f'{metric_name}: Historical vs Forecast',
        xaxis_title='Period',
        yaxis_title=ylabel,
        hovermode='x unified',
        height=450,
        template='plotly_white'
    )
    
    return fig


def create_ratio_dashboard(ratios_dict):
    """Create multi-panel ratio visualization."""
    # Select key ratios to display
    key_ratios = ['AR_Turnover', 'AP_Turnover', 'Debt_to_Equity', 'Current_Ratio']
    available_ratios = {k: v for k, v in ratios_dict.items() if k in key_ratios and not v.isna().all()}
    
    if not available_ratios:
        return None
    
    n_ratios = len(available_ratios)
    rows = (n_ratios + 1) // 2
    
    fig = make_subplots(
        rows=rows, 
        cols=2,
        subplot_titles=list(available_ratios.keys())
    )
    
    colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444']
    
    for idx, (name, series) in enumerate(available_ratios.items()):
        row = idx // 2 + 1
        col = idx % 2 + 1
        
        fig.add_trace(
            go.Scatter(
                x=series.index,
                y=series.values,
                mode='lines+markers',
                name=name,
                line=dict(color=colors[idx % len(colors)], width=2),
                marker=dict(size=6),
                showlegend=False
            ),
            row=row, col=col
        )
    
    fig.update_layout(
        height=300 * rows,
        template='plotly_white',
        title_text="Financial Ratios Dashboard"
    )
    
    return fig


# ============================================================================
# MAIN APPLICATION
# ============================================================================

def main():
    """Main Streamlit application."""
    
    # Header
    st.title("Chinook -- Cash Conversion Cycle & 3-Statement Model Calculator")
    st.markdown("""
    Upload your Excel financial statements for comprehensive analysis including:
    - **Cash Conversion Cycle** (CCC) calculation with DSO, DIO, DPO
    - **Financial Ratios**
    - **Growth Analysis**
    - **Forecasting**
    """)
    
    # Sidebar
    with st.sidebar:
        st.header("Guide")
        st.markdown("""
        **Excel File Format:**
        - Multiple sheets supported (will be combined)
        - Rows = Line items (Sales, COGS, etc.)
        - Columns = Time periods (years or months)
        - Auto-detects monthly/yearly data
        
        **Required Items:**
        - Sales/Revenue
        - COGS
        - Accounts Receivable
        - Accounts Payable
        - Inventory
        """)
        
        st.header("")
        show_debug = st.checkbox("Show Debug Info", value=False)
    
    # File Upload
    st.header("1️⃣ Upload Financial Data")
    uploaded_file = st.file_uploader(
        "Choose Excel file (.xlsx or .xls)",
        type=['xlsx', 'xls'],
        help="Upload your financial statements in Excel format"
    )
    
    if uploaded_file is None:
        st.info("👆 Please upload an Excel file to begin analysis")
        return
    
    # Load and process Excel
    with st.spinner("Loading Excel file..."):
        sheets_dict = load_excel_with_sheets(uploaded_file)
    
    if sheets_dict is None:
        st.error("Failed to load Excel file")
        return
    
    # Show sheet information
    st.success(f"✅ Loaded {len(sheets_dict)} sheet(s): {', '.join(sheets_dict.keys())}")
    
    # Combine sheets
    with st.spinner("Processing and combining sheets..."):
        df = combine_sheets(sheets_dict)
    
    if df is None:
        st.error("Could not process Excel sheets. Please check your file format.")
        if show_debug:
            st.write("Debug: Showing first sheet raw data")
            st.dataframe(list(sheets_dict.values())[0].head(20))
        return
    
    # Detect time period frequency
    if hasattr(df, 'attrs') and 'parsed_dates' in df.attrs:
        dates = df.attrs['parsed_dates']
        frequency, avg_days = detect_time_frequency(dates)
        
        if avg_days:
            st.info(f"**Detected Time Period**: {frequency.upper()}")
        else:
            st.info(f"**Detected Time Period**: {frequency.upper()}")
        
        # Set period days based on frequency
        if frequency == 'monthly':
            period_days = 30
        elif frequency == 'quarterly':
            period_days = 90
        elif frequency == 'yearly':
            period_days = 365
        else:
            st.warning(f"⚠️ Could not confidently detect frequency. Average spacing: {avg_days:.0f} days" if avg_days else "⚠️ Could not detect frequency")
            period_days = st.number_input("Days per period", value=int(avg_days) if avg_days else 365, min_value=1)
    else:
        st.warning("⚠️ Could not confidently parse dates from columns")
        period_days = st.number_input("Days per period", value=365, min_value=1)
        frequency = 'unknown'
    
    # Show raw data
    with st.expander("📄 View Processed Data", expanded=False):
        st.dataframe(df, use_container_width=True)
        
        if show_debug:
            st.write("**Column Info:**")
            st.write(f"Columns: {list(df.columns)}")
            st.write(f"Index: {list(df.index[:10])}..." if len(df.index) > 10 else f"Index: {list(df.index)}")
            if hasattr(df, 'attrs') and 'parsed_dates' in df.attrs:
                st.write(f"Parsed dates: {df.attrs['parsed_dates']}")
    
    # Column selection
    st.header("2️⃣ Select Time Periods")
    
    all_columns = list(df.columns)
    
    st.info("**Historical Periods**")
    default_historical = all_columns
    
    historical_cols = st.multiselect(
        "Select historical periods",
        options=all_columns,
        default=default_historical,
        help="Choose columns containing actual historical data"
    )
    
    if not historical_cols:
        st.warning("Please select at least one time period")
        return
    
    # Auto-detect variables
    st.header("3️⃣ Map Financial Statement Items")
    
    with st.spinner("Auto-detecting line items..."):
        auto_mapping = auto_detect_variables(df)
    
    # Show mapping status
    detected_count = sum(1 for v in auto_mapping.values() if v is not None)
    st.info(f"Auto-detected {detected_count} out of 23 line items")
    
    # Variable mapping interface
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Income Statement")
        
        sales_row = st.selectbox(
            "Sales/Revenue *",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Sales"]) + 1 if auto_mapping["Sales"] and auto_mapping["Sales"] in df.index else 0,
            help="Total revenue or sales"
        )
        
        cogs_row = st.selectbox(
            "Cost of Goods Sold (COGS) *",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["COGS"]) + 1 if auto_mapping["COGS"] and auto_mapping["COGS"] in df.index else 0,
            help="Direct costs of producing goods/services"
        )
        
        gross_profit_row = st.selectbox(
            "Gross Profit",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Gross Profit"]) + 1 if auto_mapping["Gross Profit"] and auto_mapping["Gross Profit"] in df.index else 0
        )
        
        opex_row = st.selectbox(
            "Operating Expenses",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Operating Expenses"]) + 1 if auto_mapping["Operating Expenses"] and auto_mapping["Operating Expenses"] in df.index else 0
        )
        
        depreciation_row = st.selectbox(
            "Depreciation & Amortization",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Depreciation"]) + 1 if auto_mapping["Depreciation"] and auto_mapping["Depreciation"] in df.index else 0
        )
        
        interest_row = st.selectbox(
            "Interest Expense",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Interest Expense"]) + 1 if auto_mapping["Interest Expense"] and auto_mapping["Interest Expense"] in df.index else 0
        )
        
        net_income_row = st.selectbox(
            "Net Income",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Net Income"]) + 1 if auto_mapping["Net Income"] and auto_mapping["Net Income"] in df.index else 0
        )
    
    with col2:
        st.subheader("Balance Sheet - Assets")
        
        cash_row = st.selectbox(
            "Cash & Equivalents",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Cash"]) + 1 if auto_mapping["Cash"] and auto_mapping["Cash"] in df.index else 0
        )
        
        ar_row = st.selectbox(
            "Accounts Receivable *",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Accounts Receivable"]) + 1 if auto_mapping["Accounts Receivable"] and auto_mapping["Accounts Receivable"] in df.index else 0,
            help="Money owed by customers"
        )
        
        inventory_row = st.selectbox(
            "Inventory *",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Inventory"]) + 1 if auto_mapping["Inventory"] and auto_mapping["Inventory"] in df.index else 0,
            help="Finished goods inventory"
        )
        
        wip_row = st.selectbox(
            "Work in Progress (WIP)",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["WIP"]) + 1 if auto_mapping["WIP"] and auto_mapping["WIP"] in df.index else 0,
            help="Partially completed goods"
        )
        
        prepaid_row = st.selectbox(
            "Prepaid Expenses",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Prepaid Expenses"]) + 1 if auto_mapping["Prepaid Expenses"] and auto_mapping["Prepaid Expenses"] in df.index else 0
        )
        
        current_assets_row = st.selectbox(
            "Total Current Assets",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Current Assets"]) + 1 if auto_mapping["Current Assets"] and auto_mapping["Current Assets"] in df.index else 0
        )
        
        fixed_assets_row = st.selectbox(
            "Fixed Assets (Net)",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Fixed Assets"]) + 1 if auto_mapping["Fixed Assets"] and auto_mapping["Fixed Assets"] in df.index else 0
        )
        
        total_assets_row = st.selectbox(
            "Total Assets",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Total Assets"]) + 1 if auto_mapping["Total Assets"] and auto_mapping["Total Assets"] in df.index else 0
        )
    
    with col3:
        st.subheader("Balance Sheet - Liabilities & Equity")
        
        ap_row = st.selectbox(
            "Accounts Payable *",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Accounts Payable"]) + 1 if auto_mapping["Accounts Payable"] and auto_mapping["Accounts Payable"] in df.index else 0,
            help="Money owed to suppliers"
        )
        
        deferred_income_row = st.selectbox(
            "Deferred Income",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Deferred Income"]) + 1 if auto_mapping["Deferred Income"] and auto_mapping["Deferred Income"] in df.index else 0,
            help="Cash received but not yet earned"
        )
        
        accrued_row = st.selectbox(
            "Accrued Expenses",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Accrued Expenses"]) + 1 if auto_mapping["Accrued Expenses"] and auto_mapping["Accrued Expenses"] in df.index else 0
        )
        
        short_debt_row = st.selectbox(
            "Short Term Debt",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Short Term Debt"]) + 1 if auto_mapping["Short Term Debt"] and auto_mapping["Short Term Debt"] in df.index else 0
        )
        
        current_liab_row = st.selectbox(
            "Total Current Liabilities",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Current Liabilities"]) + 1 if auto_mapping["Current Liabilities"] and auto_mapping["Current Liabilities"] in df.index else 0
        )
        
        long_debt_row = st.selectbox(
            "Long Term Debt",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Long Term Debt"]) + 1 if auto_mapping["Long Term Debt"] and auto_mapping["Long Term Debt"] in df.index else 0
        )
        
        total_liab_row = st.selectbox(
            "Total Liabilities",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Total Liabilities"]) + 1 if auto_mapping["Total Liabilities"] and auto_mapping["Total Liabilities"] in df.index else 0
        )
        
        equity_row = st.selectbox(
            "Total Equity",
            options=[""] + list(df.index),
            index=list(df.index).index(auto_mapping["Equity"]) + 1 if auto_mapping["Equity"] and auto_mapping["Equity"] in df.index else 0
        )
    
    # Validate required fields
    required_fields = {
        'Sales': sales_row,
        'COGS': cogs_row,
        'Accounts Receivable': ar_row,
        'Accounts Payable': ap_row,
        'Inventory': inventory_row
    }
    
    missing_fields = [name for name, value in required_fields.items() if not value]
    
    if missing_fields:
        st.error(f"⚠️ Please map the following required fields: {', '.join(missing_fields)}")
        return
    
    # Calculate button
    st.header("4️⃣ Calculate Metrics")
    
    if st.button("Calculate CCC & Financial Metrics", type="primary", use_container_width=True):
        
        with st.spinner("Calculating financial metrics..."):
            
            # Extract historical data
            hist_df = df[historical_cols]
            
            # Build data dictionary
            data = {}
            
            # Income Statement
            data['Sales'] = get_row_data(hist_df, sales_row)
            data['COGS'] = get_row_data(hist_df, cogs_row)
            data['Gross_Profit'] = get_row_data(hist_df, gross_profit_row) if gross_profit_row else data['Sales'] - data['COGS']
            data['Operating_Expenses'] = get_row_data(hist_df, opex_row)
            data['Depreciation'] = get_row_data(hist_df, depreciation_row)
            data['Interest_Expense'] = get_row_data(hist_df, interest_row)
            data['Net_Income'] = get_row_data(hist_df, net_income_row) if net_income_row else data['Gross_Profit'] - data['Operating_Expenses'] - data['Depreciation'] - data['Interest_Expense']
            
            # Calculate EBITDA
            data['EBITDA'] = data['Net_Income'] + data['Interest_Expense'] + data['Depreciation']
            
            # Balance Sheet - Assets
            data['Cash'] = get_row_data(hist_df, cash_row)
            data['AR'] = get_row_data(hist_df, ar_row)
            data['Inventory'] = get_row_data(hist_df, inventory_row)
            data['WIP'] = get_row_data(hist_df, wip_row)
            data['Total_Inventory'] = data['Inventory'] + data['WIP']
            data['Deferred_Income'] = get_row_data(hist_df, deferred_income_row)
            data['Prepaid'] = get_row_data(hist_df, prepaid_row)
            data['Current_Assets'] = get_row_data(hist_df, current_assets_row) if current_assets_row else data['Cash'] + data['AR'] + data['Total_Inventory'] + data['Prepaid']
            data['Fixed_Assets'] = get_row_data(hist_df, fixed_assets_row)
            data['Total_Assets'] = get_row_data(hist_df, total_assets_row) if total_assets_row else data['Current_Assets'] + data['Fixed_Assets']
            
            # Balance Sheet - Liabilities & Equity
            data['AP'] = get_row_data(hist_df, ap_row)
            data['Accrued'] = get_row_data(hist_df, accrued_row)
            data['Short_Term_Debt'] = get_row_data(hist_df, short_debt_row)
            data['Current_Liabilities'] = get_row_data(hist_df, current_liab_row) if current_liab_row else data['AP'] + data['Accrued'] + data['Short_Term_Debt']
            data['Long_Term_Debt'] = get_row_data(hist_df, long_debt_row)
            data['Total_Debt'] = data['Short_Term_Debt'] + data['Long_Term_Debt']
            data['Total_Liabilities'] = get_row_data(hist_df, total_liab_row) if total_liab_row else data['Current_Liabilities'] + data['Long_Term_Debt']
            data['Equity'] = get_row_data(hist_df, equity_row) if equity_row else data['Total_Assets'] - data['Total_Liabilities']
            
            # Calculate CCC metrics
            ccc_metrics = calculate_ccc_metrics(
                data['Sales'], 
                data['COGS'], 
                data['AR'], 
                data['AP'], 
                data['Inventory'], 
                data['WIP'], 
                data['Deferred_Income'],
                period_days
            )
            
            data.update(ccc_metrics)
            
            # Calculate financial ratios
            ratios = calculate_financial_ratios(data)
            
            # Calculate growth rates
            sales_growth, sales_growth_stats = calculate_growth_rates(data['Sales'])
            cogs_growth, cogs_growth_stats = calculate_growth_rates(data['COGS'])
            
            # Store in session state
            st.session_state['data'] = data
            st.session_state['ratios'] = ratios
            st.session_state['historical_cols'] = historical_cols
            st.session_state['period_days'] = period_days
            st.session_state['frequency'] = frequency
            st.session_state['sales_growth_stats'] = sales_growth_stats
            st.session_state['cogs_growth_stats'] = cogs_growth_stats
            
            # Store mappings for forecast
            st.session_state['mappings'] = {
                'sales_row': sales_row,
                'cogs_row': cogs_row,
                'ar_row': ar_row,
                'ap_row': ap_row,
                'inventory_row': inventory_row,
                'wip_row': wip_row,
                'deferred_income_row': deferred_income_row,
                'net_income_row': net_income_row,
                'depreciation_row': depreciation_row,
                'opex_row': opex_row,
                'equity_row': equity_row,
                'long_debt_row': long_debt_row,
                'short_debt_row': short_debt_row
            }
        
        st.success("✅ Calculations complete!")
        #st.rerun()
    
    # Display results if available
    if 'data' not in st.session_state:
        return
    
    data = st.session_state['data']
    ratios = st.session_state['ratios']
    historical_cols = st.session_state['historical_cols']
    period_days = st.session_state['period_days']
    frequency = st.session_state['frequency']
    sales_growth_stats = st.session_state['sales_growth_stats']
    cogs_growth_stats = st.session_state['cogs_growth_stats']
    
    # ========================================================================
    # RESULTS SECTION
    # ========================================================================
    
    st.header("5️⃣ Financial Analysis")
    
    # Key Metrics Dashboard
    st.subheader("Key Metrics Summary")
    
    # Get latest values
    latest_idx = -1
    latest_period = data['Sales'].index[latest_idx]
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "Cash Conversion Cycle",
            f"{data['CCC'].iloc[latest_idx]:.1f} days"
        )
        st.metric(
            "Working Capital",
            f"${ratios['Working_Capital'].iloc[latest_idx]:,.0f}" if 'Working_Capital' in ratios else "N/A"
        )
    
    with col2:
        st.metric(
            "DIO (Days Inventory Outstanding)",
            f"{data['DIO'].iloc[latest_idx]:.1f} days"
        )
        st.metric(
            "Debt Service Coverage",
            f"{ratios['Debt_Service_Coverage'].iloc[latest_idx]:.2f}" if 'Debt_Service_Coverage' in ratios and not pd.isna(ratios['Debt_Service_Coverage'].iloc[latest_idx]) else "N/A"
        )
    
    with col3:
        st.metric(
            "DSO (Days Sales Outstanding)",
            f"{data['DSO'].iloc[latest_idx]:.1f} days"
        )
        st.metric(
            "Current Ratio",
            f"{ratios['Current_Ratio'].iloc[latest_idx]:.2f}" if 'Current_Ratio' in ratios and not pd.isna(ratios['Current_Ratio'].iloc[latest_idx]) else "N/A"
        )
    
    with col4:
        st.metric(
            "DPO (Days Payable Outstanding)",
            f"{data['DPO'].iloc[latest_idx]:.1f} days"
        )
        st.metric(
            "Debt to Equity",
            f"{ratios['Debt_to_Equity'].iloc[latest_idx]:.2f}" if 'Debt_to_Equity' in ratios and not pd.isna(ratios['Debt_to_Equity'].iloc[latest_idx]) else "N/A"
        )
    
    # Tabs for different analyses
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 CCC Analysis",
        "📊 Financial Ratios",
        "📈 Growth Analysis",
        "📈 Forecasting"
    ])
    
    # ========================================================================
    # TAB 1: CCC ANALYSIS
    # ========================================================================
    
    with tab1:
        st.subheader("Cash Conversion Cycle Components")
        
        # CCC trend chart
        ccc_chart_data = {
            'DSO': data['DSO'],
            'DIO': data['DIO'],
            'DPO': data['DPO'],
            'CCC': data['CCC']
        }
        
        fig_ccc = create_time_series_chart(
            ccc_chart_data,
            "Cash Conversion Cycle Trends",
            "Days",
            height=500
        )
        st.plotly_chart(fig_ccc, use_container_width=True)
        
        # CCC statistics
        st.subheader("Cash Conversion Cycle Statistics")
        
        col1, col2 = st.columns(2)
        
        with col1:
            ccc_stats = pd.DataFrame({
                'Metric': ['DSO', 'DIO', 'DPO', 'CCC'],
                'Current': [
                    data['DSO'].iloc[latest_idx],
                    data['DIO'].iloc[latest_idx],
                    data['DPO'].iloc[latest_idx],
                    data['CCC'].iloc[latest_idx]
                ],
                'Average': [
                    data['DSO'].mean(),
                    data['DIO'].mean(),
                    data['DPO'].mean(),
                    data['CCC'].mean()
                ],
                'Min': [
                    data['DSO'].min(),
                    data['DIO'].min(),
                    data['DPO'].min(),
                    data['CCC'].min()
                ],
                'Max': [
                    data['DSO'].max(),
                    data['DIO'].max(),
                    data['DPO'].max(),
                    data['CCC'].max()
                ]
            })
            
            st.dataframe(
                ccc_stats.style.format({
                    'Current': '{:.1f}',
                    'Average': '{:.1f}',
                    'Min': '{:.1f}',
                    'Max': '{:.1f}'
                }),
                use_container_width=True,
                hide_index=True
            )
        
        # with col2:
        #     st.markdown("**CCC Interpretation:**")
        #     st.markdown("""
        #     - **DSO**: Average days to collect receivables
        #     - **DIO**: Average days inventory is held
        #     - **DPO**: Average days to pay suppliers
        #     - **CCC = DSO + DIO - DPO**: Net days cash is tied up
            
        #     **Lower CCC is generally better** (faster cash conversion)
        #     """)
        
        # Working Capital breakdown
        st.subheader("Working Capital Components")
        
        wc_data = {
            'Accounts Receivable': data['AR'],
            'Inventory + WIP': data['Total_Inventory'],
            'Accounts Payable': data['AP']
        }
        
        fig_wc = go.Figure()
        
        fig_wc.add_trace(go.Bar(
            x=data['AR'].index,
            y=data['AR'].values,
            name='Accounts Receivable',
            marker_color='#3b82f6'
        ))
        
        fig_wc.add_trace(go.Bar(
            x=data['Total_Inventory'].index,
            y=data['Total_Inventory'].values,
            name='Inventory + WIP',
            marker_color='#10b981'
        ))
        
        fig_wc.add_trace(go.Bar(
            x=data['AP'].index,
            y=-data['AP'].values,
            name='Accounts Payable',
            marker_color='#f59e0b'
        ))
        
        # Net working capital line
        net_wc = data['AR'] + data['Total_Inventory'] - data['AP']
        fig_wc.add_trace(go.Scatter(
            x=net_wc.index,
            y=net_wc.values,
            name='AR + Inv - AP',
            mode='lines+markers',
            line=dict(color='#ef4444', width=3)#,
            #yaxis='y2'
        ))
        
        fig_wc.update_layout(
            title='Working Capital Components',
            xaxis_title='Period',
            yaxis_title='Amount ($)',
            yaxis2=dict(
                title='',
                overlaying='y',
                side='right'
            ),
            barmode='relative',
            height=500,
            template='plotly_white',
            hovermode='x unified'
        )
        
        st.plotly_chart(fig_wc, use_container_width=True)
    
    # ========================================================================
    # TAB 2: FINANCIAL RATIOS
    # ========================================================================
    
    with tab2:
        st.subheader("Financial Ratios")
        
        # Ratios dashboard
        if ratios:
            fig_ratios = create_ratio_dashboard(ratios)
            if fig_ratios:
                st.plotly_chart(fig_ratios, use_container_width=True)
        
        # Detailed ratios table
        st.subheader("Ratios Over Time")
        
        ratios_df = pd.DataFrame(ratios)
        
        if not ratios_df.empty:
            # Format and display
            st.dataframe(
                ratios_df.style.format("{:.2f}"),
                use_container_width=True
            )
        else:
            st.info("Not enough data to calculate ratios. Please map more balance sheet items.")
    
    # ========================================================================
    # TAB 3: GROWTH ANALYSIS
    # ========================================================================
    
    with tab3:
        st.subheader("Growth Rate Analysis")
        
        # Variable selection
        st.markdown("#### Select Variables to Analyze")
        
        # Available numeric variables
        available_vars = {
            'Sales': data['Sales'],
            'COGS': data['COGS'],
            'Gross Profit': data['Gross_Profit'],
            'Operating Expenses': data['Operating_Expenses'],
            'Net Income': data['Net_Income'],
            'EBITDA': data['EBITDA'],
            'Accounts Receivable': data['AR'],
            'Inventory': data['Total_Inventory'],
            'Accounts Payable': data['AP'],
            'Working Capital': ratios.get('Working_Capital', pd.Series()),
            'Total Assets': data['Total_Assets'],
            'Equity': data['Equity']
        }
        
        # Filter out empty series
        available_vars = {k: v for k, v in available_vars.items() if not v.empty and not v.isna().all()}
        
        col1, col2 = st.columns(2)
        
        with col1:
            var1_name = st.selectbox(
                "Select first variable",
                options=list(available_vars.keys()),
                index=0 if 'Sales' in available_vars else 0
            )
        st.rerun()
        
        with col2:
            var2_name = st.selectbox(
                "Select second variable",
                options=list(available_vars.keys()),
                index=1 if len(available_vars) > 1 else 0
            )
        st.rerun()
        
        # Calculate growth rates for selected variables
        var1_data = available_vars[var1_name]
        var2_data = available_vars[var2_name]
        
        var1_growth, var1_stats = calculate_growth_rates(var1_data)
        var2_growth, var2_stats = calculate_growth_rates(var2_data)
        
        # Display side by side
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"### 📊 {var1_name} Growth")
            
            growth_metrics_1 = pd.DataFrame({
                'Metric': ['Mean Growth', 'Median Growth', 'Min Growth', 'Max Growth', 'Std Dev'],
                'Value': [
                    f"{var1_stats['mean']*100:.2f}%",
                    f"{var1_stats['median']*100:.2f}%",
                    f"{var1_stats['min']*100:.2f}%",
                    f"{var1_stats['max']*100:.2f}%",
                    f"{var1_stats['std']*100:.2f}%"
                ]
            })
            
            st.dataframe(growth_metrics_1, use_container_width=True, hide_index=True)
            
            # Trend chart
            fig_var1 = go.Figure()
            fig_var1.add_trace(go.Scatter(
                x=var1_data.index,
                y=var1_data.values,
                mode='lines+markers',
                name=var1_name,
                line=dict(color='#3b82f6', width=3)
            ))
            fig_var1.update_layout(
                title=f'{var1_name} Trend',
                xaxis_title='Period',
                yaxis_title=var1_name,
                height=400,
                template='plotly_white'
            )
            st.plotly_chart(fig_var1, use_container_width=True)
            
            # Growth rate chart
            fig_growth1 = go.Figure()
            fig_growth1.add_trace(go.Bar(
                x=var1_growth.index,
                y=var1_growth.values * 100,
                name=f'{var1_name} Growth %',
                marker_color='#3b82f6'
            ))
            fig_growth1.update_layout(
                title=f'{var1_name} Period-over-Period Growth',
                xaxis_title='Period',
                yaxis_title='Growth Rate (%)',
                height=350,
                template='plotly_white'
            )
            st.plotly_chart(fig_growth1, use_container_width=True)
        
        with col2:
            st.markdown(f"### 📊 {var2_name} Growth")
            
            growth_metrics_2 = pd.DataFrame({
                'Metric': ['Mean Growth', 'Median Growth', 'Min Growth', 'Max Growth', 'Std Dev'],
                'Value': [
                    f"{var2_stats['mean']*100:.2f}%",
                    f"{var2_stats['median']*100:.2f}%",
                    f"{var2_stats['min']*100:.2f}%",
                    f"{var2_stats['max']*100:.2f}%",
                    f"{var2_stats['std']*100:.2f}%"
                ]
            })
            
            st.dataframe(growth_metrics_2, use_container_width=True, hide_index=True)
            
            # Trend chart
            fig_var2 = go.Figure()
            fig_var2.add_trace(go.Scatter(
                x=var2_data.index,
                y=var2_data.values,
                mode='lines+markers',
                name=var2_name,
                line=dict(color='#f59e0b', width=3)
            ))
            fig_var2.update_layout(
                title=f'{var2_name} Trend',
                xaxis_title='Period',
                yaxis_title=var2_name,
                height=400,
                template='plotly_white'
            )
            st.plotly_chart(fig_var2, use_container_width=True)
            
            # Growth rate chart
            fig_growth2 = go.Figure()
            fig_growth2.add_trace(go.Bar(
                x=var2_growth.index,
                y=var2_growth.values * 100,
                name=f'{var2_name} Growth %',
                marker_color='#f59e0b'
            ))
            fig_growth2.update_layout(
                title=f'{var2_name} Period-over-Period Growth',
                xaxis_title='Period',
                yaxis_title='Growth Rate (%)',
                height=350,
                template='plotly_white'
            )
            st.plotly_chart(fig_growth2, use_container_width=True)
        
        # Comparison chart
        st.markdown("### 📈 Growth Rate Comparison")
        
        fig_compare = go.Figure()
        
        fig_compare.add_trace(go.Scatter(
            x=var1_growth.index,
            y=var1_growth.values * 100,
            mode='lines+markers',
            name=f'{var1_name} Growth %',
            line=dict(color='#3b82f6', width=2)
        ))
        
        fig_compare.add_trace(go.Scatter(
            x=var2_growth.index,
            y=var2_growth.values * 100,
            mode='lines+markers',
            name=f'{var2_name} Growth %',
            line=dict(color='#f59e0b', width=2)
        ))
        
        fig_compare.update_layout(
            title=f'{var1_name} vs {var2_name} Growth Rates',
            xaxis_title='Period',
            yaxis_title='Growth Rate (%)',
            height=450,
            template='plotly_white',
            hovermode='x unified'
        )
        
        st.plotly_chart(fig_compare, use_container_width=True)
        
        # Correlation analysis if both variables have valid data
        if len(var1_data) > 2 and len(var2_data) > 2:
            st.markdown("### Relationship Analysis")
            
            # Align the two series
            aligned_data = pd.DataFrame({
                var1_name: var1_data,
                var2_name: var2_data
            }).dropna()
            
            if len(aligned_data) > 2:
                correlation = aligned_data.corr().iloc[0, 1]
                
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    st.metric("Correlation Coefficient", f"{correlation:.3f}")
                    
                    if abs(correlation) > 0.7:
                        st.success("Strong correlation")
                    elif abs(correlation) > 0.4:
                        st.info("Moderate correlation")
                    else:
                        st.warning("Weak correlation")
                
                with col2:
                    # Scatter plot
                    fig_scatter = go.Figure()
                    fig_scatter.add_trace(go.Scatter(
                        x=aligned_data[var2_name],
                        y=aligned_data[var1_name],
                        mode='markers',
                        marker=dict(size=10, color='#8b5cf6'),
                        text=aligned_data.index,
                        hovertemplate=f'<b>%{{text}}</b><br>{var2_name}: %{{x:,.0f}}<br>{var1_name}: %{{y:,.0f}}<extra></extra>'
                    ))
                    
                    fig_scatter.update_layout(
                        title=f'{var1_name} vs {var2_name}',
                        xaxis_title=var2_name,
                        yaxis_title=var1_name,
                        height=350,
                        template='plotly_white'
                    )
                    
                    st.plotly_chart(fig_scatter, use_container_width=True)
    
    # ========================================================================
    # TAB 4: FORECASTING
    # ========================================================================
    
    with tab4:
        st.subheader("Financial Forecasting")
        
        st.markdown("""
        Generate forecasts based on historical growth rates and operational assumptions.
        Adjust the parameters below to customize your forecast.
        """)
        
        # Forecast parameters
        st.markdown("#### Forecast Parameters")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            num_periods = st.number_input(
                f"Number of periods to forecast",
                min_value=1,
                max_value=24,
                value=6 if frequency == 'monthly' else 3,
                help=f"Number of {frequency} periods to forecast"
            )
            
            # Convert growth rate based on frequency
            if frequency == 'monthly':
                default_sales_growth = sales_growth_stats['mean']
            elif frequency == 'yearly':
                default_sales_growth = sales_growth_stats['mean']
            else:
                default_sales_growth = sales_growth_stats['mean']
            
            sales_growth_rate = st.number_input(
                f"Sales Growth Rate per {frequency} period (%)",
                min_value=-50.0,
                max_value=100.0,
                value=float(default_sales_growth * 100) if not np.isnan(default_sales_growth) else 5.0,
                step=0.5,
                help=f"Expected growth rate for each {frequency} period"
            ) / 100
            
            current_cogs_pct = (data['COGS'].iloc[-1] / data['Sales'].iloc[-1] * 100) if data['Sales'].iloc[-1] != 0 else 65.0
            cogs_pct_of_sales = st.number_input(
                "COGS as % of Sales",
                min_value=0.0,
                max_value=100.0,
                value=float(current_cogs_pct) if not np.isnan(current_cogs_pct) else 65.0,
                step=0.5
            ) / 100

        with col2:
            # Safe value extraction with bounds
            current_dso = float(data['DSO'].iloc[-1]) if not pd.isna(data['DSO'].iloc[-1]) else 45.0
            current_dso = max(0.0, min(current_dso, 365.0))
            
            forecast_dso = st.number_input(
                "Target DSO (days)",
                min_value=0.0,
                max_value=365.0,
                value=current_dso,
                step=1.0,
                help="Days Sales Outstanding target"
            )
            
            current_dio = float(data['DIO'].iloc[-1]) if not pd.isna(data['DIO'].iloc[-1]) else 90.0
            current_dio = max(0.0, min(current_dio, 365.0))
            
            forecast_dio = st.number_input(
                "Target DIO (days)",
                min_value=0.0,
                max_value=365.0,
                value=current_dio,
                step=1.0,
                help="Days Inventory Outstanding target"
            )
            
            current_dpo = float(data['DPO'].iloc[-1]) if not pd.isna(data['DPO'].iloc[-1]) else 30.0
            current_dpo = max(0.0, min(current_dpo, 365.0))
            
            forecast_dpo = st.number_input(
                "Target DPO (days)",
                min_value=0.0,
                max_value=365.0,
                value=current_dpo,
                step=1.0,
                help="Days Payable Outstanding target"
            )
        
        with col3:
            opex_growth_rate = st.number_input(
                f"Operating Expenses Growth Rate per period (%)",
                min_value=-50.0,
                max_value=100.0,
                value=float(default_sales_growth * 80) if not np.isnan(default_sales_growth) else 4.0,
                step=0.5
            ) / 100
            
            depreciation_annual = st.number_input(
                "Depreciation",
                min_value=0.0,
                #value=float(data['Depreciation'].sum()) if data['Depreciation'].sum() > 0 else 10000.0,
                value=float(data['Depreciation'].mean()) if data['Depreciation'].mean() > 0 else 10000.0,
                step=1000.0
            )
            
            # Calculate depreciation per period
            if frequency == 'monthly':
                depreciation_per_period = depreciation_annual / 12
            elif frequency == 'quarterly':
                depreciation_per_period = depreciation_annual / 4
            else:
                depreciation_per_period = depreciation_annual
            
            interest_rate_pct = st.number_input(
                "Interest Rate on Debt (% per year)",
                min_value=0.0,
                max_value=50.0,
                value=5.0,
                step=0.25
            ) / 100
        
        # Generate forecast button
        if st.button("Generate Forecast", type="primary", use_container_width=True):
            
            with st.spinner("Generating forecast..."):
                
                # Forecast sales
                fcst_sales = forecast_series(data['Sales'], num_periods, sales_growth_rate)
                
                # Generate period labels
                last_date_str = str(data['Sales'].index[-1])
                
                if frequency == 'monthly':
                    # Parse the last date to get starting point
                    try:
                        last_parsed = parse_date_flexible(last_date_str)
                        if last_parsed:
                            fcst_periods = pd.date_range(start=last_parsed, periods=num_periods + 1, freq='MS')[1:]
                            fcst_index = [d.strftime('%b %Y') for d in fcst_periods]
                        else:
                            fcst_index = [f"Fcst {i+1}" for i in range(num_periods)]
                    except:
                        fcst_index = [f"Fcst {i+1}" for i in range(num_periods)]
                
                elif frequency == 'yearly':
                    # Extract year and increment
                    try:
                        last_parsed = parse_date_flexible(last_date_str)
                        if last_parsed:
                            last_year = last_parsed.year
                            fcst_index = [f"{last_year + i + 1}" for i in range(num_periods)]
                        else:
                            # Try to extract year from string
                            year_match = re.search(r'(\d{4})', last_date_str)
                            if year_match:
                                last_year = int(year_match.group(1))
                                fcst_index = [f"{last_year + i + 1}" for i in range(num_periods)]
                            else:
                                fcst_index = [f"Year {i+1}" for i in range(num_periods)]
                    except:
                        fcst_index = [f"Year {i+1}" for i in range(num_periods)]
                
                else:
                    fcst_index = [f"Fcst {i+1}" for i in range(num_periods)]
                
                fcst_sales.index = fcst_index
                
                # Forecast COGS
                fcst_cogs = fcst_sales * cogs_pct_of_sales
                
                # Forecast OpEx
                fcst_opex = forecast_series(data['Operating_Expenses'], num_periods, opex_growth_rate)
                fcst_opex.index = fcst_index
                
                # Forecast Depreciation
                fcst_depreciation = pd.Series([depreciation_per_period] * num_periods, index=fcst_index)
                
                # Calculate working capital based on CCC targets
                fcst_ar = (fcst_sales / 365) * forecast_dso
                fcst_inventory = (fcst_cogs / 365) * forecast_dio
                fcst_ap = (fcst_cogs / 365) * forecast_dpo
                
                # Calculate income statement items
                fcst_gross_profit = fcst_sales - fcst_cogs
                
                # Interest expense (on existing debt)
                avg_debt = data['Total_Debt'].iloc[-1] if 'Total_Debt' in data else 0
                if frequency == 'monthly':
                    fcst_interest = pd.Series([avg_debt * interest_rate_pct / 12] * num_periods, index=fcst_index)
                elif frequency == 'quarterly':
                    fcst_interest = pd.Series([avg_debt * interest_rate_pct / 4] * num_periods, index=fcst_index)
                else:
                    fcst_interest = pd.Series([avg_debt * interest_rate_pct] * num_periods, index=fcst_index)
                
                fcst_net_income = fcst_gross_profit - fcst_opex - fcst_depreciation - fcst_interest
                
                # Calculate EBITDA
                fcst_ebitda = fcst_net_income + fcst_interest + fcst_depreciation
                
                # Calculate CCC
                fcst_dso = pd.Series([forecast_dso] * num_periods, index=fcst_index)
                fcst_dio = pd.Series([forecast_dio] * num_periods, index=fcst_index)
                fcst_dpo = pd.Series([forecast_dpo] * num_periods, index=fcst_index)
                fcst_ccc = fcst_dso + fcst_dio - fcst_dpo
                
                # Store forecast in session state
                forecast_data = {
                    'Sales': fcst_sales,
                    'COGS': fcst_cogs,
                    'Gross_Profit': fcst_gross_profit,
                    'Operating_Expenses': fcst_opex,
                    'Depreciation': fcst_depreciation,
                    'Interest_Expense': fcst_interest,
                    'Net_Income': fcst_net_income,
                    'EBITDA': fcst_ebitda,
                    'AR': fcst_ar,
                    'Inventory': fcst_inventory,
                    'AP': fcst_ap,
                    'DSO': fcst_dso,
                    'DIO': fcst_dio,
                    'DPO': fcst_dpo,
                    'CCC': fcst_ccc
                }
                
                st.session_state['forecast_data'] = forecast_data
                st.session_state['forecast_periods'] = num_periods
            
            st.success("✅ Forecast generated successfully!")
            st.rerun()
        
        # Display forecast if available
        if 'forecast_data' in st.session_state:
            forecast_data = st.session_state['forecast_data']
            
            st.markdown("---")
            st.subheader("Forecast Results")
            
            # Forecast summary metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                fcst_sales_growth = ((forecast_data['Sales'].iloc[-1] / data['Sales'].iloc[-1]) - 1) * 100
                st.metric(
                    "Total Sales Growth",
                    f"{fcst_sales_growth:.1f}%",
                    delta=f"${forecast_data['Sales'].iloc[-1] - data['Sales'].iloc[-1]:,.0f}"
                )
            
            with col2:
                avg_fcst_margin = (forecast_data['Gross_Profit'].mean() / forecast_data['Sales'].mean() * 100)
                current_margin = (data['Gross_Profit'].iloc[-1] / data['Sales'].iloc[-1] * 100)
                st.metric(
                    "Avg Forecast Gross Margin",
                    f"{avg_fcst_margin:.1f}%",
                    delta=f"{avg_fcst_margin - current_margin:.1f}% vs current"
                )
            
            with col3:
                st.metric(
                    "Forecast CCC",
                    f"{forecast_data['CCC'].iloc[-1]:.1f} days",
                    delta=f"{forecast_data['CCC'].iloc[-1] - data['CCC'].iloc[-1]:.1f} vs current"
                )
            
            with col4:
                total_fcst_net_income = forecast_data['Net_Income'].sum()
                st.metric(
                    "Total Forecast Net Income",
                    f"${total_fcst_net_income:,.0f}"
                )
            
            # Forecast charts
            st.markdown("### Forecast Figures")
            
            # Sales forecast
            fig_sales_fcst = create_forecast_comparison_chart(
                data['Sales'],
                forecast_data['Sales'],
                'Sales',
                'Sales ($)'
            )
            st.plotly_chart(fig_sales_fcst, use_container_width=True)
            
            # CCC forecast
            fig_ccc_fcst = create_forecast_comparison_chart(
                data['CCC'],
                forecast_data['CCC'],
                'Cash Conversion Cycle',
                'Days'
            )
            st.plotly_chart(fig_ccc_fcst, use_container_width=True)
            
            # Net Income forecast
            fig_ni_fcst = create_forecast_comparison_chart(
                data['Net_Income'],
                forecast_data['Net_Income'],
                'Net Income',
                'Net Income ($)'
            )
            st.plotly_chart(fig_ni_fcst, use_container_width=True)
            
            # Combined historical + forecast table
            st.markdown("### Complete Forecast Table")
            
            # Combine historical and forecast
            combined_df = pd.DataFrame()
            
            for key in ['Sales', 'COGS', 'Gross_Profit', 'Operating_Expenses', 
                       'Depreciation', 'Interest_Expense', 'Net_Income', 
                       'AR', 'Inventory', 'AP', 'DSO', 'DIO', 'DPO', 'CCC']:
                if key in data and key in forecast_data:
                    combined_df[key] = pd.concat([data[key], forecast_data[key]])
            
            # Add type indicator
            combined_df['Type'] = ['Historical'] * len(data['Sales']) + ['Forecast'] * len(forecast_data['Sales'])
            
            # Reorder columns
            cols = ['Type'] + [col for col in combined_df.columns if col != 'Type']
            combined_df = combined_df[cols]
            
            # Format and display
            st.dataframe(
                combined_df.style.format({
                    'Sales': '${:,.0f}',
                    'COGS': '${:,.0f}',
                    'Gross_Profit': '${:,.0f}',
                    'Operating_Expenses': '${:,.0f}',
                    'Depreciation': '${:,.0f}',
                    'Interest_Expense': '${:,.0f}',
                    'Net_Income': '${:,.0f}',
                    'AR': '${:,.0f}',
                    'Inventory': '${:,.0f}',
                    'AP': '${:,.0f}',
                    'DSO': '{:.1f}',
                    'DIO': '{:.1f}',
                    'DPO': '{:.1f}',
                    'CCC': '{:.1f}'
                }),
                use_container_width=True
            )


# ============================================================================
# RUN APPLICATION
# ============================================================================

if __name__ == "__main__":
    main()