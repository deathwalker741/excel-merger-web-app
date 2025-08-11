import pandas as pd
import re
from datetime import datetime

# 📌 Fields to aggregate using sum
SUM_COLUMNS = [
    "Total Order Value (Exclusive GST)",
    "Total Order Value (Inclusive GST)",
    "ASSET Revenue",
    "ASSETStudents",
    "CARES Revenue",
    "CARESStudents",
    "Mindspark Revenue",
    "MindsparkStudents"
]

# 🧹 Helper: Clean and convert currency strings to float
def clean_currency(value):
    if pd.isnull(value):
        return 0.0
    try:
        # Remove INR, commas, and any non-numeric except .
        value_str = str(value)
        value_str = re.sub(r'[^\d.]', '', value_str.replace(',', ''))
        return float(value_str) if value_str else 0.0
    except:
        return 0.0

# 🧠 Merge rows grouped by 'School No'
def merge_group(group: pd.DataFrame) -> pd.Series:
    merged_row = {}

    # Loop through all columns
    for col in group.columns:
        values = group[col].dropna().astype(str).unique()

        if col in SUM_COLUMNS:
            # Sum cleaned currency fields
            merged_row[col] = group[col].apply(clean_currency).sum()
        elif col == "School No":
            # Keep one instance of School No
            merged_row[col] = group[col].iloc[0]
        else:
            if len(values) == 1:
                merged_row[col] = values[0].strip()
            else:
                unique_vals = sorted(set(v.strip() for v in values if v.strip()))
                merged_row[col] = ", ".join(unique_vals)

    return pd.Series(merged_row)

def process_file(input_path, output_path):
    """
    Process the Excel file and merge rows by School No
    """
    try:
        # Read Excel file as strings to avoid conversion issues
        df = pd.read_excel(input_path, dtype=str)
        
        # Check if 'School No' column exists
        if 'School No' not in df.columns:
            raise ValueError("Excel file must contain a 'School No' column")
        
        # Group and merge by School No (fixed pandas deprecation warning)
        merged_df = df.groupby("School No", as_index=False).apply(merge_group, include_groups=False)
        
        # Format numeric columns to 2-decimal format
        for col in SUM_COLUMNS:
            if col in merged_df.columns:
                merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce').fillna(0).round(2)
        
        # Save to output file
        merged_df.to_excel(output_path, index=False)
        
        return len(df), len(merged_df)  # Return original and merged row counts
        
    except Exception as e:
        raise Exception(f"Error processing file: {str(e)}")
