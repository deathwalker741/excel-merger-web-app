import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io

# Configure Streamlit page
st.set_page_config(
    page_title="Excel School Merger",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for better UI
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 1rem 0;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 2rem;
    }
    .stFileUploader > div > div > div > div {
        text-align: center;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
        margin: 1rem 0;
    }
    .stats-container {
        display: flex;
        justify-content: space-around;
        margin: 2rem 0;
    }
    .stat-box {
        text-align: center;
        padding: 1rem;
        background: #f8f9fa;
        border-radius: 0.5rem;
        border: 1px solid #dee2e6;
    }
    .stat-number {
        font-size: 2rem;
        font-weight: bold;
        color: #667eea;
    }
    .stat-label {
        color: #6c757d;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# Fields to aggregate using sum
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

@st.cache_data
def clean_currency(value):
    """Clean and convert currency strings to float"""
    if pd.isnull(value):
        return 0.0
    try:
        value_str = str(value)
        value_str = re.sub(r'[^\d.]', '', value_str.replace(',', ''))
        return float(value_str) if value_str else 0.0
    except:
        return 0.0

def merge_group(group: pd.DataFrame) -> pd.Series:
    """Merge rows grouped by School No"""
    merged_row = {}
    
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

@st.cache_data
def process_excel_data(file_content):
    """Process Excel data with caching for speed"""
    try:
        # Read Excel directly from memory (no file I/O!)
        df = pd.read_excel(io.BytesIO(file_content), dtype=str)
        
        # Validate required column
        if 'School No' not in df.columns:
            return None, "❌ Excel file must contain a 'School No' column"
        
        # Fast processing with optimized groupby
        with st.spinner('🔄 Merging duplicate schools...'):
            merged_df = df.groupby("School No", as_index=False).apply(
                merge_group, include_groups=False
            )
        
        # Format numeric columns
        for col in SUM_COLUMNS:
            if col in merged_df.columns:
                merged_df[col] = pd.to_numeric(
                    merged_df[col], errors='coerce'
                ).fillna(0).round(2)
        
        return merged_df, None
        
    except Exception as e:
        return None, f"❌ Error processing file: {str(e)}"

def create_download_excel(df):
    """Create Excel file in memory for download"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Merged Schools')
    
    # Add timestamp to filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"merged_schools_{timestamp}.xlsx"
    
    return output.getvalue(), filename

# Main App UI
def main():
    # Header
    st.markdown('<h1 class="main-header">📊 Excel School Merger</h1>', unsafe_allow_html=True)
    st.markdown("**Upload your Excel file to merge duplicate school records instantly**")
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx)",
        type=['xlsx'],
        help="Upload an Excel file with school data containing a 'School No' column"
    )
    
    if uploaded_file is not None:
        # Show file info
        file_size = len(uploaded_file.getvalue()) / 1024  # KB
        st.info(f"📁 **File:** {uploaded_file.name} ({file_size:.1f} KB)")
        
        # Process button
        if st.button("🚀 Process & Merge Schools", type="primary", use_container_width=True):
            
            # Start timer
            start_time = datetime.now()
            
            # Process the file
            merged_df, error = process_excel_data(uploaded_file.getvalue())
            
            # Calculate processing time
            processing_time = (datetime.now() - start_time).total_seconds()
            
            if error:
                st.markdown(f'<div class="error-box">{error}</div>', unsafe_allow_html=True)
            else:
                # Success message with stats
                original_rows = len(pd.read_excel(io.BytesIO(uploaded_file.getvalue())))
                merged_rows = len(merged_df)
                
                st.markdown(f'''
                <div class="success-box">
                    ✅ <strong>Success!</strong> Processed in {processing_time:.2f} seconds<br>
                    📊 Merged {original_rows} rows into {merged_rows} unique schools
                </div>
                ''', unsafe_allow_html=True)
                
                # Display stats
                st.markdown(f'''
                <div class="stats-container">
                    <div class="stat-box">
                        <div class="stat-number">{original_rows}</div>
                        <div class="stat-label">Original Rows</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-number">{merged_rows}</div>
                        <div class="stat-label">Merged Schools</div>
                    </div>
                    <div class="stat-box">
                        <div class="stat-number">{processing_time:.1f}s</div>
                        <div class="stat-label">Processing Time</div>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
                
                # Preview data
                st.subheader("📋 Preview of Merged Data")
                st.dataframe(merged_df.head(10), use_container_width=True)
                
                # Download button
                excel_data, filename = create_download_excel(merged_df)
                st.download_button(
                    label="📥 Download Merged Excel File",
                    data=excel_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
    
    # Information section
    with st.expander("ℹ️ How it works"):
        st.markdown("""
        **This tool merges duplicate school records by School Number:**
        
        1. **Upload** your Excel (.xlsx) file
        2. **Groups** all rows by "School No"
        3. **Sums** financial fields (revenues, student counts)
        4. **Merges** text fields intelligently
        5. **Download** the processed file instantly
        
        **Fields that get summed:**
        - Total Order Value (Exclusive/Inclusive GST)
        - ASSET Revenue & Students
        - CARES Revenue & Students  
        - Mindspark Revenue & Students
        """)

    # Footer
    st.markdown("---")
    st.markdown("**⚡ Built with Streamlit for blazing-fast performance**")

if __name__ == "__main__":
    main()
