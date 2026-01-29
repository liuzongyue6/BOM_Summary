"""
BOM Comparison - Streamlit Page
Author: Zongyue Liu
Date: 2026-01-27

Compare two BOM Excel files and generate comparison report
"""

import streamlit as st
import sys
from pathlib import Path
from datetime import datetime
from io import BytesIO

# Add parent directory to path to import bom_comparison
sys.path.append(str(Path(__file__).parent.parent.parent))

from bom_comparison import BOMComparator
import config


# Page configuration
st.set_page_config(
    page_title="BOM Comparison",
    page_icon="🔄",
    layout="wide"
)


def initialize_comparison_state():
    """Initialize session state for comparison page"""
    if 'comparison_initialized' not in st.session_state:
        st.session_state.comparison_initialized = True
        st.session_state.comparison_complete = False
        st.session_state.comparison_result = None
        st.session_state.comparison_stats = None


def main():
    """Main comparison page"""
    initialize_comparison_state()
    
    # Header
    st.title("🔄 BOM File Comparison")
    st.markdown("""
    Compare two BOM Excel files to identify:
    - **Matching parts** with differences in Material Spec, CAD Oem Name, or Thickness (highlighted in yellow)
    - **Unique parts** in each file
    """)
    
    st.divider()
    
    # File upload section
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📁 First Excel File")
        file1 = st.file_uploader(
            "Upload first BOM file",
            type=['xlsx', 'xlsm'],
            key='file1',
            help="Upload the first Excel file for comparison"
        )
        if file1:
            st.success(f"✓ Loaded: {file1.name}")
    
    with col2:
        st.subheader("📁 Second Excel File")
        file2 = st.file_uploader(
            "Upload second BOM file",
            type=['xlsx', 'xlsm'],
            key='file2',
            help="Upload the second Excel file for comparison"
        )
        if file2:
            st.success(f"✓ Loaded: {file2.name}")
    
    st.divider()
    
    # Required columns info
    with st.expander("ℹ️ Required Columns", expanded=False):
        st.markdown("Both files must contain the following columns:")
        for col in config.COMPARISON_COLUMNS:
            st.markdown(f"- `{col}`")
    
    # Comparison button
    if file1 and file2:
        if st.button("🔍 Compare Files", type="primary", use_container_width=True):
            perform_comparison(file1, file2)
    else:
        st.info("👆 Please upload both Excel files to begin comparison")
    
    # Display results if comparison is complete
    if st.session_state.comparison_complete and st.session_state.comparison_result:
        display_results()


def perform_comparison(file1, file2):
    """Perform the BOM comparison"""
    
    with st.status("🔄 Comparing BOM files...", expanded=True) as status:
        try:
            # Read file bytes
            st.write("📖 Reading File 1...")
            file1_bytes = file1.read()
            file1.seek(0)  # Reset file pointer
            
            st.write("📖 Reading File 2...")
            file2_bytes = file2.read()
            file2.seek(0)  # Reset file pointer
            
            # Initialize comparator
            st.write("🔧 Initializing comparator...")
            comparator = BOMComparator(
                file1_bytes, file1.name,
                file2_bytes, file2.name
            )
            
            # Load data
            st.write("📊 Loading and validating data...")
            success, message = comparator.load_data()
            
            if not success:
                st.error(f"❌ Validation failed: {message}")
                status.update(label="❌ Comparison failed", state="error")
                return
            
            st.write(f"✓ File 1: {len(comparator.data1)} parts loaded")
            st.write(f"✓ File 2: {len(comparator.data2)} parts loaded")
            
            # Perform comparison
            st.write("🔍 Comparing parts...")
            success, message = comparator.compare()
            
            if not success:
                st.error(f"❌ Comparison failed: {message}")
                status.update(label="❌ Comparison failed", state="error")
                return
            
            st.write(message)
            
            # Generate report
            st.write("📝 Generating comparison report...")
            result_buffer = comparator.generate_comparison_report()
            
            # Get statistics
            stats = comparator.get_stats()
            
            # Store in session state
            st.session_state.comparison_complete = True
            st.session_state.comparison_result = result_buffer
            st.session_state.comparison_stats = stats
            st.session_state.file1_name = file1.name
            st.session_state.file2_name = file2.name
            
            status.update(label="✅ Comparison complete!", state="complete")
            st.rerun()
        
        except Exception as e:
            st.error(f"❌ Error during comparison: {str(e)}")
            status.update(label="❌ Comparison failed", state="error")


def display_results():
    """Display comparison results and download button"""
    
    st.success("✅ Comparison completed successfully!")
    
    # Display statistics
    stats = st.session_state.comparison_stats
    
    # Main totals
    col1, col2 = st.columns(2)
    with col1:
        st.metric("File 1 Total Parts", stats['total_file1'])
    with col2:
        st.metric("File 2 Total Parts", stats['total_file2'])
    
    st.divider()
    
    # Category breakdown
    st.subheader("📊 Comparison Categories")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("NoChange", stats['no_change'], 
                 help="Parts that are completely identical")
    
    with col2:
        st.metric("Part+Rev Same", stats['part_rev_no_change'],
                 help="Same Part Number and Rev, but other attributes differ")
    
    with col3:
        st.metric("Rev Changed", stats['rev_change'],
                 help="Same Part Number, different Rev")
    
    with col4:
        st.metric("Part Changed", stats['part_change'],
                 help="Parts found in only one file")
    
    st.divider()
    
    # Download section
    st.subheader("📥 Download Results")
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"compared_result_{timestamp}.xlsx"
    
    col_download, col_info = st.columns([1, 2])
    
    with col_download:
        st.download_button(
            label="⬇️ Download Comparison Report",
            data=st.session_state.comparison_result,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    
    with col_info:
        st.info("""
        **Report Contents (5 Sheets):**
        - **NoChange:** Parts completely identical
        - **Part_Rev_NoChange:** Same Part+Rev, attributes differ
        - **RevChange:** Same Part, different Rev
        - **PartChange:** Parts in only one file
        - **Summary:** All categories combined
        
        Format: File1 columns | File2 columns (side-by-side)
        """)
    
    # Comparison details
    with st.expander("📊 Detailed Comparison Information", expanded=False):
        st.markdown(f"""
        **Files Compared:**
        - File 1: `{st.session_state.file1_name}`
        - File 2: `{st.session_state.file2_name}`
        
        **Comparison Logic:**
        1. **NoChange**: Part Number + Rev + all attributes identical
        2. **Part_Rev_NoChange**: Part Number + Rev same, but Material Spec / CAD Oem Name / Thickness differ
        3. **RevChange**: Part Number same, Rev different
        4. **PartChange**: Part Number not found in the other file
        
        **Output Format:**
        - Side-by-side: File1 columns | File2 columns
        - Yellow highlighting indicates differences
        - Summary sheet combines all categories with section headers
        """)
    
    # Reset button
    st.divider()
    if st.button("🔄 Start New Comparison", use_container_width=True):
        # Clear session state
        st.session_state.comparison_complete = False
        st.session_state.comparison_result = None
        st.session_state.comparison_stats = None
        st.rerun()


if __name__ == "__main__":
    main()
