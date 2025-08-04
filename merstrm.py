import pandas as pd
import streamlit as st
import warnings
import io
import time
import os

# Suppress all warnings
warnings.filterwarnings("ignore")

# Initialize session state
if 'merged_data' not in st.session_state:
    st.session_state.merged_data = None
if 'output_file' not in st.session_state:
    st.session_state.output_file = None
if 'performance_metrics' not in st.session_state:
    st.session_state.performance_metrics = None
if 'uploaded_files_hash' not in st.session_state:
    st.session_state.uploaded_files_hash = None

def get_files_hash(uploaded_files):
    """Create a hash of uploaded files to detect changes"""
    if not uploaded_files:
        return None
    # Create a more stable hash based on file names and sizes
    file_info = [(f.name, f.size) for f in uploaded_files]
    file_info.sort()  # Sort to ensure consistent hash
    return hash(tuple(file_info))

def merge_excel_files_fast(uploaded_files):
    """
    Fast Excel merger with optimized output creation
    """
    if not uploaded_files:
        return {}
    
    try:
        # Get sheet names from first file
        first_file = pd.ExcelFile(uploaded_files[0], engine='openpyxl')
        sheet_names = first_file.sheet_names
        
        merged_data = {}
        
        # Process each sheet
        for sheet_name in sheet_names:
            st.write(f"📄 Processing sheet: {sheet_name}")
            
            all_data = []
            
            # Process each file
            for uploaded_file in uploaded_files:
                try:
                    # Read the sheet
                    df = pd.read_excel(
                        uploaded_file, 
                        sheet_name=sheet_name, 
                        engine='openpyxl'
                    )
                    
                    # Add source file column
                    df['Source_File'] = uploaded_file.name
                    all_data.append(df)
                    
                    st.write(f"   ✅ Read {len(df)} rows from {uploaded_file.name}")
                    
                except Exception as e:
                    st.error(f"   ❌ Error reading {uploaded_file.name}: {str(e)}")
                    continue
            
            # Merge all data for this sheet
            if all_data:
                merged_df = pd.concat(all_data, ignore_index=True)
                merged_data[sheet_name] = merged_df
                st.success(f"   📊 Merged {len(merged_df)} total rows for {sheet_name}")
            else:
                st.warning(f"   ⚠️ No data found for sheet: {sheet_name}")
        
        return merged_data
        
    except Exception as e:
        st.error(f"Error in merge process: {str(e)}")
        return {}

def create_output_file_fast(merged_data):
    """
    Fast output file creation with progress updates
    """
    st.write("📝 Creating output file...")
    
    # Create output buffer
    output = io.BytesIO()
    
    try:
        # Use optimized Excel writer
        with pd.ExcelWriter(output, engine='openpyxl', mode='w') as writer:
            for i, (sheet_name, df) in enumerate(merged_data.items()):
                st.write(f"   📋 Writing sheet: {sheet_name} ({len(df):,} rows)")
                
                # Write sheet efficiently
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Update progress
                progress = (i + 1) / len(merged_data) * 100
                st.progress(progress / 100)
        
        # Reset buffer position
        output.seek(0)
        st.success("✅ Output file created successfully!")
        return output
        
    except Exception as e:
        st.error(f"❌ Error creating output file: {str(e)}")
        return None

# Streamlit application
st.title("⚡ Fast Excel Merger with Quick Output")
st.write("Optimized for speed with fast output file creation")

# File upload
uploaded_files = st.file_uploader(
    "Choose Excel files to merge", 
    type=["xlsx"], 
    accept_multiple_files=True,
    help="Select multiple Excel files with the same sheet structure"
)

# Check if files have changed
current_files_hash = get_files_hash(uploaded_files)
files_changed = current_files_hash != st.session_state.uploaded_files_hash

# Clear previous results if files changed
if files_changed and st.session_state.merged_data is not None:
    st.session_state.merged_data = None
    st.session_state.output_file = None
    st.session_state.performance_metrics = None
    st.session_state.uploaded_files_hash = current_files_hash

# Show file info
if uploaded_files:
    st.write("### 📁 Selected Files:")
    total_size = 0
    for i, file in enumerate(uploaded_files, 1):
        size_mb = file.size / (1024 * 1024)
        total_size += size_mb
        st.write(f"{i}. **{file.name}** ({size_mb:.1f} MB)")
    
    st.info(f"**Total size:** {total_size:.1f} MB")

# Show existing results if available
if st.session_state.merged_data is not None and st.session_state.output_file is not None:
    st.success("✅ **Previous merge results available!**")
    
    # Show performance metrics
    if st.session_state.performance_metrics:
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("⏱️ Time", f"{st.session_state.performance_metrics['time']:.1f}s")
        with col2:
            st.metric("📊 Size", f"{st.session_state.performance_metrics['size']:.1f}MB")
        with col3:
            st.metric("⚡ Speed", f"{st.session_state.performance_metrics['speed']:.1f}MB/s")
    
    # Show data summary
    total_rows = sum(len(df) for df in st.session_state.merged_data.values())
    st.info(f"📈 **Total rows merged:** {total_rows:,}")
    
    # Stable download button
    st.write("### 📥 Download Your File:")
    st.download_button(
        label="📥 Download Merged Excel File",
        data=st.session_state.output_file,
        file_name="merged_excel_files.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Click to download the merged Excel file",
        use_container_width=True
    )
    
    # Show sheet summary
    st.write("### 📋 Sheet Summary:")
    for sheet_name, df in st.session_state.merged_data.items():
        st.write(f"- **{sheet_name}**: {len(df):,} rows")
    
    # Option to re-merge
    if st.button("🔄 Re-merge Files", type="secondary"):
        st.session_state.merged_data = None
        st.session_state.output_file = None
        st.session_state.performance_metrics = None
        st.rerun()

# Merge button (only show if no results exist)
if uploaded_files and st.session_state.merged_data is None:
    if len(uploaded_files) < 2:
        st.error("❌ Please select at least 2 Excel files to merge.")
    else:
        if st.button("🚀 Start Fast Merge", type="primary"):
            try:
                # Progress container
                progress_container = st.container()
                
                with progress_container:
                    st.write("### 🔄 Processing...")
                    
                    # Start timing
                    start_time = time.time()
                    
                    # Step 1: Merge files
                    st.write("**Step 1:** Reading and merging files...")
                    merged_data = merge_excel_files_fast(uploaded_files)
                    
                    if merged_data:
                        # Step 2: Create output file
                        st.write("**Step 2:** Creating output file...")
                        output = create_output_file_fast(merged_data)
                        
                        if output:
                            # Calculate performance
                            end_time = time.time()
                            total_time = end_time - start_time
                            total_size = sum(file.size for file in uploaded_files) / (1024 * 1024)
                            speed = total_size / total_time if total_time > 0 else 0
                            
                            # Store results in session state
                            st.session_state.merged_data = merged_data
                            st.session_state.output_file = output
                            st.session_state.performance_metrics = {
                                'time': total_time,
                                'size': total_size,
                                'speed': speed
                            }
                            st.session_state.uploaded_files_hash = current_files_hash
                            
                            # Success message
                            st.success("🎉 **Merge completed successfully!**")
                            
                            # Rerun to show stable results
                            st.rerun()
                            
                        else:
                            st.error("❌ Failed to create output file.")
                            
                    else:
                        st.error("❌ No data was merged. Please check your files.")
                        
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
                st.exception(e)

 
