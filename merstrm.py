import pandas as pd
import streamlit as st
import warnings
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
import gc
import time

# Suppress openpyxl date warnings
warnings.filterwarnings(
    "ignore",
    message="Cell .* is marked as a date but the serial value .* is outside the limits for dates",
    module="openpyxl"
)

def read_sheet_optimized(file_name, file_path, sheet_name, chunk_size=10000):
    """
    Reads a specific sheet from the Excel file using chunked reading for better memory management.
    Returns a DataFrame with an extra 'Source_File' column if the sheet exists.
    """
    try:
        # Use engine='openpyxl' for better performance with large files
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        if sheet_name in excel_file.sheet_names:
            # Read in chunks to reduce memory usage
            df = pd.read_excel(
                excel_file, 
                sheet_name=sheet_name,
                engine='openpyxl',
                nrows=chunk_size  # Read in chunks
            )
            df['Source_File'] = file_name
            return df
    except Exception as e:
        raise e
    return None

def merge_excel_files_optimized(file_paths, chunk_size=10000):
    """
    Optimized version that merges Excel files with better memory management.
    
    Parameters:
        file_paths (list): List of file paths
        chunk_size (int): Number of rows to read at once
        
    Returns:
        dict: Keys are sheet names, values are merged DataFrames.
    """
    if not file_paths:
        return {}
    
    # Get sheet names from the first file
    first_file = pd.ExcelFile(file_paths[0], engine='openpyxl')
    sheet_names = first_file.sheet_names
    
    # Dictionary to hold merged DataFrames per sheet
    merged_data = {sheet: [] for sheet in sheet_names}
    
    # Process files sequentially to avoid memory issues
    for file_path in file_paths:
        file_name = file_path.split('/')[-1] if '/' in file_path else file_path.split('\\')[-1]
        
        try:
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            
            for sheet_name in sheet_names:
                if sheet_name in excel_file.sheet_names:
                    # Read the entire sheet at once for better performance
                    df = pd.read_excel(
                        excel_file, 
                        sheet_name=sheet_name,
                        engine='openpyxl'
                    )
                    df['Source_File'] = file_name
                    merged_data[sheet_name].append(df)
                    
                    # Force garbage collection to free memory
                    gc.collect()
                    
        except Exception as e:
            st.error(f"Error reading file '{file_name}': {e}")
            continue
    
    # Concatenate DataFrames for each sheet
    final_merged_data = {}
    for sheet_name, dfs in merged_data.items():
        if dfs:
            final_merged_data[sheet_name] = pd.concat(dfs, ignore_index=True)
            # Clear the list to free memory
            merged_data[sheet_name].clear()
    
    # Force garbage collection
    gc.collect()
    
    return final_merged_data

def merge_excel_files_streamlit_optimized(uploaded_files):
    """
    Optimized version specifically for Streamlit that handles uploaded files efficiently.
    """
    if not uploaded_files:
        return {}
    
    # Get sheet names from the first file
    first_file = pd.ExcelFile(uploaded_files[0], engine='openpyxl')
    sheet_names = first_file.sheet_names
    
    # Dictionary to hold merged DataFrames per sheet
    merged_data = {sheet: [] for sheet in sheet_names}
    
    # Process files sequentially
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        
        try:
            excel_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
            
            for sheet_name in sheet_names:
                if sheet_name in excel_file.sheet_names:
                    # Read the entire sheet at once
                    df = pd.read_excel(
                        excel_file, 
                        sheet_name=sheet_name,
                        engine='openpyxl'
                    )
                    df['Source_File'] = file_name
                    merged_data[sheet_name].append(df)
                    
                    # Force garbage collection
                    gc.collect()
                    
        except Exception as e:
            st.error(f"Error reading file '{file_name}': {e}")
            continue
    
    # Concatenate DataFrames for each sheet
    final_merged_data = {}
    for sheet_name, dfs in merged_data.items():
        if dfs:
            final_merged_data[sheet_name] = pd.concat(dfs, ignore_index=True)
            merged_data[sheet_name].clear()
    
    # Force garbage collection
    gc.collect()
    
    return final_merged_data

# Streamlit application with progress tracking
st.title("🚀 Fast Excel File Merger")
st.write("Upload multiple Excel files to merge sheets with the same name efficiently.")

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx"], accept_multiple_files=True)

if st.button("🚀 Merge Files (Optimized)"):
    if not uploaded_files or len(uploaded_files) < 2:
        st.error("At least two files are required for merging.")
    else:
        try:
            # Show progress
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("Starting merge process...")
            progress_bar.progress(10)
            
            # Start timing
            start_time = time.time()
            
            status_text.text("Reading and merging files...")
            progress_bar.progress(30)
            
            # Use optimized merge function
            merged_data = merge_excel_files_streamlit_optimized(uploaded_files)
            
            progress_bar.progress(70)
            status_text.text("Creating output file...")
            
            # Write merged data to BytesIO
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, df in merged_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            output.seek(0)
            
            progress_bar.progress(100)
            
            # Calculate and display performance metrics
            end_time = time.time()
            total_time = end_time - start_time
            total_size = sum(file.size for file in uploaded_files) / (1024 * 1024)  # MB
            
            st.success(f"✅ Excel files merged successfully!")
            st.info(f"📊 Performance: {total_size:.1f}MB processed in {total_time:.2f} seconds")
            st.info(f"⚡ Speed: {total_size/total_time:.1f} MB/s")
            
            st.download_button(
                "📥 Download Merged Excel File",
                data=output,
                file_name="merged_data_optimized.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"An error occurred: {e}")
            st.exception(e) 
