import pandas as pd
import streamlit as st
import warnings
import io
from concurrent.futures import ThreadPoolExecutor, as_completed

# Suppress openpyxl date warnings
warnings.filterwarnings(
    "ignore",
    message="Cell .* is marked as a date but the serial value .* is outside the limits for dates",
    module="openpyxl"
)

def read_sheet(file_name, file_bytes, sheet_name):
    """
    Reads a specific sheet from the Excel file (provided as bytes).
    Returns a DataFrame with an extra 'Source_File' column if the sheet exists.
    """
    try:
        # Create a new BytesIO stream for each thread to avoid pointer issues
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        if sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            df['Source_File'] = file_name
            return df
    except Exception as e:
        # Propagate exception so it can be caught in the main thread
        raise e
    return None

def merge_excel_files(uploaded_files_data):
    """
    Merges sheets (with the same name) from all uploaded Excel files concurrently.
    
    Parameters:
        uploaded_files_data (list): List of tuples (file_name, file_bytes)
        
    Returns:
        dict: Keys are sheet names, values are merged DataFrames.
    """
    # Read sheet names from the first file to enforce a consistent structure.
    first_file_name, first_file_bytes = uploaded_files_data[0]
    sheet_names = pd.ExcelFile(io.BytesIO(first_file_bytes)).sheet_names
    
    # Dictionary to hold lists of DataFrames per sheet name.
    results_by_sheet = {sheet: [] for sheet in sheet_names}
    futures = {}
    
    # Create a thread pool to process each (sheet, file) combination concurrently.
    with ThreadPoolExecutor() as executor:
        for sheet_name in sheet_names:
            for file_name, file_bytes in uploaded_files_data:
                future = executor.submit(read_sheet, file_name, file_bytes, sheet_name)
                futures[future] = (sheet_name, file_name)
                
        # Gather the results as they complete.
        for future in as_completed(futures):
            sheet_name, file_name = futures[future]
            try:
                result = future.result()
                if result is not None:
                    results_by_sheet[sheet_name].append(result)
            except Exception as e:
                st.error(f"Error reading sheet '{sheet_name}' from file '{file_name}': {e}")
    
    # Concatenate dataframes for each sheet.
    merged_data = {}
    for sheet_name, dfs in results_by_sheet.items():
        if dfs:
            merged_data[sheet_name] = pd.concat(dfs, ignore_index=True)
    return merged_data

# Streamlit application
st.title("Upload Excel Files")
st.write("Upload multiple Excel files to merge sheets with the same name concurrently.")

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx"], accept_multiple_files=True)

if st.button("Merge Files"):
    if not uploaded_files or len(uploaded_files) < 2:
        st.error("At least two files are required for merging.")
    else:
        try:
            # Convert uploaded files into a list of (file_name, file_bytes)
            uploaded_files_data = [(file.name, file.getvalue()) for file in uploaded_files]
            
            merged_data = merge_excel_files(uploaded_files_data)
            
            # Write merged data to a BytesIO object instead of disk
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, df in merged_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            output.seek(0)  # Reset pointer to the beginning

            st.success("Excel files merged successfully!")
            st.download_button(
                "Download Merged Excel File",
                data=output,
                file_name="merged_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"An error occurred: {e}")
