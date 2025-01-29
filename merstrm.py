import pandas as pd
import os
import streamlit as st
import io

# Inject custom CSS to hide the cat symbol (or any other symbol)
st.markdown("""
    <style>
        .cat-icon {
            display: none;
        }
    </style>
""", unsafe_allow_html=True)

# Function to merge Excel files
def merge_excel_files(uploaded_files):
    merged_data = {}
    
    # Load sheet names from the first file to ensure matching structure
    excel_file1 = pd.ExcelFile(uploaded_files[0])
    sheet_names = excel_file1.sheet_names

    # Iterate through each sheet
    for sheet_name in sheet_names:
        merged_sheets = []

        # Process the same sheet across all uploaded files
        for file in uploaded_files:
            try:
                excel_file = pd.ExcelFile(file)
                if sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    df['Source_File'] = file.name  # Add a column with the source file name
                    merged_sheets.append(df)
            except Exception as e:
                st.error(f"Error reading {sheet_name} from {file.name}: {e}")

        # Merge the data for the current sheet
        if merged_sheets:
            merged_data[sheet_name] = pd.concat(merged_sheets, ignore_index=True)

    return merged_data

# Streamlit application
st.title("Upload Excel Files")
st.write("Upload multiple Excel files to merge sheets with the same name.")

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx"], accept_multiple_files=True)

if st.button("Merge Files"):
    if len(uploaded_files) < 2:
        st.error("At least two files are required for merging.")
    else:
        try:
            merged_data = merge_excel_files(uploaded_files)

            # Save each merged sheet to the output file
            output_path = os.path.join(os.getcwd(), 'merged_data.xlsx')
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in merged_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Provide download link for the merged file
            with open(output_path, "rb") as f:
                st.download_button("Download Merged Excel File", f, file_name="merged_data.xlsx")
        except Exception as e:
            st.error(f"An error occurred: {e}")
