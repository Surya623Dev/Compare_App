import streamlit as st
import pandas as pd
import json
from io import BytesIO

# Function to read Excel files
def read_excel(file, sheet_name):
    df = pd.read_excel(file, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()  # Strip any leading or trailing spaces from column names
    return df

# Function to load configuration from uploaded JSON file
def load_config(config_file):
    # Read the contents of the uploaded file
    config_contents = config_file.read()

    # Decode the bytes to string (assuming UTF-8 encoding for JSON)
    config_str = config_contents.decode('utf-8')

    # Load the JSON string into a Python dictionary
    config = json.loads(config_str)
    
    return config

# Function to normalize dates in DataFrame
def normalize_dates(df, date_columns):
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.date
    return df

# Function to compare files based on configuration
def compare_files(df1, df2, config):
    columns_to_compare = config['columns_to_compare']
    date_columns = config.get('date_columns', [])
    decimal_columns = config.get('decimal_columns', [])

    # Rename key columns to a common name
    df1 = df1.rename(columns={config['file1']['key_column']: 'KeyColumn'})
    df2 = df2.rename(columns={config['file2']['key_column']: 'KeyColumn'})

    # Standardize column names to lowercase for consistent comparison
    df1.columns = df1.columns.str.lower()
    df2.columns = df2.columns.str.lower()
    columns_to_compare = [col.lower() for col in columns_to_compare]
    date_columns = [col.lower() for col in date_columns]

    # Normalize date columns
    df1 = normalize_dates(df1, date_columns)
    df2 = normalize_dates(df2, date_columns)

    # Sort dataframes by key column
    df1 = df1.sort_values(by='keycolumn').reset_index(drop=True)
    df2 = df2.sort_values(by='keycolumn').reset_index(drop=True)

    # Merge dataframes on the key column
    merged_df = pd.merge(df1, df2, on='keycolumn', suffixes=('_file1', '_file2'), how='outer', indicator=True)
    
    # Summary of matched and unmatched records
    matched_records = merged_df[merged_df['_merge'] == 'both'].shape[0]
    not_matched_records = merged_df[merged_df['_merge'] != 'both'].shape[0]

    summary = pd.DataFrame({
        'Description': ['Matched Records', 'Not Matched Records'],
        'Count': [matched_records, not_matched_records]
    })

    # Detailed comparison
    comparison_results = []
    for _, row in merged_df.iterrows():
        if row['_merge'] == 'both':
            record_comparison = [row['keycolumn']]
            for column in columns_to_compare:
                col_file1 = f"{column}_file1"
                col_file2 = f"{column}_file2"
                if col_file1 in row and col_file2 in row:
                    val1 = str(row[col_file1]).strip()
                    val2 = str(row[col_file2]).strip()
                    if val1 == val2:
                        record_comparison.append('Matched')
                    else:
                        record_comparison.append(f"{val1} | {val2}")
                else:
                    record_comparison.append('Not Available')
            comparison_results.append(record_comparison)
    
    comparison_df = pd.DataFrame(comparison_results, columns=['KeyColumn'] + columns_to_compare)

    # Missing keys
    missing_in_file1 = merged_df[merged_df['_merge'] == 'right_only']['keycolumn'].to_frame()
    missing_in_file2 = merged_df[merged_df['_merge'] == 'left_only']['keycolumn'].to_frame()

    return summary, comparison_df, missing_in_file1, missing_in_file2

# Function to write comparison results to an Excel file
def write_to_excel(summary, comparison_df, missing_in_file1, missing_in_file2, config):
    output_path = config.get('output_path', 'comparison_report.xlsx')
    
    # Create a Pandas Excel writer using xlsxwriter as the engine
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

    # Write each DataFrame to a specific sheet
    summary.to_excel(writer, sheet_name='Summary', index=False)
    comparison_df.to_excel(writer, sheet_name='Detailed Comparison', index=False)
    missing_in_file1.to_excel(writer, sheet_name='Missing in File 1', index=False)
    missing_in_file2.to_excel(writer, sheet_name='Missing in File 2', index=False)

    # Close the Pandas Excel writer and output the Excel file
    writer.save()

    # Read the generated file as bytes to return for download
    with open(output_path, 'rb') as file:
        excel_data = file.read()

    return excel_data
    
def main():
    st.title('Excel Comparison Tool')

    # Upload files and config
    file1 = st.file_uploader("Upload File 1", type=["xlsx"])
    file2 = st.file_uploader("Upload File 2", type=["xlsx"])
    config_file = st.file_uploader("Upload Config File", type=["json"])

    if config_file:
        config = load_config(config_file)
        st.write("Config:")
        st.json(config)

    if file1 and file2 and config_file:
        df1 = pd.read_excel(file1, sheet_name=config['file1']['sheet_name'])
        df2 = pd.read_excel(file2, sheet_name=config['file2']['sheet_name'])

        # Compare files
        summary, comparison_df, missing_in_file1, missing_in_file2 = compare_files(df1, df2, config)
        
        # Display comparison results
        st.write("Summary:")
        st.write(summary)
        st.write("Detailed Comparison:")
        st.write(comparison_df)
        st.write("Missing in File 1:")
        st.write(missing_in_file1)
        st.write("Missing in File 2:")
        st.write(missing_in_file2)
        
        # Save comparison report
        if st.button("Save Comparison Report"):
            excel_data = write_to_excel(summary, comparison_df, missing_in_file1, missing_in_file2, config)
            st.download_button(label="Download Comparison Report", data=excel_data, file_name="comparison_report.xlsx")

if __name__ == "__main__":
    main()
