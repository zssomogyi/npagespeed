import streamlit as st
import pandas as pd
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import numpy as np
from datetime import datetime

# Define site sections
SITE_SECTIONS = [
    '/networks/', '/about-us/', '/about-us/news/releases/', 
    '/customer-success/', '/industries/', '/we-are-nokia/', 
    '/thought-leadership/', '/blog/'
]

# Function to load CSV file
def load_file(uploaded_file, version):
    df = pd.read_csv(uploaded_file)
    if 'Address' not in df.columns:
        st.error(f"The 'Address' column is missing in the file: {uploaded_file.name}")
        return None
    df.columns = [f'{col} v{version}' if col != 'Address' else col for col in df.columns]
    return df

# Function to load Excel file
def load_xls_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    if 'Detailed Metrics' not in xls.sheet_names:
        st.error(f"The 'Detailed Metrics' sheet is missing in the file: {uploaded_file.name}")
        return None
    df = pd.read_excel(xls, sheet_name='Detailed Metrics')
    return df

# Function to merge dataframes
def merge_dataframes(dfs):
    merged_df = dfs[0]
    for df in dfs[1:]:
        merged_df = pd.merge(merged_df, df, on='Address', how='outer', suffixes=('', '_dup'))
        for col in merged_df.columns:
            if col.endswith('_dup'):
                base_col = col.rstrip('_dup')
                if base_col in merged_df.columns:
                    merged_df.drop(columns=col, inplace=True)
                else:
                    merged_df.rename(columns={col: base_col}, inplace=True)
    return merged_df

# Function to insert min, max, and median columns
def insert_min_max_median(df, metric, versions):
    version_columns = [f'{metric} v{i}' for i in versions if f'{metric} v{i}' in df.columns]
    if version_columns:
        # Convert to numeric, coercing errors to NaN
        numeric_df = df[version_columns].apply(pd.to_numeric, errors='coerce')
        
        df[f'MIN {metric}'] = numeric_df.min(axis=1)
        df[f'MAX {metric}'] = numeric_df.max(axis=1)
        df[f'MEDIAN {metric}'] = numeric_df.median(axis=1)
    else:
        st.warning(f"Version columns for {metric} are missing or not named correctly.")
    return df

# Function to reorder columns
def reorder_columns(df, metrics, versions):
    ordered_columns = ['Address']
    for metric in metrics:
        version_columns = [f'{metric} v{i}' for i in versions if f'{metric} v{i}' in df.columns]
        ordered_columns.extend(version_columns)
        if version_columns:
            ordered_columns.extend([f'MIN {metric}', f'MAX {metric}', f'MEDIAN {metric}'])
    return df[ordered_columns]

# Function to calculate sitewide averages
def calculate_sitewide_averages(df, metrics):
    averages = {}
    for metric in metrics:
        median_column = f'MEDIAN {metric}'
        if median_column in df.columns:
            averages[metric] = df[median_column].mean()
    averages_df = pd.DataFrame.from_dict(averages, orient='index', columns=['Sitewide Average'])
    averages_df.index.name = 'Metric'
    return averages_df.reset_index()

# Function to calculate section averages
def calculate_section_averages(df, metrics):
    section_averages = {}
    for section in SITE_SECTIONS:
        section_df = df[df['Address'].str.contains(section)]
        if not section_df.empty:
            section_averages[section] = calculate_sitewide_averages(section_df, metrics)
            section_averages[section].rename(columns={'Sitewide Average': 'Segment Average'}, inplace=True)
    return section_averages

# Function to convert dataframes to Excel
def to_excel(df, df_averages, section_averages):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Detailed Metrics', index=False)
        df_averages.to_excel(writer, sheet_name='Sitewide Averages', index=False)
        for section, averages in section_averages.items():
            sheet_name = section.replace('/', '_').strip('_')
            averages.to_excel(writer, sheet_name=sheet_name, index=False)
    processed_data = output.getvalue()
    return processed_data

# Function to clean data
def clean_data(df):
    return df.replace([np.inf, -np.inf, np.nan], None)

# Function to append data to Google Sheets
def append_to_gsheets(df, sheet_name):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gsheets"], scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open('[Nokia] Page_Speed_GSheet')
        
        try:
            sheet = spreadsheet.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            sheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")
        
        df_cleaned = clean_data(df)
        
        # Convert Timestamp to string
        df_cleaned['Date'] = df_cleaned['Date'].astype(str)
        
        # Ensure all values are JSON compliant
        df_cleaned = df_cleaned.applymap(lambda x: None if isinstance(x, (float, int)) and (np.isnan(x) or np.isinf(x)) else x)
        
        # Append new data to the existing data
        sheet.append_rows(df_cleaned.values.tolist(), value_input_option='RAW')
    except gspread.exceptions.APIError as e:
        st.error(f"APIError: {e}")
        st.error("Please ensure the Google Drive API is enabled and the service account has the necessary permissions.")
    except ValueError as e:
        st.error(f"ValueError: {e}")
        st.error("Please ensure all data is JSON compliant.")
    except Exception as e:
        st.error(f"Unexpected error: {e}")

# Main function
def main():
    st.title('Page Speed Score Aggregation')

    # Session state for clear button
    if 'clear' not in st.session_state:
        st.session_state.clear = False

    uploaded_files = st.file_uploader("Upload your CSV files", accept_multiple_files=True, type=['csv'])
    selected_date = st.date_input("Select the date of the upload", datetime.now().date())
    previous_aggregated_file = st.file_uploader("Upload a previously aggregated sitewide sheet", type=['xls', 'xlsx'])
    
    # Checkbox to make Google Sheets upload optional
    upload_to_gsheets = st.checkbox("Upload to Google Sheets", value=True)
    
    if (uploaded_files and len(uploaded_files) == 5) or previous_aggregated_file:
        if st.button('Start Processing'):
            st.session_state.clear = False  # Reset clear state
            dataframes = []
            versions = range(1, len(uploaded_files)+1)
            
            if uploaded_files and len(uploaded_files) == 5:
                for i, uploaded_file in enumerate(uploaded_files):
                    df = load_file(uploaded_file, i+1)
                    if df is not None:
                        dataframes.append(df)
                    else:
                        return  # Stop processing if any file is missing the 'Address' column

                if len(dataframes) == 5:  # Proceed only if all files are correctly loaded
                    combined_df = merge_dataframes(dataframes)
                    metrics = [
                        "Performance Score", "Server Responds Quickly", "First Contentful Paint Time (ms)",
                        "First Meaningful Paint Time (ms)", "Max Potential First Input Delay (ms)", "Largest Contentful Paint Time (ms)",
                        "Cumulative Layout Shift", "Speed Index Time (ms)", "Total Blocking Time (ms)", "Reduce Unused JavaScript Savings (ms)",
                        "Time to Interactive (ms)", "Minify JavaScript Savings (ms)", "Total Size Savings (Bytes)", "Total Time Savings (ms)",
                        "Total Requests", "Total Page Size (Bytes)", "HTML Size (Bytes)", "HTML Count", "Image Size (Bytes)", "Image Count",
                        "CSS Size (Bytes)", "CSS Count", "JavaScript Size (Bytes)", "JavaScript Count", "Font Size (Bytes)", "Font Count",
                        "Media Size (Bytes)", "Media Count", "Other Size (Bytes)", "Other Count", "Third Party Size (Bytes)", "Third Party Count",
                        "Core Web Vitals Assessment", "CrUX Largest Contentful Paint Time (ms)", "CrUX Interaction to Next Paint (ms)",
                        "CrUX Cumulative Layout Shift", "CrUX First Contentful Paint Time (ms)", "CrUX First Input Delay Time (ms)",
                        "CrUX Time to First Byte (ms)", "CrUX Origin Largest Contentful Paint Time (ms)", "CrUX Origin Interaction to Next Paint (ms)",
                        "CrUX Origin Cumulative Layout Shift", "CrUX Origin First Contentful Paint Time (ms)", "CrUX Origin First Input Delay Time (ms)",
                        "CrUX Origin Time to First Byte (ms)", "Eliminate Render-Blocking Resources Savings (ms)", "Defer Offscreen Images Savings (ms)",
                        "Defer Offscreen Images Savings (Bytes)", "Efficiently Encode Images Savings (ms)", "Efficiently Encode Images Savings (Bytes)",
                        "Properly Size Images Savings (ms)", "Properly Size Images Savings (Bytes)", "Minify CSS Savings (ms)", "Minify CSS Savings (Bytes)",
                        "Minify JavaScript Savings (Bytes)", "Reduce Unused CSS Savings (ms)", "Reduce Unused CSS Savings (Bytes)",
                        "Reduce Unused JavaScript Savings (Bytes)", "Serve Images in Next-Gen Formats Savings (ms)", "Serve Images in Next-Gen Formats Savings (Bytes)",
                        "Enable Text Compression Savings (ms)", "Enable Text Compression Savings (Bytes)", "Preconnect to Required Origins Savings (ms)",
                        "Server Response Times (TTFB) (ms)", "Server Response Times (TTFB) Category (ms)", "Multiple Redirects Savings (ms)",
                        "Preload Key Requests Savings (ms)", "Use Video Format for Animated Images Savings (ms)", "Use Video Format for Animated Images Savings (Bytes)",
                        "Total Image Optimization Savings (ms)", "Avoid Serving Legacy JavaScript to Modern Browsers Savings (ms)", "JavaScript Execution Time (ms)",
                        "Efficient Cache Policy Savings (Bytes)", "Minimize Main-Thread Work (ms)"
                    ]

                    for metric in metrics:
                        combined_df = insert_min_max_median(combined_df, metric, versions)

                    combined_df = reorder_columns(combined_df, metrics, versions)
                    st.dataframe(combined_df)

                    sitewide_averages = calculate_sitewide_averages(combined_df, metrics)
                    st.write("Sitewide Averages:")
                    st.dataframe(sitewide_averages)

                    section_averages = calculate_section_averages(combined_df, metrics)
                    for section, averages in section_averages.items():
                        st.write(f"Averages for {section}:")
                        st.dataframe(averages)

                    # Add selected date to the data
                    date_str = selected_date.strftime("%Y-%m-%d")
                    combined_df['Date'] = date_str
                    sitewide_averages['Date'] = date_str
                    for section, averages in section_averages.items():
                        averages['Date'] = date_str

                    # Save to Google Sheets if the checkbox is checked
                    if upload_to_gsheets:
                        append_to_gsheets(sitewide_averages, 'Sitewide Averages')
                        for section, averages in section_averages.items():
                            append_to_gsheets(averages, section.replace('/', '_').strip('_'))

                    df_xlsx = to_excel(combined_df, sitewide_averages, section_averages)
                    st.session_state.df_xlsx = df_xlsx
                    st.session_state.combined_df = combined_df

            if previous_aggregated_file:
                df_previous = load_xls_file(previous_aggregated_file)
                if df_previous is not None:
                    metrics = [
                        "Performance Score", "Server Responds Quickly", "First Contentful Paint Time (ms)",
                        "Largest Contentful Paint Time (ms)", "Cumulative Layout Shift", "Speed Index Time (ms)",
                        "Total Blocking Time (ms)", "Reduce Unused JavaScript Savings (ms)", "Time to Interactive (ms)", "Minify JavaScript Savings (ms)"
                    ]

                    section_averages = calculate_section_averages(df_previous, metrics)
                    for section, averages in section_averages.items():
                        st.write(f"Averages for {section}:")
                        st.dataframe(averages)

                    # Add selected date to the data
                    date_str = selected_date.strftime("%Y-%m-%d")
                    df_previous['Date'] = date_str
                    for section, averages in section_averages.items():
                        averages['Date'] = date_str

                    # Save to Google Sheets if the checkbox is checked
                    if upload_to_gsheets:
                        for section, averages in section_averages.items():
                            append_to_gsheets(averages, section.replace('/', '_').strip('_'))

                    df_xlsx = to_excel(df_previous, pd.DataFrame(), section_averages)
                    st.session_state.df_xlsx = df_xlsx
                    st.session_state.combined_df = df_previous

        if 'df_xlsx' in st.session_state and 'combined_df' in st.session_state:
            st.download_button(label='📥 Download Full Report',
                               data=st.session_state.df_xlsx,
                               file_name='page_speed_scores_report.xlsx',
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.download_button(label='📥 Download Full URL List',
                               data=st.session_state.combined_df.to_csv(index=False).encode('utf-8'),
                               file_name='full_url_list.csv',
                               mime='text/csv')

        if st.button('Clear'):
            st.session_state.clear = True  # Set clear state to true
            st.session_state.pop('df_xlsx', None)
            st.session_state.pop('combined_df', None)
            st.experimental_rerun()

        if st.session_state.clear:
            st.info("Data cleared. You can re-upload and process new files.")

    else:
        st.error("Please upload exactly 5 CSV files or a single sitewide aggregation Excel file.")

if __name__ == "__main__":
    main()

