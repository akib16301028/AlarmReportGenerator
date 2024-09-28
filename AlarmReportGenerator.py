import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Function to extract client name from Site Alias
def extract_client(site_alias):
    match = re.search(r'\((.*?)\)', site_alias)
    return match.group(1) if match else None

# Function to create pivot table for offline report
def create_offline_pivot(df):
    # Create pivot table counting occurrences based on duration
    pivot = df.groupby(['Cluster', 'Zone']).agg({
        'Duration': lambda x: (x == 'Less than 24 hours').sum(),
        'Duration': lambda x: (x == 'More than 24 hours').sum(),
        'Duration': lambda x: (x == 'More than 72 hours').sum()
    }).reset_index()

    # Calculate the grand total for each duration
    pivot['Grand Total'] = pivot[['Duration']].sum(axis=1)

    # Add a total row for each column
    total_row = pivot[['Duration']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    pivot = pd.concat([pivot, total_row], ignore_index=True)

    # Rename columns for clarity
    pivot.columns = ['Cluster', 'Zone', 'Less than 24 hours', 'More than 24 hours', 'More than 72 hours', 'Grand Total']

    # Merge same Cluster cells (simulate merged cells)
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']

    return pivot

# Function to extract the file name's timestamp
def extract_timestamp(file_name):
    match = re.search(r'\((.*?)\)', file_name)
    if match:
        timestamp_str = match.group(1)  # e.g., "September 28th 2024, 10_26_23 pm"
        return timestamp_str.replace('_', ':')  # Replace underscores with colons for display
    return "Unknown Time"

# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            # Replace any characters in sheet name that are invalid in Excel
            valid_sheet_name = re.sub(r'[\\/*?:[\]]', '_', sheet_name)[:31]  # Excel sheet name limit is 31 chars
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    return output.getvalue()

# Streamlit app
st.title("Alarm and Offline Data Pivot Table Generator")

# Upload Excel files for both Alarm Report and Offline Report
uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

# Process both reports only after both files are uploaded
if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        # Read the Alarm Report file, assuming headers start from row 3 (0-indexed)
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        
        # Read the Offline Report file, assuming headers start from row 3 (0-indexed)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)
        
        # Extract date and time from the uploaded file names
        formatted_alarm_time = extract_timestamp(uploaded_alarm_file.name)
        formatted_offline_time = extract_timestamp(uploaded_offline_file.name)
        
        alarm_download_file_name = f"RMS Alarm Report {formatted_alarm_time}.xlsx"
        
        # Check if required columns exist for Alarm Report
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            # Add a new column for Client extracted from Site Alias
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            
            # Drop rows where Client extraction failed
            alarm_df = alarm_df.dropna(subset=['Client'])
            
            # Define priority alarms
            priority_alarms = [
                'Mains Fail', 
                'Battery Low', 
                'DCDB-01 Primary Disconnect', 
                'MDB Fault', 
                'PG Run', 
                'Door Open'
            ]
            
            # Get unique alarms, including non-priority ones
            all_alarms = list(alarm_df['Alarm Name'].unique())
            non_priority_alarms = [alarm for alarm in all_alarms if alarm not in priority_alarms]
            ordered_alarms = priority_alarms + sorted(non_priority_alarms)
            
            # Create and display the Offline Report first
            pivot_offline = create_offline_pivot(offline_df)
            st.markdown("### Offline Report")
            st.markdown(f"<small><i>till {formatted_offline_time}</i></small>", unsafe_allow_html=True)
            st.dataframe(pivot_offline)

            # Dictionary to store pivot tables
            alarm_pivot_tables = {}
            for alarm in ordered_alarms:
                pivot, total_count = create_pivot_table(alarm_df, alarm)
                alarm_pivot_tables[alarm] = pivot

                # Display the alarm name
                st.markdown(f"### {alarm}")
                st.markdown(f"<small><i>till {formatted_alarm_time}</i></small>", unsafe_allow_html=True)
                st.markdown(f"**Total Count:** {int(total_count)}")
                
                # Display the pivot table
                st.dataframe(pivot)

            # Create download button for Alarm Report
            alarm_excel_data = to_excel(alarm_pivot_tables)
            st.download_button(
                label="Download All Alarm Pivot Tables as Excel",
                data=alarm_excel_data,
                file_name=alarm_download_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
