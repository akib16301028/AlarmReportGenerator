import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Function to extract client name from Site Alias
def extract_client(site_alias):
    match = re.search(r'\((.*?)\)', site_alias)
    return match.group(1) if match else None

# Function to create pivot table for a specific alarm
def create_pivot_table(df, alarm_name):
    # Filter for the specific alarm
    alarm_df = df[df['Alarm Name'] == alarm_name].copy()
    
    # If the alarm is 'DCDB-01 Primary Disconnect', exclude leased sites
    if alarm_name == 'DCDB-01 Primary Disconnect':
        alarm_df = alarm_df[~alarm_df['RMS Station'].str.startswith('L')]
    
    # Create pivot table: Cluster and Zone as index, Clients as columns
    pivot = pd.pivot_table(
        alarm_df,
        index=['Cluster', 'Zone'],
        columns='Client',
        values='Site Alias',
        aggfunc='count',
        fill_value=0
    )
    
    # Flatten the columns
    pivot = pivot.reset_index()
    
    # Calculate Total per row
    client_columns = [col for col in pivot.columns if col not in ['Cluster', 'Zone']]
    pivot['Total'] = pivot[client_columns].sum(axis=1)
    
    # Calculate Total per client and overall total
    total_row = pivot[client_columns + ['Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    # Append the total row
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    # Calculate Total Alarm Count
    total_alarm_count = pivot['Total'].iloc[-1]
    
    # Merge same Cluster cells (simulate merged cells)
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_alarm_count

# Function to create offline pivot table
def create_offline_pivot_table(df):
    # Create pivot table with RIO and Duration as index
    offline_pivot = pd.pivot_table(
        df,
        index=['Cluster', 'Site Alias', 'Last Online Time'],  # Group by Cluster and Site
        values='Site Alias',  # Use Site Alias for counting
        aggfunc='count',
        fill_value=0
    ).reset_index()

    # Calculate duration columns
    duration_counts = df['Duration'].value_counts().reindex(
        ['Less than 24 hours', 'More than 24 hours', 'More than 48 hours', 'More than 72 hours'],
        fill_value=0
    )

    # Add duration columns to the pivot
    for duration in duration_counts.index:
        offline_pivot[duration] = (df['Duration'] == duration).astype(int)

    # Calculate Total Offline Sites
    total_offline_count = offline_pivot['Site Alias'].count()
    
    # Merge same Cluster cells
    last_cluster = None
    for i in range(len(offline_pivot)):
        if offline_pivot.at[i, 'Cluster'] == last_cluster:
            offline_pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = offline_pivot.at[i, 'Cluster']

    return offline_pivot, total_offline_count

# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for alarm, (df, _) in dfs_dict.items():
            # Replace any characters in sheet name that are invalid in Excel
            sheet_name = re.sub(r'[\\/*?:[\]]', '_', alarm)[:31]  # Excel sheet name limit is 31 chars
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

# Streamlit app
st.title("Alarm and Offline Data Pivot Table Generator")

# Upload Excel files
current_alarms_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"], key="current_alarms")
offline_report_file = st.file_uploader("Upload Offline Report", type=["xlsx"], key="offline_report")

if current_alarms_file and offline_report_file:
    try:
        # Read the uploaded Current Alarms file, assuming headers start from row 3 (0-indexed)
        alarms_df = pd.read_excel(current_alarms_file, header=2)

        # Read the uploaded Offline Report file, assuming headers start from row 3 (0-indexed)
        offline_df = pd.read_excel(offline_report_file, header=2)

        # Extract date and time from the uploaded files' names
        alarms_file_name = current_alarms_file.name
        match = re.search(r'\((.*?)\)', alarms_file_name)
        if match:
            alarms_timestamp_str = match.group(1)  # e.g., "September 28th 2024, 2_46_02 pm"
            alarms_formatted_time = alarms_timestamp_str.replace('_', ':')  # Replace underscores with colons for display
        else:
            alarms_formatted_time = "Unknown Time"

        offline_file_name = offline_report_file.name
        match = re.search(r'\((.*?)\)', offline_file_name)
        if match:
            offline_timestamp_str = match.group(1)  # e.g., "September 28th 2024, 2_46_02 pm"
            offline_formatted_time = offline_timestamp_str.replace('_', ':')  # Replace underscores with colons for display
        else:
            offline_formatted_time = "Unknown Time"

        # Check if required columns exist in alarms data
        required_columns_alarms = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name']
        if not all(col in alarms_df.columns for col in required_columns_alarms):
            st.error(f"The uploaded Current Alarms file is missing one of the required columns: {required_columns_alarms}")
        else:
            # Add a new column for Client extracted from Site Alias
            alarms_df['Client'] = alarms_df['Site Alias'].apply(extract_client)
            
            # Drop rows where Client extraction failed
            alarms_df = alarms_df.dropna(subset=['Client'])
            
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
            all_alarms = list(alarms_df['Alarm Name'].unique())
            non_priority_alarms = [alarm for alarm in all_alarms if alarm not in priority_alarms]
            ordered_alarms = priority_alarms + sorted(non_priority_alarms)
            
            # Create offline pivot table
            offline_pivot, total_offline_count = create_offline_pivot_table(offline_df)

            # Display the Offline Report
            st.markdown("### Offline Report")
            st.markdown(f"<small><i>till {offline_formatted_time}</i></small>", unsafe_allow_html=True)
            st.markdown(f"**Total Offline Count:** {total_offline_count}")  # Total count of offline sites
            st.dataframe(offline_pivot)  # Display the offline pivot table

            # Dictionary to store pivot tables and total counts for alarms
            pivot_tables = {}
            
            for alarm in ordered_alarms:
                pivot, total_count = create_pivot_table(alarms_df, alarm)
                pivot_tables[alarm] = (pivot, total_count)
                
                # Display the alarm name
                st.markdown(f"### {alarm}")  # Main header
                # Italicized and smaller date and time
                st.markdown(f"<small><i>till {alarms_formatted_time}</i></small>", unsafe_allow_html=True)
                st.markdown(f"**Total Count:** {int(total_count)}")  # Separate line for total count
                st.dataframe(pivot)  # Display the pivot table
            
            # Create download button
            excel_data = to_excel(pivot_tables)
            st.download_button(
                label="Download All Pivot Tables as Excel",
                data=excel_data,
                file_name=f"RMS Alarm Report {alarms_timestamp_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
