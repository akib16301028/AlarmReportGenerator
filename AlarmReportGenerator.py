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
st.title("Alarm Data Pivot Table Generator")

# Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file is not None:
    try:
        # Read the uploaded Excel file, assuming headers start from row 3 (0-indexed)
        df = pd.read_excel(uploaded_file, header=2)
        
        # Extract date and time from the uploaded file name
        file_name = uploaded_file.name
        match = re.search(r'\((.*?)\)', file_name)
        if match:
            # Use the original timestamp string directly
            timestamp_str = match.group(1)  # e.g., "September 28th 2024, 2_46_02 pm"
            formatted_time = timestamp_str.replace('_', ':')  # Replace underscores with colons for display
            download_file_name = f"RMS Alarm Report {timestamp_str}.xlsx"
        else:
            formatted_time = "Unknown Time"
            download_file_name = "RMS Alarm Report Unknown Time.xlsx"

        # Check if required columns exist
        required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name']
        if not all(col in df.columns for col in required_columns):
            st.error(f"The uploaded file is missing one of the required columns: {required_columns}")
        else:
            # Add a new column for Client extracted from Site Alias
            df['Client'] = df['Site Alias'].apply(extract_client)
            
            # Drop rows where Client extraction failed
            df = df.dropna(subset=['Client'])
            
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
            all_alarms = list(df['Alarm Name'].unique())
            non_priority_alarms = [alarm for alarm in all_alarms if alarm not in priority_alarms]
            ordered_alarms = priority_alarms + sorted(non_priority_alarms)
            
            # Dictionary to store pivot tables and total counts
            pivot_tables = {}
            
            for alarm in ordered_alarms:
                pivot, total_count = create_pivot_table(df, alarm)
                pivot_tables[alarm] = (pivot, total_count)
                
                # Display the alarm name and total count
                st.markdown(f"### {alarm} till {formatted_time} (Total Count: {int(total_count)})")
                st.dataframe(pivot)  # Display the pivot table
            
            # Create download button
            excel_data = to_excel(pivot_tables)
            st.download_button(
                label="Download All Pivot Tables as Excel",
                data=excel_data,
                file_name=download_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
