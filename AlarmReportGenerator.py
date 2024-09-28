import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

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

# Function to create pivot table for offline report
def create_offline_pivot(df):
    # Remove duplicates
    df = df.drop_duplicates()
    
    # Create columns for different duration categories
    df['Less than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'Less than 24 hours' in x else 0)
    df['More than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 24 hours' in x and '72' not in x else 0)
    df['More than 72 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 72 hours' in x else 0)
    
    # Pivot table structure
    pivot = df.groupby(['Cluster', 'Zone']).agg({
        'Less than 24 hours': 'sum',
        'More than 24 hours': 'sum',
        'More than 72 hours': 'sum',
        'Site Alias': 'nunique'  # This will give the total number of unique sites
    }).reset_index()

    # Rename 'Site Alias' to 'Total'
    pivot = pivot.rename(columns={'Site Alias': 'Total'})
    
    # Add a total row for each column
    total_row = pivot[['Less than 24 hours', 'More than 24 hours', 'More than 72 hours', 'Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    # Append the total row
    pivot = pd.concat([pivot, total_row], ignore_index=True)

    # Calculate total offline count from the last cell of the Total column
    total_offline_count = int(pivot['Total'].iloc[-1])  # Get the last cell of the Total column

    # Merge same Cluster cells (simulate merged cells)
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_offline_count

# Function to calculate days from Last Online Time
def calculate_days_from_last_online(df):
    now = datetime.now()
    df['Last Online Time'] = pd.to_datetime(df['Last Online Time'], format='%d/%m/%Y %I:%M:%S %p')
    df['Days Offline'] = (now - df['Last Online Time']).dt.days
    return df[['Days Offline', 'Site Alias', 'Cluster', 'Zone', 'Last Online Time']]

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

# Define the priority order for alarm names
alarm_priority = [
    'Mains Fail',
    'Battery Low',
    'DCDB-01 Primary Disconnect',
    'MDB Fault',
    'PG Run',
    'Door Open'
    # Add more alarm names as needed...
]

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
        
        # Process the Offline Report first
        pivot_offline, total_offline_count = create_offline_pivot(offline_df)
        
        # Display header for Offline Report
        st.markdown("### Offline Report")
        st.markdown(f"<small><i>till {formatted_offline_time}</i></small>", unsafe_allow_html=True)
        st.markdown(f"**Total Offline Count:** {total_offline_count}")
        
        # Display the pivot table for Offline Report
        st.dataframe(pivot_offline)

        # Create download button for Offline Report
        offline_excel_data = to_excel({'Offline Report': pivot_offline})
        st.download_button(
            label="Download Offline Report as Excel",
            data=offline_excel_data,
            file_name=f"Offline RMS Report {formatted_offline_time}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Calculate days from Last Online Time
        days_offline_df = calculate_days_from_last_online(offline_df)
        
        # Create a summary table based on days offline
        summary_dict = {}
        for index, row in days_offline_df.iterrows():
            days = row['Days Offline']
            if days not in summary_dict:
                summary_dict[days] = []
            summary_dict[days].append(row)

        # Display each group of sites under days offline as a table
        day_keys = list(summary_dict.keys())
        for i in range(0, len(day_keys), 2):  # Loop through day_keys in pairs
            cols = st.columns(2)  # Create two columns
            for j, days in enumerate(day_keys[i:i+2]):
                with cols[j]:
                    st.markdown(f"### {days} Days")
                    
                    # Create a DataFrame for the sites
                    summary_df = pd.DataFrame(summary_dict[days], columns=['Site Alias', 'Cluster', 'Zone', 'Last Online Time'])
                    # Display the DataFrame as a table
                    st.dataframe(summary_df)

        # Process Alarms Report
        pivot_dict = {}
        total_alarm_counts = {}
        
        for alarm in alarm_priority:
            pivot, total_count = create_pivot_table(alarm_df, alarm)
            pivot_dict[alarm] = pivot
            total_alarm_counts[alarm] = total_count
            
            # Display each pivot table
            st.markdown(f"### {alarm} Alarm Report")
            st.markdown(f"**Total Count:** {total_count}")
            st.dataframe(pivot)

        # Combine all alarm pivot tables into one Excel file
        alarm_excel_data = to_excel(pivot_dict)
        st.download_button(
            label="Download Alarm Reports as Excel",
            data=alarm_excel_data,
            file_name=alarm_download_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
