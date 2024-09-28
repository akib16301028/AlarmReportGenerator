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
    alarm_df = df[df['Alarm Name'] == alarm_name].copy()
    
    if alarm_name == 'DCDB-01 Primary Disconnect':
        alarm_df = alarm_df[~alarm_df['RMS Station'].str.startswith('L')]
    
    pivot = pd.pivot_table(
        alarm_df,
        index=['Cluster', 'Zone'],
        columns='Client',
        values='Site Alias',
        aggfunc='count',
        fill_value=0
    )
    
    pivot = pivot.reset_index()
    client_columns = [col for col in pivot.columns if col not in ['Cluster', 'Zone']]
    pivot['Total'] = pivot[client_columns].sum(axis=1)
    
    total_row = pivot[client_columns + ['Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    total_alarm_count = pivot['Total'].iloc[-1]
    
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_alarm_count

# Function to create pivot table for offline report
def create_offline_pivot(df):
    df = df.drop_duplicates()
    
    df['Less than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'Less than 24 hours' in x else 0)
    df['More than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 24 hours' in x and '72' not in x else 0)
    df['More than 72 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 72 hours' in x else 0)
    
    pivot = df.groupby(['Cluster', 'Zone']).agg({
        'Less than 24 hours': 'sum',
        'More than 24 hours': 'sum',
        'More than 72 hours': 'sum',
        'Site Alias': 'nunique'
    }).reset_index()

    pivot = pivot.rename(columns={'Site Alias': 'Total'})
    
    total_row = pivot[['Less than 24 hours', 'More than 24 hours', 'More than 72 hours', 'Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)

    total_offline_count = int(pivot['Total'].iloc[-1])

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
        timestamp_str = match.group(1)
        return timestamp_str.replace('_', ':')
    return "Unknown Time"

# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, (df, _) in dfs_dict.items():
            valid_sheet_name = re.sub(r'[\\/*?:[\]]', '_', sheet_name)[:31]
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    return output.getvalue()

# Streamlit app
st.title("Alarm and Offline Data Pivot Table Generator")

uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)

        formatted_alarm_time = extract_timestamp(uploaded_alarm_file.name)
        formatted_offline_time = extract_timestamp(uploaded_offline_file.name)

        # Process the Offline Report
        pivot_offline, total_offline_count = create_offline_pivot(offline_df)

        st.markdown("### Offline Report")
        st.markdown(f"<small><i>till {formatted_offline_time}</i></small>", unsafe_allow_html=True)
        st.markdown(f"**Total Offline Count:** {total_offline_count}")
        st.dataframe(pivot_offline)

        # Calculate days from Last Online Time
        days_offline_df = calculate_days_from_last_online(offline_df)

        # Create a summary table based on days offline
        summary_dict = {}
        for index, row in days_offline_df.iterrows():
            days = row['Days Offline']
            if days not in summary_dict:
                summary_dict[days] = []
            summary_dict[days].append(row)

        # Prepare DataFrame for display
        summary_data = []
        for days, sites in summary_dict.items():
            for site in sites:
                summary_data.append([days, site['Site Alias'], site['Cluster'], site['Zone'], site['Last Online Time']])
        
        # Display the summary table
        summary_df = pd.DataFrame(summary_data, columns=["Days Offline", "Site Name (Site Alias)", "Cluster", "Zone", "Last Online Time"])
        st.markdown("### Summary of Offline Sites")
        st.dataframe(summary_df)

        # Check for required columns in Alarm Report
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            alarm_df = alarm_df[~alarm_df['Client'].isnull()]
            alarm_names = alarm_df['Alarm Name'].unique()

            # Define the priority order for the alarm names
            priority_order = [
                'Mains Fail',
                'Battery Low',
                'DCDB-01 Primary Disconnect',
                'PG Run',
                'MDB Fault',
                'Door Open'
            ]

            # Separate prioritized alarms from the rest
            prioritized_alarms = [name for name in priority_order if name in alarm_names]
            non_prioritized_alarms = [name for name in alarm_names if name not in priority_order]

            # Combine both lists to maintain the desired order
            ordered_alarm_names = prioritized_alarms + non_prioritized_alarms

            # Create a dictionary to store all pivot tables
            alarm_data = {}
            for alarm_name in ordered_alarm_names:
                pivot_table, total_alarm_count = create_pivot_table(alarm_df, alarm_name)
                alarm_data[alarm_name] = (pivot_table, total_alarm_count)

            # Combine all pivot tables into one Excel file
            combined_alarm_df = pd.concat([pivot_table.assign(Alarm=alarm_name) for alarm_name, (pivot_table, _) in alarm_data.items()], ignore_index=True)

            # Display each alarm report
            for alarm_name in ordered_alarm_names:
                pivot_table, total_alarm_count = alarm_data[alarm_name]
                st.markdown(f"### {alarm_name}")
                st.markdown(f"**Total Alarm Count:** {total_alarm_count}")
                st.dataframe(pivot_table)

            # Create download button for combined Alarm Report
            combined_alarm_excel_data = to_excel({f"Combined Alarm Report": (combined_alarm_df, None)})
            st.download_button(
                label="Download All Alarms Report as Excel",
                data=combined_alarm_excel_data,
                file_name=f"All_Alarms_Report_{formatted_alarm_time}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
