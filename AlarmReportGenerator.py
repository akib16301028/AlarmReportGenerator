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

# Function to calculate time offline smartly (minutes, hours, or days)
def calculate_time_offline(df, current_time):
    df['Last Online Time'] = pd.to_datetime(df['Last Online Time'], format='%Y-%m-%d %H:%M:%S')
    df['Hours Offline'] = (current_time - df['Last Online Time']).dt.total_seconds() / 3600

    def format_offline_duration(hours):
        if hours < 1:
            return f"{int(hours * 60)} minutes"
        elif hours < 24:
            return f"{int(hours)} hours"
        else:
            return f"{int(hours // 24)} days"

    df['Offline Duration'] = df['Hours Offline'].apply(format_offline_duration)

    return df[['Offline Duration', 'Site Alias', 'Cluster', 'Zone', 'Last Online Time']]

# Function to extract the file name's timestamp
def extract_timestamp(file_name):
    match = re.search(r'\((.*?)\)', file_name)
    if match:
        timestamp_str = match.group(1)
        timestamp_str = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', timestamp_str).replace('_', ':')
        return pd.to_datetime(timestamp_str, format='%B %d %Y, %I:%M:%S %p', errors='coerce')
    return None

# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            valid_sheet_name = re.sub(r'[\\/*?:[\]]', '_', sheet_name)[:31]
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    return output.getvalue()

# Streamlit app
st.title("StatusMatrix")

uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)

        current_time = extract_timestamp(uploaded_alarm_file.name)
        offline_time = extract_timestamp(uploaded_offline_file.name)

        # Process the Offline Report
        pivot_offline, total_offline_count = create_offline_pivot(offline_df)

        st.markdown("### Offline Report")
        st.markdown(f"<small><i>till {offline_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)
        st.markdown(f"**Total Offline Count:** {total_offline_count}")

        # Filters for Offline Report
        cluster_filter_offline = st.selectbox("Filter by Cluster", options=["All"] + offline_df['Cluster'].unique().tolist())
        zone_filter_offline = st.selectbox("Filter by Zone", options=["All"] + offline_df['Zone'].unique().tolist())
        
        if cluster_filter_offline != "All":
            offline_df = offline_df[offline_df['Cluster'] == cluster_filter_offline]
        if zone_filter_offline != "All":
            offline_df = offline_df[offline_df['Zone'] == zone_filter_offline]

        st.dataframe(pivot_offline)

        # Calculate time offline smartly using the offline time
        time_offline_df = calculate_time_offline(offline_df, offline_time)

        # Create a summary table based on offline duration
        summary_dict = {}
        for index, row in time_offline_df.iterrows():
            duration = row['Offline Duration']
            if duration not in summary_dict:
                summary_dict[duration] = []
            summary_dict[duration].append(row)

        # Prepare DataFrame for display
        summary_data = []
        for duration, sites in summary_dict.items():
            for site in sites:
                summary_data.append([duration, site['Site Alias'], site['Cluster'], site['Zone'], site['Last Online Time']])
        
        # Display the summary table
        summary_df = pd.DataFrame(summary_data, columns=["Offline Duration", "Site Name", "Cluster", "Zone", "Last Online Time"])
        st.markdown("### Summary of Offline Sites")
        st.dataframe(summary_df)

        # Check for required columns in Alarm Report
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name', 'Alarm Time']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            alarm_df = alarm_df[~alarm_df['Client'].isnull()]

            # Prepare download for Offline Report
            offline_report_data = {
                "Offline Summary": pivot_offline,
                "Offline Details": summary_df
            }
            offline_excel_data = to_excel(offline_report_data)

            st.download_button(
                label="Download Offline Report",
                data=offline_excel_data,
                file_name=f"Offline Report_{offline_time.strftime('%Y-%m-%d %H-%M-%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Add the current time to the alarm header
            st.markdown(f"### Current Alarms Report")
            st.markdown(f"<small><i>till {current_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)

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
            prioritized_alarms = sorted([alarm for alarm in alarm_names if alarm in priority_order], key=lambda x: priority_order.index(x))
            other_alarms = sorted([alarm for alarm in alarm_names if alarm not in priority_order])

            # Combine prioritized and other alarms
            all_alarm_names = prioritized_alarms + other_alarms

            # Create a dictionary to hold pivot tables and totals for download
            alarm_report_data = {}

            # Filters for each current alarm table
            for alarm_name in all_alarm_names:
                pivot_table, total_alarm_count = create_pivot_table(alarm_df, alarm_name)
                alarm_report_data[alarm_name] = pivot_table
                
                st.markdown(f"#### {alarm_name} - Total Alarms: {total_alarm_count}")

                cluster_filter = st.selectbox(f"Filter by Cluster for {alarm_name}", options=["All"] + alarm_df['Cluster'].unique().tolist())
                zone_filter = st.selectbox(f"Filter by Zone for {alarm_name}", options=["All"] + alarm_df['Zone'].unique().tolist())
                date_filter = st.date_input(f"Select Date for {alarm_name}", value=pd.to_datetime('today'))

                if cluster_filter != "All":
                    pivot_table = pivot_table[pivot_table['Cluster'] == cluster_filter]
                if zone_filter != "All":
                    pivot_table = pivot_table[pivot_table['Zone'] == zone_filter]
                if date_filter is not None:
                    pivot_table = pivot_table[pd.to_datetime(pivot_table['Alarm Time']).dt.date == date_filter]

                st.dataframe(pivot_table)

            # Prepare download for Current Alarms Report
            current_alarm_excel_data = to_excel(alarm_report_data)

            st.download_button(
                label="Download Current Alarms Report",
                data=current_alarm_excel_data,
                file_name=f"Current Alarms Report_{current_time.strftime('%Y-%m-%d %H-%M-%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
