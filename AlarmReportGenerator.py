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
        # Normalize day suffixes and replace underscores with colons for time
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

        # === Offline Report Filters ===
        st.sidebar.header("Offline Report Filters")
        # Get unique clusters for filtering
        offline_clusters = offline_df['Cluster'].dropna().unique().tolist()
        selected_offline_clusters = st.sidebar.multiselect(
            "Select Clusters for Offline Report",
            options=offline_clusters,
            default=offline_clusters
        )

        # Apply cluster filter to pivot_offline
        if 'Total' in pivot_offline['Cluster'].values:
            filtered_pivot_offline = pivot_offline[
                (pivot_offline['Cluster'].isin(selected_offline_clusters)) | (pivot_offline['Cluster'] == 'Total')
            ]
        else:
            filtered_pivot_offline = pivot_offline[pivot_offline['Cluster'].isin(selected_offline_clusters)]

        # Similarly, filter the summary_df
        # Calculate summary_df after filtering pivot_offline
        summary_pivot = filtered_pivot_offline.copy()
        # Exclude the total row if present
        if 'Total' in summary_pivot['Cluster'].values:
            summary_pivot = summary_pivot[summary_pivot['Cluster'] != 'Total']
        summary_df = summary_pivot.melt(id_vars=['Cluster', 'Zone'], 
                                        value_vars=['Less than 24 hours', 'More than 24 hours', 'More than 72 hours', 'Total'],
                                        var_name='Duration',
                                        value_name='Count')
        # Aggregate to get summary
        summary_df = summary_df.groupby('Duration')['Count'].sum().reset_index()

        # Display the Offline Report
        st.markdown("### Offline Report")
        st.markdown(f"<small><i>till {offline_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)
        st.markdown(f"**Total Offline Count:** {total_offline_count}")
        st.dataframe(filtered_pivot_offline)

        # === Summary of Offline Sites Filters ===
        st.markdown("### Summary of Offline Sites")
        st.markdown(f"**Total Offline Sites:** {summary_df['Count'].sum()}")
        st.dataframe(summary_df)

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
        
        # Apply cluster filter to summary_df
        summary_df_full = pd.DataFrame(summary_data, columns=["Offline Duration", "Site Name", "Cluster", "Zone", "Last Online Time"])
        filtered_summary_df = summary_df_full[summary_df_full['Cluster'].isin(selected_offline_clusters)]

        st.markdown("### Detailed Summary of Offline Sites")
        st.dataframe(filtered_summary_df)

        # Check for required columns in Alarm Report
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name', 'Alarm Time']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            alarm_df = alarm_df[~alarm_df['Client'].isnull()]

            # Prepare download for Offline Report
            offline_report_data = {
                "Offline Summary": filtered_pivot_offline,
                "Offline Details": filtered_summary_df
            }
            offline_excel_data = to_excel(offline_report_data)

            st.download_button(
                label="Download Offline Report",
                data=offline_excel_data,
                file_name=f"Offline_Report_{offline_time.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
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
            prioritized_alarms = [name for name in priority_order if name in alarm_names]
            non_prioritized_alarms = [name for name in alarm_names if name not in priority_order]

            # Combine both lists to maintain the desired order
            ordered_alarm_names = prioritized_alarms + non_prioritized_alarms

            # Create a dictionary to store all pivot tables for current alarms
            alarm_data = {}

            # Add a time filter for the "DCDB-01 Primary Disconnect" alarm
            dcdb_time_filter = None
            if 'DCDB-01 Primary Disconnect' in ordered_alarm_names:
                dcdb_time_filter = st.sidebar.date_input(
                    "Select Date Range for DCDB-01 Primary Disconnect",
                    [],
                    key='dcdb_date_filter'
                )
                # Ensure it's a range
                if len(dcdb_time_filter) == 2:
                    dcdb_start_date, dcdb_end_date = dcdb_time_filter
                else:
                    dcdb_start_date, dcdb_end_date = None, None

            # === Current Alarms Filters ===
            # We'll create a separate filter for each alarm
            for alarm_name in ordered_alarm_names:
                st.markdown(f"#### Filters for {alarm_name}")
                # Cluster-wise filter
                alarm_clusters = alarm_df['Cluster'].dropna().unique().tolist()
                selected_alarm_clusters = st.multiselect(
                    f"Select Clusters for {alarm_name}",
                    options=alarm_clusters,
                    default=alarm_clusters,
                    key=f"cluster_filter_{alarm_name}"
                )
                # Date range filter
                alarm_dates = pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce')
                min_date = alarm_dates.min().date()
                max_date = alarm_dates.max().date()
                selected_date_range = st.date_input(
                    f"Select Date Range for {alarm_name}",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date,
                    key=f"date_filter_{alarm_name}"
                )
                if isinstance(selected_date_range, tuple) and len(selected_date_range) == 2:
                    start_date, end_date = selected_date_range
                else:
                    start_date, end_date = min_date, max_date

                # Filter the DataFrame based on selected clusters and date range
                filtered_alarm_df = alarm_df[
                    (alarm_df['Alarm Name'] == alarm_name) &
                    (alarm_df['Cluster'].isin(selected_alarm_clusters)) &
                    (pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce').dt.date >= start_date) &
                    (pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce').dt.date <= end_date)
                ]

                # If it's "DCDB-01 Primary Disconnect", apply additional filter if date range is selected
                if alarm_name == 'DCDB-01 Primary Disconnect' and dcdb_start_date and dcdb_end_date:
                    filtered_alarm_df = alarm_df[
                        (alarm_df['Alarm Name'] == alarm_name) &
                        (pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce').dt.date >= dcdb_start_date) &
                        (pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce').dt.date <= dcdb_end_date)
                    ]

                # Create pivot table for the filtered data
                pivot, total_count = create_pivot_table(filtered_alarm_df, alarm_name)
                alarm_data[alarm_name] = (pivot, total_count)

            # Display each pivot table for the current alarms
            for alarm_name, (pivot, total_count) in alarm_data.items():
                st.markdown(f"### **{alarm_name}**")
                st.markdown(f"<small><i>till {current_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)
                st.markdown(f"**Alarm Count:** {total_count}")
                st.dataframe(pivot)

            # Prepare download for Current Alarms Report only if there is data
            if alarm_data:
                # Create a dictionary with each alarm's pivot table
                current_alarm_excel_dict = {alarm_name: data[0] for alarm_name, data in alarm_data.items()}
                current_alarm_excel_data = to_excel(current_alarm_excel_dict)
                st.download_button(
                    label="Download Current Alarms Report",
                    data=current_alarm_excel_data,
                    file_name=f"Current_Alarms_Report_{current_time.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No current alarm data available for export.")
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
