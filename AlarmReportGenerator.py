import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime, timedelta

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
st.set_page_config(page_title="StatusMatrix", layout="wide")
st.title("StatusMatrix")

# File Uploaders
uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        # Read Excel files with header starting at row 3 (index 2)
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploadged_offline_file, header=2)

        # Extract timestamps from file names
        current_time = extract_timestamp(uploaded_alarm_file.name)
        offline_time = extract_timestamp(uploaded_offline_file.name)

        # Define Alarm Priority Order
        priority_order = [
            'Mains Fail',
            'Battery Low',
            'DCDB-01 Primary Disconnect',
            'PG Run',
            'MDB Fault',
            'Door Open',
            'Battery Critical'
            'Water Leakage'  # Added as per user's example
        ]

        # ------------------ Offline Report Filters ------------------
        st.header("Offline Report")

        # Arrange filters in three columns
        col1, col2, col3 = st.columns(3)

        with col1:
            # Date Filter: Assuming 'Last Online Time' exists in offline_df
            if 'Last Online Time' in offline_df.columns:
                offline_df['Last Online Time'] = pd.to_datetime(offline_df['Last Online Time'], format='%Y-%m-%d %H:%M:%S')
                min_date = offline_df['Last Online Time'].min().date()
                max_date = offline_df['Last Online Time'].max().date()
                selected_date = st.select_slider(
                    "Select Date Range",
                    options=pd.date_range(min_date, max_date, freq='D'),
                    value=(min_date, max_date),
                    format="YYYY-MM-DD"
                )
                if isinstance(selected_date, tuple) and len(selected_date) == 2:
                    start_date, end_date = selected_date
                    mask = (offline_df['Last Online Time'].dt.date >= start_date.date()) & (offline_df['Last Online Time'].dt.date <= end_date.date())
                    filtered_offline_df = offline_df[mask]
                else:
                    filtered_offline_df = offline_df.copy()
            else:
                filtered_offline_df = offline_df.copy()

        with col2:
            # Zone Filter
            zones = filtered_offline_df['Zone'].dropna().unique().tolist()
            selected_zones = st.multiselect("Select Zones", options=sorted(zones), default=sorted(zones))

        with col3:
            # Cluster Filter
            clusters = filtered_offline_df['Cluster'].dropna().unique().tolist()
            selected_clusters = st.multiselect("Select Clusters", options=sorted(clusters), default=sorted(clusters))

        # Apply Zone and Cluster Filters
        filtered_offline_df = filtered_offline_df[
            (filtered_offline_df['Zone'].isin(selected_zones)) &
            (filtered_offline_df['Cluster'].isin(selected_clusters))
        ]

        # Recreate pivot table based on filters
        pivot_offline, total_offline_count = create_offline_pivot(filtered_offline_df)

        # Display Offline Report
        st.markdown(f"*till {offline_time.strftime('%Y-%m-%d %H:%M:%S')}*")
        st.markdown(f"**Total Offline Count:** {total_offline_count}")
        st.dataframe(pivot_offline)

        # Calculate time offline smartly using the offline time
        time_offline_df = calculate_time_offline(filtered_offline_df, offline_time)

        # ------------------ Summary of Offline Sites ------------------
        st.subheader("Summary of Offline Sites")
        st.dataframe(time_offline_df)

        # ------------------ Current Alarms Filters ------------------
        st.header("Current Alarms Report")

        # Arrange alarm filters in three columns
        col_a1, col_a2, col_a3, col_a4 = st.columns(4)

        with col_a1:
            # Date Filter: Assuming 'Alarm Time' exists in alarm_df
            if 'Alarm Time' in alarm_df.columns:
                alarm_df['Alarm Time'] = pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p')
                min_alarm_date = alarm_df['Alarm Time'].min().date()
                max_alarm_date = alarm_df['Alarm Time'].max().date()
                selected_alarm_date = st.select_slider(
                    "Select Alarm Date Range",
                    options=pd.date_range(min_alarm_date, max_alarm_date, freq='D'),
                    value=(min_alarm_date, max_alarm_date),
                    format="YYYY-MM-DD"
                )
                if isinstance(selected_alarm_date, tuple) and len(selected_alarm_date) == 2:
                    start_alarm_date, end_alarm_date = selected_alarm_date
                    mask_alarm_date = (alarm_df['Alarm Time'].dt.date >= start_alarm_date.date()) & (alarm_df['Alarm Time'].dt.date <= end_alarm_date.date())
                    filtered_alarm_df = alarm_df[mask_alarm_date]
                else:
                    filtered_alarm_df = alarm_df.copy()
            else:
                filtered_alarm_df = alarm_df.copy()

        with col_a2:
            # Zone Filter
            alarm_zones = filtered_alarm_df['Zone'].dropna().unique().tolist()
            selected_alarm_zones = st.multiselect("Select Alarm Zones", options=sorted(alarm_zones), default=sorted(alarm_zones))

        with col_a3:
            # Cluster Filter
            alarm_clusters = filtered_alarm_df['Cluster'].dropna().unique().tolist()
            selected_alarm_clusters = st.multiselect("Select Alarm Clusters", options=sorted(alarm_clusters), default=sorted(alarm_clusters))

        with col_a4:
            # Dropdown to select specific alarms
            alarm_names = filtered_alarm_df['Alarm Name'].unique().tolist()
            selected_alarm = st.multiselect("Select Alarms to Display", options=sorted(alarm_names), default=sorted(alarm_names))

        # Apply Zone and Cluster Filters
        filtered_alarm_df = filtered_alarm_df[
            (filtered_alarm_df['Zone'].isin(selected_alarm_zones)) &
            (filtered_alarm_df['Cluster'].isin(selected_alarm_clusters)) &
            (filtered_alarm_df['Alarm Name'].isin(selected_alarm))
        ]

        # Extract Client and filter out nulls
        filtered_alarm_df['Client'] = filtered_alarm_df['Site Alias'].apply(extract_client)
        filtered_alarm_df = filtered_alarm_df[~filtered_alarm_df['Client'].isnull()]

        # Recreate pivot tables based on filters
        # Order alarms based on priority_order
        ordered_alarm_names = [alarm for alarm in priority_order if alarm in filtered_alarm_df['Alarm Name'].unique()]
        # Append any alarms not in priority_order at the end
        other_alarms = sorted(set(filtered_alarm_df['Alarm Name'].unique()) - set(priority_order))
        ordered_alarm_names += other_alarms

        # Create a dictionary to store all pivot tables for current alarms
        alarm_data = {}

        for alarm_name in ordered_alarm_names:
            data = create_pivot_table(filtered_alarm_df, alarm_name)
            alarm_data[alarm_name] = data

        # Display each pivot table for the current alarms in priority order
        for alarm_name in ordered_alarm_names:
            pivot, total_count = alarm_data[alarm_name]
            st.markdown(f"### {alarm_name}")
            st.markdown(f"<small><i>till {current_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)
            st.markdown(f"**Alarm Count:** {total_count}")
            st.dataframe(pivot)

        # Prepare download for Current Alarms Report only if there is data
        if alarm_data:
            current_alarm_excel_data = to_excel({alarm_name: data[0] for alarm_name, data in alarm_data.items()})
            st.download_button(
                label="Download Current Alarms Report",
                data=current_alarm_excel_data,
                file_name=f"Current_Alarms_Report_{current_time.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No current alarm data available for export.")

        # Prepare download for Offline Report with filters applied
        offline_report_data = {
            "Offline Summary": pivot_offline,
            "Offline Details": time_offline_df
        }
        offline_excel_data = to_excel(offline_report_data)

        st.download_button(
            label="Download Offline Report",
            data=offline_excel_data,
            file_name=f"Offline_Report_{offline_time.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
