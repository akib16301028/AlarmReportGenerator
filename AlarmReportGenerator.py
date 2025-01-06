import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

# Function to extract client name from Site Alias
def extract_client(site_alias):
    match = re.search(r'\((.*?)\)', site_alias)
    return match.group(1) if match else None

# Function to categorize Duration Slot (Hours)
def categorize_duration(hours):
    if 0 <= hours < 2:
        return '0+'
    elif 2 <= hours < 4:
        return '2+'
    elif 4 <= hours < 8:
        return '4+'
    elif hours >= 8:
        return '8+'
    else:
        return 'Unknown'

# Function to create pivot table for a specific alarm
def create_pivot_table(df, alarm_name):
    alarm_df = df[df['Alarm Name'] == alarm_name].copy()
    if alarm_name == 'DCDB-01 Primary Disconnect':
        alarm_df = alarm_df[~alarm_df['RMS Station'].str.startswith('L')]

    alarm_df['Duration Category'] = alarm_df['Duration Slot (Hours)'].apply(categorize_duration)

    pivot_total = pd.pivot_table(
        alarm_df,
        index=['Cluster', 'Zone'],
        columns='Client',
        values='Site Alias',
        aggfunc='count',
        fill_value=0
    ).reset_index()

    client_columns = [col for col in pivot_total.columns if col not in ['Cluster', 'Zone']]
    pivot_total['Total'] = pivot_total[client_columns].sum(axis=1)

    pivot_duration = pd.pivot_table(
        alarm_df,
        index=['Cluster', 'Zone'],
        columns='Duration Category',
        values='Site Alias',
        aggfunc='count',
        fill_value=0
    ).reset_index()

    pivot = pd.merge(pivot_total, pivot_duration, on=['Cluster', 'Zone'], how='left')

    for cat in ['0+', '2+', '4+', '8+']:
        if cat not in pivot.columns:
            pivot[cat] = 0

    pivot = pivot[['Cluster', 'Zone'] + client_columns + ['Total', '0+', '2+', '4+', '8+']]
    total_row = pivot.select_dtypes(include=['number']).sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    pivot = pd.concat([pivot, total_row], ignore_index=True)

    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']

    return pivot, pivot['Total'].iloc[-1]

# Function to create pivot table for offline report
def create_offline_pivot(df):
    df = df.drop_duplicates()
    df['Less than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'Less than 24 hours' in x else 0)
    df['More than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 24 hours' in x and '72' not in x else 0)
    df['More than 48 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 48 hours' in x else 0)
    df['More than 72 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 72 hours' in x else 0)

    pivot = df.groupby(['Cluster', 'Zone']).agg({
        'Less than 24 hours': 'sum',
        'More than 24 hours': 'sum',
        'More than 48 hours': 'sum',
        'More than 72 hours': 'sum',
        'Site Alias': 'nunique'
    }).reset_index()

    pivot = pivot.rename(columns={'Site Alias': 'Total'})
    total_row = pivot[['Less than 24 hours', 'More than 24 hours', 'More than 48 hours', 'More than 72 hours', 'Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    pivot = pd.concat([pivot, total_row], ignore_index=True)

    return pivot, int(pivot['Total'].iloc[-1])

# Function to calculate duration for Offline Summary
def calculate_duration(df):
    current_time = datetime.now()
    df['Last Online Time'] = pd.to_datetime(df['Last Online Time'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce')
    df['Duration'] = (current_time - df['Last Online Time']).dt.days
    return df[['Site Alias', 'Zone', 'Cluster', 'Duration']]

# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            valid_sheet_name = re.sub(r'[\\/*?:[\]]', '_', sheet_name)[:31]
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    return output.getvalue()

# Function to create site-wise log table
def create_site_wise_log(df, selected_alarm):
    if selected_alarm == "All":
        filtered_df = df.copy()
    else:
        filtered_df = df[df['Alarm Name'] == selected_alarm].copy()
    return filtered_df[['Site Alias', 'Cluster', 'Zone', 'Alarm Name', 'Alarm Time', 'Duration']].sort_values(by='Alarm Time', ascending=False)

# Streamlit app
st.title("StatusMatrix@STL")

uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file and uploaded_offline_file:
    try:
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)

        st.sidebar.header("Filters")
        offline_clusters = sorted(offline_df['Cluster'].dropna().unique().tolist())
        offline_clusters.insert(0, "All")
        selected_offline_cluster = st.sidebar.selectbox("Select Cluster", options=offline_clusters, index=0)

        st.sidebar.subheader("Current Alarms Filters")
        alarm_names = sorted(alarm_df['Alarm Name'].dropna().unique().tolist())
        alarm_names.insert(0, "All")
        selected_alarm = st.sidebar.selectbox("Select Alarm to Filter", options=alarm_names, index=0)

        pivot_offline, total_offline_count = create_offline_pivot(offline_df)
        st.markdown("### Offline Report")
        st.markdown(f"**Total Offline Count:** {total_offline_count}")
        st.dataframe(pivot_offline)

        if selected_alarm != "All":
            pivot, total_count = create_pivot_table(alarm_df, selected_alarm)
            st.markdown(f"### Current Alarms for {selected_alarm}")
            st.markdown(f"**Total Alarm Count:** {total_count}")
            st.dataframe(pivot)
    except Exception as e:
        st.error(f"Error processing files: {e}")
