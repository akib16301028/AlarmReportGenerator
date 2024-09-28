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
    
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot

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
        timestamp_str = match.group(1)
        return pd.to_datetime(timestamp_str.replace('_', ':'), format='%B %dth %Y, %I:%M:%S %p')
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
st.title("Alarm and Offline Data Pivot Table Generator")

uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)

        current_time = extract_timestamp(uploaded_alarm_file.name)
        offline_time = extract_timestamp(uploaded_offline_file.name)

        # Process the Offline Report
        pivot_offline = create_offline_pivot(offline_df)

        st.markdown("### Offline Report")
        st.markdown(f"<small><i>till {offline_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)
        st.dataframe(pivot_offline)

        # Check for required columns in Alarm Report
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name', 'Alarm Time']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            alarm_df = alarm_df[~alarm_df['Client'].isnull()]

            # Prepare download for Offline Report
            offline_report_data = {
                "Offline Summary": pivot_offline
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

            # Display all alarm tables initially
            alarm_names = alarm_df['Alarm Name'].unique()

            # Create a dictionary to store all pivot tables for current alarms
            alarm_data = {}
            for alarm_name in alarm_names:
                pivot = create_pivot_table(alarm_df, alarm_name)
                alarm_data[alarm_name] = pivot

                # Display the alarm name
                st.markdown(f"### <b>{alarm_name}</b>", unsafe_allow_html=True)
                st.markdown(f"<small><i>till {current_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)
                st.markdown(f"**Alarm Count:** {len(alarm_df[alarm_df['Alarm Name'] == alarm_name])}")
                st.dataframe(pivot)

            # Dropdown for time filter
            time_filter_alarm = st.selectbox("Select Alarm for Time Filter (DCDB-01 Primary Disconnect)", ["None"] + [name for name in alarm_names if name == "DCDB-01 Primary Disconnect"])

            if time_filter_alarm == "DCDB-01 Primary Disconnect":
                # Get time filter inputs
                start_time = st.date_input("Start Date", value=pd.to_datetime("2024-09-28"))
                start_hour = st.time_input("Start Time", value=pd.to_datetime("10:00:00").time())
                end_time = st.date_input("End Date", value=pd.to_datetime("2024-09-29"))
                end_hour = st.time_input("End Time", value=pd.to_datetime("23:59:59").time())

                # Create a datetime range for filtering
                start_datetime = pd.to_datetime(f"{start_time} {start_hour}")
                end_datetime = pd.to_datetime(f"{end_time} {end_hour}")

                # Filter alarm_df for the selected time range
                filtered_alarm_df = alarm_df[(alarm_df['Alarm Name'] == "DCDB-01 Primary Disconnect") & 
                                              (pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p') >= start_datetime) & 
                                              (pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p') <= end_datetime)]

                # Display the filtered pivot table
                if not filtered_alarm_df.empty:
                    st.markdown(f"### Filtered Table for {time_filter_alarm} between {start_datetime} and {end_datetime}")
                    filtered_pivot = create_pivot_table(filtered_alarm_df, time_filter_alarm)
                    st.dataframe(filtered_pivot)
                else:
                    st.warning("No alarms found for the selected time range.")

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
