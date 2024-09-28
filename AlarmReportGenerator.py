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

# Upload Excel files for both Alarm Report and Offline Report
uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)
        
        formatted_alarm_time = extract_timestamp(uploaded_alarm_file.name)
        formatted_offline_time = extract_timestamp(uploaded_offline_file.name)
        
        alarm_download_file_name = f"RMS Alarm Report {formatted_alarm_time}.xlsx"
        
        # Process the Offline Report first
        pivot_offline, total_offline_count = create_offline_pivot(offline_df)
        
        # Display header for Offline Report
        st.markdown("### Offline Report")
        st.markdown(f"<small><i>till {formatted_offline_time}</i></small>", unsafe_allow_html=True)
        st.markdown(f"**Total Offline Count:** {total_offline_count}")
        st.dataframe(pivot_offline)

        # Create download button for Offline Report
        offline_excel_data = to_excel({'Offline Report': (pivot_offline, total_offline_count)})
        st.download_button(
            label="Download Offline Report as Excel",
            data=offline_excel_data,
            file_name=f"Offline RMS Report {formatted_offline_time}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        days_offline_df = calculate_days_from_last_online(offline_df)

        # Create a summary table based on days offline
        summary_dict = {}
        for index, row in days_offline_df.iterrows():
            days = row['Days Offline']
            if days not in summary_dict:
                summary_dict[days] = []
            summary_dict[days].append(row)

        # Display the summary table side by side for Less than 1 Day and 1 Day
        st.markdown("### Days Offline Summary")
        columns = st.columns(2)
        
        # Prepare tables for Less than 1 Day and 1 Day
        less_than_1_day = summary_dict.get(0, [])
        one_day = summary_dict.get(1, [])
        
        # Display Less than 1 Day
        with columns[0]:
            st.markdown("#### Less than 1 Day")
            if less_than_1_day:
                st.markdown("Site Name (Site Alias) | Cluster | Zone | Last Online Time")
                for site in less_than_1_day:
                    st.markdown(f"{site['Site Alias']} | {site['Cluster']} | {site['Zone']} | {site['Last Online Time']}")
            else:
                st.markdown("No sites found.")

        # Display 1 Day
        with columns[1]:
            st.markdown("#### 1 Day")
            if one_day:
                st.markdown("Site Name (Site Alias) | Cluster | Zone | Last Online Time")
                for site in one_day:
                    st.markdown(f"{site['Site Alias']} | {site['Cluster']} | {site['Zone']} | {site['Last Online Time']}")
            else:
                st.markdown("No sites found.")

        # Check if required columns exist for Alarm Report
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            alarm_df = alarm_df[~alarm_df['Client'].isnull()]
            alarm_names = alarm_df['Alarm Name'].unique()
            
            # Create a dictionary to store the pivot tables and total alarm counts
            alarm_data = {}
            for alarm_name in alarm_names:
                pivot_table, total_alarm_count = create_pivot_table(alarm_df, alarm_name)
                alarm_data[alarm_name] = (pivot_table, total_alarm_count)
                
            # Combine all alarm reports into one DataFrame for downloading
            combined_alarm_data = pd.concat([table[0] for table in alarm_data.values()], keys=alarm_data.keys())
            
            # Display the overall alarm report
            st.markdown("### Current Alarms")
            st.dataframe(combined_alarm_data)
            
            # Create a single download button for the entire Current Alarms report
            combined_alarm_excel_data = to_excel({f"Combined Alarm Report {formatted_alarm_time}": (combined_alarm_data, "")})
            st.download_button(
                label="Download Current Alarms Report as Excel",
                data=combined_alarm_excel_data,
                file_name=f"Current Alarms Report {formatted_alarm_time}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
