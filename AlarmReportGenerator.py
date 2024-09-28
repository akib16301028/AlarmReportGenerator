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
    total_offline_count = df['Site Alias'].nunique()
    
    pivot = df.groupby(['Cluster', 'Zone']).agg({
        'Duration': lambda x: (x == 'Less than 24 hours').sum(),
        'Duration': lambda x: (x == 'More than 24 hours').sum(),
        'Duration': lambda x: (x == 'More than 72 hours').sum(),
        'Site Alias': 'nunique'
    }).rename(columns={'Site Alias': 'Total'}).reset_index()
    
    duration_counts = df['Duration'].value_counts()
    for duration in duration_counts.index:
        pivot[duration] = (df['Duration'] == duration).astype(int)
    
    total_row = pivot[['Less than 24 hours', 'More than 24 hours', 'More than 72 hours', 'Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)

    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']

    return pivot, total_offline_count

def extract_timestamp(file_name):
    match = re.search(r'\((.*?)\)', file_name)
    if match:
        timestamp_str = match.group(1)
        return timestamp_str.replace('_', ':')
    return "Unknown Time"

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
        
        alarm_download_file_name = f"RMS Alarm Report {formatted_alarm_time}.xlsx"
        
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            alarm_df = alarm_df.dropna(subset=['Client'])
            
            priority_alarms = [
                'Mains Fail', 
                'Battery Low', 
                'DCDB-01 Primary Disconnect', 
                'MDB Fault', 
                'PG Run', 
                'Door Open'
            ]
            
            all_alarms = list(alarm_df['Alarm Name'].unique())
            non_priority_alarms = [alarm for alarm in all_alarms if alarm not in priority_alarms]
            ordered_alarms = priority_alarms + sorted(non_priority_alarms)
            
            pivot_offline, total_offline_count = create_offline_pivot(offline_df)
            st.markdown("### Offline Report")
            st.markdown(f"<small><i>till {formatted_offline_time}</i></small>", unsafe_allow_html=True)
            st.markdown(f"**Total Offline Count:** {total_offline_count}")
            st.dataframe(pivot_offline)

            alarm_pivot_tables = {}
            for alarm in ordered_alarms:
                pivot, total_count = create_pivot_table(alarm_df, alarm)
                alarm_pivot_tables[alarm] = (pivot, total_count)

                st.markdown(f"### {alarm}")
                st.markdown(f"<small><i>till {formatted_alarm_time}</i></small>", unsafe_allow_html=True)
                st.markdown(f"**Total Count:** {int(total_count)}")
                
                top_5_alarms = sorted(alarm_pivot_tables.items(), key=lambda x: x[1][1], reverse=True)[:5]
                if alarm in dict(top_5_alarms):
                    st.dataframe(pivot.style.applymap(lambda x: 'background-color: yellow' if x else '', subset=['Total']))
                else:
                    st.dataframe(pivot)

            alarm_excel_data = to_excel(alarm_pivot_tables)
            st.download_button(
                label="Download All Alarm Pivot Tables as Excel",
                data=alarm_excel_data,
                file_name=alarm_download_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
