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
    
    # Exclude leased sites for DCDB-01 Primary Disconnect
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
    
    # Flatten columns
    pivot = pivot.reset_index()
    
    # Calculate totals per row
    client_columns = [col for col in pivot.columns if col not in ['Cluster', 'Zone']]
    pivot['Total'] = pivot[client_columns].sum(axis=1)
    
    # Calculate totals per client and overall total
    total_row = pivot[client_columns + ['Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    # Append total row
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    # Total alarm count
    total_alarm_count = pivot['Total'].iloc[-1]
    
    # Merge same cluster cells
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_alarm_count

# Function to create pivot for offline report
def create_offline_pivot(df_offline):
    pivot = pd.pivot_table(
        df_offline,
        index='Cluster',
        columns='Duration',
        values='Site Alias',
        aggfunc='count',
        fill_value=0
    ).reset_index()
    
    # Merging same RIO cells
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    # Calculate total column
    duration_columns = [col for col in pivot.columns if col != 'Cluster']
    pivot['Total'] = pivot[duration_columns].sum(axis=1)
    
    # Calculate total count
    total_row = pivot[duration_columns + ['Total']].sum().to_frame().T
    total_row['Cluster'] = 'Total'
    
    # Append total row
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    # Total offline count
    total_offline_count = df_offline['Site Alias'].nunique()
    
    return pivot, total_offline_count

# Convert multiple dataframes to Excel
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, (df, _) in dfs_dict.items():
            # Clean sheet name
            sheet_name = re.sub(r'[\\/*?:[\]]', '_', name)[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

# Streamlit app
st.title("Alarm and Offline Data Pivot Table Generator")

# Upload Excel files
uploaded_alarm = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm and uploaded_offline:
    try:
        # Read both files
        df_alarm = pd.read_excel(uploaded_alarm, header=2)
        df_offline = pd.read_excel(uploaded_offline, header=2)
        
        # Extract time from file names
        alarm_file_time = re.search(r'\((.*?)\)', uploaded_alarm.name).group(1).replace('_', ':')
        offline_file_time = re.search(r'\((.*?)\)', uploaded_offline.name).group(1).replace('_', ':')
        
        # Offline report pivot
        offline_pivot, total_offline_count = create_offline_pivot(df_offline)
        st.markdown(f"### Offline Report")
        st.markdown(f"<small><i>till {offline_file_time}</i></small>", unsafe_allow_html=True)
        st.markdown(f"**Total Offline Count:** {int(total_offline_count)}")
        st.dataframe(offline_pivot)

        # Alarm report pivot
        df_alarm['Client'] = df_alarm['Site Alias'].apply(extract_client)
        df_alarm = df_alarm.dropna(subset=['Client'])
        
        priority_alarms = ['Mains Fail', 'Battery Low', 'DCDB-01 Primary Disconnect', 'MDB Fault', 'PG Run', 'Door Open']
        all_alarms = list(df_alarm['Alarm Name'].unique())
        ordered_alarms = priority_alarms + sorted([alarm for alarm in all_alarms if alarm not in priority_alarms])
        
        alarm_pivots = {}
        col1, col2 = st.columns(2)  # For displaying tables side by side

        with col1:
            for alarm in ordered_alarms[:len(ordered_alarms)//2]:
                pivot, total_count = create_pivot_table(df_alarm, alarm)
                alarm_pivots[alarm] = (pivot, total_count)
                st.markdown(f"### {alarm}")
                st.markdown(f"<small><i>till {alarm_file_time}</i></small>", unsafe_allow_html=True)
                st.markdown(f"**Total Count:** {int(total_count)}")
                st.dataframe(pivot)

        with col2:
            for alarm in ordered_alarms[len(ordered_alarms)//2:]:
                pivot, total_count = create_pivot_table(df_alarm, alarm)
                alarm_pivots[alarm] = (pivot, total_count)
                st.markdown(f"### {alarm}")
                st.markdown(f"<small><i>till {alarm_file_time}</i></small>", unsafe_allow_html=True)
                st.markdown(f"**Total Count:** {int(total_count)}")
                st.dataframe(pivot)
        
        # Excel download button
        excel_data = to_excel(alarm_pivots)
        st.download_button(
            label="Download All Pivot Tables as Excel",
            data=excel_data,
            file_name=f"RMS Alarm Report {alarm_file_time}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
