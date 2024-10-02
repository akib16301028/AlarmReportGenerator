import streamlit as st
import pandas as pd
import re
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
    
    # Categorize Duration Slot (Hours)
    alarm_df['Duration Category'] = alarm_df['Duration Slot (Hours)'].apply(categorize_duration)
    
    # Create pivot tables for each Duration Category
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
    
    # Pivot for Duration Categories
    pivot_duration = pd.pivot_table(
        alarm_df,
        index=['Cluster', 'Zone'],
        columns='Duration Category',
        values='Site Alias',
        aggfunc='count',
        fill_value=0
    ).reset_index()
    
    # Merge the two pivot tables
    pivot = pd.merge(pivot_total, pivot_duration, on=['Cluster', 'Zone'], how='left')
    
    # Ensure all Duration Categories are present
    for cat in ['0+', '2+', '4+', '8+']:
        if cat not in pivot.columns:
            pivot[cat] = 0
    
    # Reorder columns: Original client columns + 'Total' + Duration Categories
    pivot = pivot[['Cluster', 'Zone'] + client_columns + ['Total', '0+', '2+', '4+', '8+']]
    
    # Add Total row
    numeric_cols = pivot.select_dtypes(include=['number']).columns
    total_row = pivot[numeric_cols].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    # Replace numeric columns in total_row with empty strings
    total_row[numeric_cols] = total_row[numeric_cols].replace(0, "").astype(str)
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    total_alarm_count = pivot['Total'].iloc[-1]
    
    # Remove repeated Cluster names for better readability
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
    
    # Create new columns for each duration category
    df['Less than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'Less than 24 hours' in x else 0)
    df['More than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 24 hours' in x and '72' not in x else 0)
    df['More than 48 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 48 hours' in x else 0)
    df['More than 72 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 72 hours' in x else 0)
    
    pivot = df.groupby(['Cluster', 'Zone']).agg({
        'Less than 24 hours': 'sum',
        'More than 24 hours': 'sum',
        'More than 48 hours': 'sum',  # Include the new column here
        'More than 72 hours': 'sum',
        'Site Alias': 'nunique'
    }).reset_index()

    pivot = pivot.rename(columns={'Site Alias': 'Total'})
    
    total_row = pivot[['Less than 24 hours', 'More than 24 hours', 'More than 48 hours', 'More than 72 hours', 'Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    # Replace numeric columns in total_row with empty strings
    numeric_cols = ['Less than 24 hours', 'More than 24 hours', 'More than 48 hours', 'More than 72 hours', 'Total']
    total_row[numeric_cols] = total_row[numeric_cols].replace(0, "").astype(str)
    
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

    # Format 'Last Online Time' to exclude microseconds
    df['Last Online Time'] = df['Last Online Time'].dt.strftime('%Y-%m-%d %H:%M:%S')

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

# Function to create site-wise log table
def create_site_wise_log(df, selected_alarm):
    if selected_alarm == "All":
        filtered_df = df.copy()
    else:
        filtered_df = df[df['Alarm Name'] == selected_alarm].copy()
    filtered_df = filtered_df[['Site Alias', 'Cluster', 'Zone', 'Alarm Name', 'Alarm Time', 'Duration']]
    filtered_df = filtered_df.sort_values(by='Alarm Time', ascending=False)
    return filtered_df

# Function to style DataFrame: fill cells with #f0f0f0 if value is 0 or empty and handle total row
def style_dataframe(df, duration_cols, is_dark_mode):
    # Create a copy for styling
    df_style = df.copy()
    
    # Identify the total row based on 'Cluster' column
    total_row_mask = df_style['Cluster'] == 'Total'
    
    # Replace 0 with empty strings in duration columns
    df_style[duration_cols] = df_style[duration_cols].replace(0, "")
    
    # Define background colors
    cell_bg_color = '#f0f0f0'
    font_color = 'black' if not is_dark_mode else 'black'
    
    # Create a Styler object
    styler = df_style.style
    
    # Apply background color to cells with 0 or empty values
    def highlight_zero(val):
        if val == 0 or val == "":
            return f'background-color: {cell_bg_color}; color: {font_color}'
        return ''
    
    styler = styler.applymap(highlight_zero, subset=duration_cols)
    
    # Make the total row bold
    if total_row_mask.any():
        styler = styler.apply(lambda x: ['font-weight: bold' if total_row_mask.iloc[i] else '' for i in range(len(x))], axis=1)
    
    return styler

# Streamlit application layout
st.title("Alarm Report Application")
st.sidebar.title("Upload Data")

# File upload for current alarms
uploaded_file_alarms = st.sidebar.file_uploader("Upload Current Alarms CSV", type=["csv"])
if uploaded_file_alarms is not None:
    df_alarms = pd.read_csv(uploaded_file_alarms)
    df_alarms['Alarm Time'] = pd.to_datetime(df_alarms['Alarm Time'])
    st.sidebar.success("Alarms data uploaded successfully.")

# File upload for offline reports
uploaded_file_offline = st.sidebar.file_uploader("Upload Offline Reports CSV", type=["csv"])
if uploaded_file_offline is not None:
    df_offline = pd.read_csv(uploaded_file_offline)
    st.sidebar.success("Offline reports data uploaded successfully.")

# Select an alarm to filter
selected_alarm = st.sidebar.selectbox("Select Alarm", options=['All'] + df_alarms['Alarm Name'].unique().tolist() if 'df_alarms' in locals() else [])
if 'df_alarms' in locals():
    if selected_alarm:
        # Create the main pivot table for the selected alarm
        pivot_table, total_alarm_count = create_pivot_table(df_alarms, selected_alarm)
        
        # Display the pivot table
        st.subheader(f"Pivot Table for '{selected_alarm}' Alarm")
        st.dataframe(style_dataframe(pivot_table, ['Total', '0+', '2+', '4+', '8+'], is_dark_mode=False), use_container_width=True)
        
        # Create and display site-wise log
        site_wise_log = create_site_wise_log(df_alarms, selected_alarm)
        st.subheader("Site-wise Log Table")
        st.dataframe(site_wise_log, use_container_width=True)

# Display offline report if uploaded
if 'df_offline' in locals():
    current_time = pd.to_datetime("now")
    df_offline = calculate_time_offline(df_offline, current_time)
    
    # Displaying the offline report
    st.subheader("Offline Report")
    st.dataframe(df_offline, use_container_width=True)
    
    # Create the offline pivot table
    offline_pivot, total_offline_count = create_offline_pivot(df_offline)
    
    # Display the offline pivot table
    st.subheader("Offline Pivot Table")
    st.dataframe(style_dataframe(offline_pivot, ['Less than 24 hours', 'More than 24 hours', 'More than 48 hours', 'More than 72 hours', 'Total'], is_dark_mode=False), use_container_width=True)

# Button to download the report as Excel
if st.button("Download Report"):
    if 'df_alarms' in locals() and 'df_offline' in locals():
        report_dict = {
            "Alarm Report": pivot_table,
            "Site-wise Log": site_wise_log,
            "Offline Report": df_offline,
            "Offline Pivot": offline_pivot
        }
        excel_data = to_excel(report_dict)
        st.download_button("Download Excel", data=excel_data, file_name="alarm_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Please upload both Current Alarms and Offline Reports to download the report.")
