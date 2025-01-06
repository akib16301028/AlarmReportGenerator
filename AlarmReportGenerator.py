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
    total_row[numeric_cols] = total_row[numeric_cols].replace(0, "0").astype(str)
    
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
    total_row[numeric_cols] = total_row[numeric_cols].replace(0, "0").astype(str)
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    total_offline_count = int(pivot['Total'].iloc[-1])
    
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_offline_count

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
    filtered_df = filtered_df[['Site Alias', 'Cluster', 'Zone', 'Alarm Name', 'Alarm Time','Duration']]
    filtered_df = filtered_df.sort_values(by='Alarm Time', ascending=False)
    return filtered_df

# Function to style the dataframe
def style_dataframe(df, duration_cols, is_dark_mode):
    df_style = df.copy()
    total_row_mask = df_style['Cluster'] == 'Total'
    cell_bg_color = '#f0f0f0'
    font_color = 'black' if not is_dark_mode else 'black'
    styler = df_style.style

    def highlight_zero(val):
        if val != 0:
            return f'background-color: {cell_bg_color}; color: {font_color}'
        return ''
    
    styler = styler.applymap(highlight_zero)

    if total_row_mask.any():
        styler = styler.apply(
            lambda x: ['background-color: #f0f0f0; color: black' if total_row_mask.loc[x.name] else '' for _ in x],
            axis=1
        )
        styler = styler.applymap(
            lambda x: f'background-color: {cell_bg_color}; color: {font_color}',
            subset=['Cluster', 'Zone']
        )
    
    styler.set_table_styles(
        [{
            'selector': 'th',
            'props': [('border', '1px solid black')]
        },
        {
            'selector': 'td',
            'props': [('border', '1px solid black')]
        }]
    )
    
    return styler

# Function to determine if the current theme is dark
def is_dark_mode():
    if 'theme' in st.session_state:
        theme = st.session_state['theme']
    else:
        theme = 'light'

    return theme.lower() == 'dark'

# Streamlit app
st.title("StatusMatrix@STL")

# File Uploads
uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        # Read Excel files starting from the third row (header=2)
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)

        # Initialize Sidebar Filters
        st.sidebar.header("Filters")
        offline_clusters = sorted(offline_df['Cluster'].dropna().unique().tolist())
        offline_clusters.insert(0, "All")  
        selected_offline_cluster = st.sidebar.selectbox(
            "Select Cluster",
            options=offline_clusters,
            index=0
        )

        st.sidebar.subheader("Current Alarms Filters")
        st.sidebar.text("[select alarm first]")

        alarm_names = sorted(alarm_df['Alarm Name'].dropna().unique().tolist())
        alarm_names.insert(0, "All")  
        selected_alarm = st.sidebar.selectbox(
            "Select Alarm to Filter",
            options=alarm_names,
            index=0
        )
        
        # Alarm Filtering and Processing
        pivot_alarm, total_alarm_count = create_pivot_table(alarm_df, selected_alarm)

        # Offline Filtering and Processing
        filtered_offline_df = offline_df.copy()
        if selected_offline_cluster != 'All':
            filtered_offline_df = filtered_offline_df[filtered_offline_df['Cluster'] == selected_offline_cluster]
        offline_pivot, total_offline_count = create_offline_pivot(filtered_offline_df)

        # Display Pivot Tables
        st.subheader(f"Alarm Counts by Cluster and Zone - {selected_alarm}")
        st.dataframe(pivot_alarm)

        st.subheader(f"Offline Summary")
        st.dataframe(offline_pivot)

        # Generate Downloadable Reports
        dfs_dict = {
            f"Alarm Report ({selected_alarm})": pivot_alarm,
            f"Offline Report": offline_pivot
        }
        st.download_button(
            label="Download Reports",
            data=to_excel(dfs_dict),
            file_name="Alarm_Offline_Reports.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.warning("Please upload the required Excel files.")

