import streamlit as st
import pandas as pd
import re
from io import BytesIO
import pytz
from datetime import datetime

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

# Function to create offline report pivot with day-wise duration
def create_offline_daywise_pivot(df):
    # Ensure 'Last Online Time' is in datetime format
    df['Last Online Time'] = pd.to_datetime(df['Last Online Time'], errors='coerce')
    
    # Dhaka timezone
    dhaka_tz = pytz.timezone('Asia/Dhaka')
    current_time = datetime.now(dhaka_tz)
    
    # Calculate the duration (difference) in days
    df['Duration (Days)'] = (current_time - df['Last Online Time']).dt.total_seconds() / (24 * 3600)
    
    # Create a pivot table to categorize by 'Duration (Days)'
    pivot = df.groupby(['Cluster', 'Zone']).agg({
        'Duration (Days)': 'mean',  # Calculate average duration per cluster and zone
        'Site Alias': 'nunique'
    }).reset_index()

    pivot = pivot.rename(columns={'Site Alias': 'Total'})
    
    # Add a total row at the end
    total_row = pivot[['Duration (Days)', 'Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    total_offline_count = int(pivot['Total'].iloc[-1])
    
    # Remove repeated Cluster names for better readability
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_offline_count

# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            valid_sheet_name = re.sub(r'[\\/*?:[\]]', '_', sheet_name)[:31]
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    return output.getvalue()

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
    
    styler = styler.applymap(highlight_zero)
    
    # Handle total row: set all cells to empty except 'Cluster' and 'Zone' if needed
    if total_row_mask.any():
        styler = styler.apply(
            lambda x: ['background-color: #f0f0f0; color: black' if total_row_mask.loc[x.name] else '' for _ in x],
            axis=1
        )
        # Optionally, you can set the 'Cluster' and 'Zone' cells to have a different style
        styler = styler.applymap(
            lambda x: f'background-color: {cell_bg_color}; color: {font_color}',
            subset=['Cluster', 'Zone']
        )
    
    # Optional: Remove borders for a cleaner look
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

        # Get unique clusters for filtering
        offline_clusters = sorted(offline_df['Cluster'].dropna().unique().tolist())
        offline_clusters.insert(0, "All")  # Add 'All' option
        selected_offline_cluster = st.sidebar.selectbox(
            "Select Cluster",
            options=offline_clusters,
            index=0
        )

        # Process the Offline Report
        pivot_offline_daywise, total_offline_count = create_offline_daywise_pivot(offline_df)

        # Apply Offline Cluster Filters
        if selected_offline_cluster != "All":
            filtered_pivot_offline = pivot_offline_daywise[
                (pivot_offline_daywise['Cluster'] == selected_offline_cluster) | (pivot_offline_daywise['Cluster'] == 'Total')
            ]
        else:
            filtered_pivot_offline = pivot_offline_daywise.copy()

        # Display the Offline Report
        st.markdown("### Offline Report - Day Wise Duration")
        st.markdown(f"**Total Sites Offline**: {total_offline_count}")
        st.dataframe(filtered_pivot_offline)

        # Download Option for Offline Report
        offline_report_excel_data = to_excel({'Offline Report Day-Wise': filtered_pivot_offline})
        st.download_button(
            label="Download Offline Day-Wise Duration Report",
            data=offline_report_excel_data,
            file_name="Offline_DayWise_Duration_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload both the Alarm and Offline report files.")
