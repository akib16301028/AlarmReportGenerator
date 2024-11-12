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

# Function to convert DataFrame to HTML table with custom styling
def dataframe_to_styled_html(df, duration_cols):
    # Define CSS style for the table
    style = """
    <style>
        table {
            border-collapse: collapse;
            width: 100%;
            font-size: 12px;
            text-align: center;
        }
        th, td {
            border: 1px solid #dddddd;
            padding: 4px;
        }
        th {
            background-color: #f2f2f2;
        }
        .highlight {
            background-color: #ADD8E6;
        }
        .total-row {
            background-color: #D3D3D3;
            font-weight: bold;
        }
    </style>
    """
    # Apply styles based on columns
    df_styled = df.copy()
    df_styled[duration_cols] = df_styled[duration_cols].replace(0, "")
    df_styled_html = df_styled.to_html(index=False, classes="styled-table")
    return f"{style}{df_styled_html}"

# Streamlit app
st.title("StatusMatrix")

# File Uploads
uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        # Read Excel files starting from the third row (header=2)
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)

        # Example filtered dataframe display
        pivot_offline, total_offline_count = create_pivot_table(alarm_df, "DCDB-01 Primary Disconnect")

        # Get duration columns
        duration_cols = ['0+', '2+', '4+', '8+']

        # Convert the pivot table DataFrame to HTML
        html_table = dataframe_to_styled_html(pivot_offline, duration_cols)

        # Display the styled HTML table in Streamlit
        st.markdown(html_table, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
