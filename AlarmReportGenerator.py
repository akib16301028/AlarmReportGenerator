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
    
    # If the alarm is 'DCDB-01 Primary Disconnect', exclude leased sites
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
    
    # Flatten the columns
    pivot = pivot.reset_index()
    
    # Calculate Total per row
    client_columns = [col for col in pivot.columns if col not in ['Cluster', 'Zone']]
    pivot['Total'] = pivot[client_columns].sum(axis=1)
    
    # Calculate Total per client and overall total
    total_row = pivot[client_columns + ['Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    # Append the total row
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    # Calculate Total Alarm Count
    total_alarm_count = pivot['Total'].iloc[-1]
    
    return pivot, total_alarm_count

# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for alarm, (df, _) in dfs_dict.items():
            # Replace any characters in sheet name that are invalid in Excel
            sheet_name = re.sub(r'[\\/*?:[\]]', '_', alarm)[:31]  # Excel sheet name limit is 31 chars
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

# Streamlit app
st.title("Alarm Data Pivot Table Generator")

# Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file is not None:
    try:
        # Read the uploaded Excel file, assuming headers start from row 3 (0-indexed)
        df = pd.read_excel(uploaded_file, header=2)
        
        # Check if required columns exist
        required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name']
        if not all(col in df.columns for col in required_columns):
            st.error(f"The uploaded file is missing one of the required columns: {required_columns}")
        else:
            # Add a new column for Client extracted from Site Alias
            df['Client'] = df['Site Alias'].apply(extract_client)
            
            # Drop rows where Client extraction failed
            df = df.dropna(subset=['Client'])
            
            # Get unique alarms
            unique_alarms = df['Alarm Name'].unique()
            
            if len(unique_alarms) == 0:
                st.warning("No alarms found in the uploaded data.")
            else:
                # Dictionary to store pivot tables and total counts
                pivot_tables = {}
                
                for alarm in unique_alarms:
                    pivot, total_count = create_pivot_table(df, alarm)
                    pivot_tables[alarm] = (pivot, total_count)
                    
                    # Display headers and total counts
                    st.markdown(f"### Alarm Name: {alarm}")
                    st.markdown(f"**Total Alarm Count:** {int(total_count)}")
                    
                    # Display pivot table
                    st.dataframe(pivot)
                    st.markdown("---")  # Separator between tables
                
                # Function to convert multiple DataFrames to Excel with formatted headers
                def format_excel(dfs_dict):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for alarm, (df, total_count) in dfs_dict.items():
                            # Replace invalid sheet name characters and limit to 31 chars
                            sheet_name = re.sub(r'[\\/*?:[\]]', '_', alarm)[:31]
                            # Write Alarm Name and Total Alarm Count above the table
                            workbook = writer.book
                            worksheet = workbook.create_sheet(title=sheet_name)
                            
                            # Write headers
                            worksheet.cell(row=1, column=1, value=f"Alarm Name: {alarm}")
                            worksheet.cell(row=2, column=1, value=f"Total Alarm Count: {int(total_count)}")
                            
                            # Write the DataFrame starting from row 4
                            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=4):
                                for c_idx, value in enumerate(row, start=1):
                                    worksheet.cell(row=r_idx, column=c_idx, value=value)
                    
                    return output.getvalue()
                
                # Create download button
                excel_data = to_excel(pivot_tables)
                st.download_button(
                    label="Download All Pivot Tables as Excel",
                    data=excel_data,
                    file_name="pivot_tables_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
