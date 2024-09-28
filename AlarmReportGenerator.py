import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Function to extract client name from Site Alias
def extract_client(site_alias):
    match = re.search(r'\((.*?)\)', site_alias)
    return match.group(1) if match else None

# Streamlit app
st.title("Alarm Data Pivot Table Generator")

# Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file is not None:
    # Read the uploaded Excel file
    df = pd.read_excel(uploaded_file, header=2)

    # Add a new column for Client extracted from Site Alias
    df['Client'] = df['Site Alias'].apply(extract_client)

    # Filter out leased sites for the specific alarm
    filtered_df = df[~df['RMS Station'].str.startswith('L') | (df['Alarm Name'] != 'DCDB-01 Primary Disconnect')]

    # Create a pivot table for the alarm 'DCDB-01 Primary Disconnect'
    pivot_table = pd.pivot_table(
        filtered_df[filtered_df['Alarm Name'] == 'DCDB-01 Primary Disconnect'],
        index=['Cluster', 'Zone', 'Client'],
        values='Site Alias',  # This is just a placeholder, will be counted
        aggfunc='count',
        fill_value=0
    ).reset_index()

    # Add a total count for each Client and overall
    pivot_table['Total'] = pivot_table.iloc[:, 3:].sum(axis=1)  # Sum all client counts
    total_counts = pivot_table[['Client', 'Total']].groupby('Client').sum().reset_index()

    # Append a row for the total
    total_row = pd.DataFrame({'Client': ['Total'], 'Total': [total_counts['Total'].sum()]})
    total_counts = pd.concat([total_counts, total_row], ignore_index=True)

    # Display pivot table
    st.subheader("Pivot Table")
    st.dataframe(pivot_table)

    # Display total counts
    st.subheader("Total Counts")
    st.dataframe(total_counts)

    # Function to convert DataFrame to Excel for download
    def to_excel(df1, df2):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Pivot Table', index=False)
            df2.to_excel(writer, sheet_name='Total Counts', index=False)
        return output.getvalue()

    # Create download button
    excel_data = to_excel(pivot_table, total_counts)
    st.download_button(
        label="Download Output Excel File",
        data=excel_data,
        file_name="pivot_table_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
