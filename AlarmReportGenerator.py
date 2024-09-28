import pandas as pd
import streamlit as st
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
import tempfile

# Function to process the uploaded file
def process_file(file):
    # Load the Excel file and use row 3 (index 2) as the header
    df = pd.read_excel(file, header=2)

    # Step 0: Identify leased sites (sites starting with 'L' in column B)
    if 'Site' in df.columns:
        leased_sites = df[df['Site'].str.startswith('L', na=False)]
        leased_site_names = leased_sites['Site'].tolist()
    else:
        raise ValueError("'Site' column not found in the data!")

    # Step 1: Identify BANJO and Non BANJO sites in the 'Site Alias' column (Column C)
    if 'Site Alias' in df.columns:
        df['Site Type'] = df['Site Alias'].apply(lambda x: 'BANJO' if '(BANJO)' in str(x) else 'Non BANJO')
        df['Client'] = df['Site Alias'].apply(lambda x: 'BL' if '(BL)' in str(x)
            else 'GP' if '(GP)' in str(x)
            else 'Robi' if '(ROBI)' in str(x)
            else 'BANJO')
    else:
        raise ValueError("'Site Alias' column not found in the data!")

    # Step 2: Filter relevant alarm categories in the 'Alarm Name' column
    alarm_categories = ['Battery Low', 'Mains Fail', 'DCDB-01 Primary Disconnect', 'PG Run']
    if 'Alarm Name' in df.columns:
        filtered_df = df[df['Alarm Name'].isin(alarm_categories)]
    else:
        raise ValueError("'Alarm Name' column not found in the data!")

    # Step 3: Exclude leased sites for "DCDB-01 Primary Disconnect" alarm
    dcdb_df = filtered_df[filtered_df['Alarm Name'] == 'DCDB-01 Primary Disconnect']
    non_leased_dcdb_df = dcdb_df[~dcdb_df['Site'].isin(leased_site_names)]
    filtered_df = pd.concat([filtered_df[filtered_df['Alarm Name'] != 'DCDB-01 Primary Disconnect'], non_leased_dcdb_df])

    # Step 5: Generate a client-wise table for tenant counts for each alarm
    if 'Tenant' in df.columns:
        client_table = df.groupby(['Alarm Name', 'Client', 'Tenant']).size().reset_index(name='Count')
        pivot_table = client_table.pivot_table(index=['Alarm Name', 'Client'], columns='Tenant', values='Count', fill_value=0)
        pivot_table.reset_index(inplace=True)
        pivot_table.columns.name = None
    else:
        raise ValueError("'Tenant' column not found in the data!")

    # Add Total Count
    pivot_table['Total'] = pivot_table.sum(axis=1)

    # Display the pivot table on the webpage
    st.subheader("Pivot Table")
    st.dataframe(pivot_table)

    return filtered_df, pivot_table

# Main Streamlit app
st.title("Alarm Summary Report Generator")

# File uploader for Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    filtered_df, pivot_table = process_file(uploaded_file)

    # Define the path for saving the output file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        output_file = temp_file.name

        # Create a new Workbook
        wb = Workbook()
        
        # Save the original DataFrame to a new sheet
        ws_original = wb.active
        ws_original.title = "Original Data"
        for r_idx, row in enumerate(dataframe_to_rows(filtered_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws_original.cell(row=r_idx, column=c_idx, value=value)

        # Adding the pivot table to the Excel file
        ws_pivot = wb.create_sheet(title="Pivot Table")
        for r_idx, row in enumerate(dataframe_to_rows(pivot_table, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws_pivot.cell(row=r_idx, column=c_idx, value=value)

        # Save the workbook
        wb.save(output_file)

    # Provide a download button for the generated report
    with open(output_file, "rb") as f:
        st.download_button("Download Excel Report", f, file_name="Alarm_Summary_Report_Formatted_By_Alarm.xlsx")

    st.success("Report generated successfully!")

else:
    st.info("Please upload an Excel file to generate the report.")
