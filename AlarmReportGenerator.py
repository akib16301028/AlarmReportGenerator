import pandas as pd
import os
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment

# Function to process the uploaded file
def process_file(uploaded_file):
    # Load the Excel file and use row 3 (index 2) as the header
    df = pd.read_excel(uploaded_file, header=2)

    # Step 0: Identify leased sites (sites starting with 'L' in column B)
    if 'Site' in df.columns:
        leased_sites = df[df['Site'].str.startswith('L', na=False)]
        leased_site_names = leased_sites['Site'].tolist()  # List of leased site names for filtering
    else:
        st.error("'Site' column not found in the data!")
        return

    # Step 1: Identify BANJO and Non BANJO sites in the 'Site Alias' column (Column C)
    if 'Site Alias' in df.columns:
        df['Site Type'] = df['Site Alias'].apply(lambda x: 'BANJO' if '(BANJO)' in str(x) else 'Non BANJO')
        # Extract client information (BL, GP, Robi) from Site Alias
        df['Client'] = df['Site Alias'].apply(lambda x: 'BL' if '(BL)' in str(x)
        else 'GP' if '(GP)' in str(x)
        else 'Robi' if '(ROBI)' in str(x)
        else 'BANJO')
    else:
        st.error("'Site Alias' column not found in the data!")
        return

    # Step 2: Filter relevant alarm categories in the 'Alarm Name' column
    alarm_categories = ['Battery Low', 'Mains Fail', 'DCDB-01 Primary Disconnect', 'PG Run']
    if 'Alarm Name' in df.columns:
        filtered_df = df[df['Alarm Name'].isin(alarm_categories)]
    else:
        st.error("'Alarm Name' column not found in the data!")
        return

    # Step 3: Exclude leased sites for "DCDB-01 Primary Disconnect" alarm
    dcdb_df = filtered_df[filtered_df['Alarm Name'] == 'DCDB-01 Primary Disconnect']
    non_leased_dcdb_df = dcdb_df[~dcdb_df['Site'].isin(leased_site_names)]  # Exclude leased sites

    # Step 4: Create a combined DataFrame with all alarms, replacing the filtered "DCDB-01 Primary Disconnect"
    filtered_df = pd.concat([filtered_df[filtered_df['Alarm Name'] != 'DCDB-01 Primary Disconnect'], non_leased_dcdb_df])

    # Step 5: Generate a client-wise table for tenant counts for each alarm
    if 'Tenant' in df.columns:
        client_table = df.groupby(['Alarm Name', 'Client', 'Tenant']).size().reset_index(name='Count')  # Group by Alarm, Client, and Tenant and count

        # Pivot the table to get separate columns for each Tenant under each Alarm and Client
        pivot_table = client_table.pivot_table(index=['Alarm Name', 'Client'], columns='Tenant', values='Count', fill_value=0)
        pivot_table.reset_index(inplace=True)
        pivot_table.columns.name = None
    else:
        st.error("'Tenant' column not found in the data!")
        return

    # Check for RIO (Cluster) and Subcenter (Zone)
    if 'Cluster' in df.columns and 'Zone' in df.columns:
        # Step 7: Group data by Alarm Name, Client (BANJO, GP, BL, Robi), RIO, Subcenter, and count occurrences
        summary_table = filtered_df.groupby(['Alarm Name', 'Cluster', 'Zone', 'Client']).size().reset_index(name='Alarm Count')

        # Renaming for better readability
        summary_table = summary_table.rename(columns={'Cluster': 'RIO', 'Zone': 'Subcenter'})

        # Sort the summary_table by RIO and Alarm Count (descending)
        summary_table = summary_table.sort_values(by=['RIO', 'Alarm Count'], ascending=[True, False])

        # Prepare output
        output_tables = []

        for alarm in alarm_categories:
            alarm_data = summary_table[summary_table['Alarm Name'] == alarm]

            # Pivot the table to get separate columns for each client (BANJO, GP, BL, Robi)
            pivot_table = alarm_data.pivot_table(index=['RIO', 'Subcenter'], columns='Client', values='Alarm Count', fill_value=0)
            pivot_table.reset_index(inplace=True)
            pivot_table.columns.name = None

            # Sort the pivot table data first by RIO, then by counts in descending order
            pivot_table = pivot_table.sort_values(by=['RIO', 'Subcenter'], ascending=[True, True])

            # Calculate total counts for each RIO
            pivot_table['Total Count'] = pivot_table[[col for col in pivot_table.columns if col not in ['RIO', 'Subcenter']]].sum(axis=1)

            # Prepare the display format
            output_table = pd.DataFrame({
                'RIO': pivot_table['RIO'],
                'Subcenter': pivot_table['Subcenter'],
                'Client 1': pivot_table.get('BANJO', 0),  # Default to 0 if 'BANJO' not present
                'Client 2': pivot_table.get('GP', 0),      # Default to 0 if 'GP' not present
                'Client 3': pivot_table.get('BL', 0),      # Default to 0 if 'BL' not present
                'Client 4': pivot_table.get('Robi', 0),    # Default to 0 if 'Robi' not present
                'Total Count': pivot_table['Total Count']
            })

            output_tables.append((alarm, output_table))

        # Display the tables
        for alarm_name, table in output_tables:
            st.write(f"### {alarm_name}")
            st.write("#### Total Alarm Count")
            st.write(table)

            # Add merged RIO cells
            # Note: We need to save the result into an Excel file to perform merging
            output_file = os.path.join(os.path.expanduser("~"), "Desktop", "Alarm_Summary_Report.xlsx")
            wb = Workbook()

            ws = wb.create_sheet(title=alarm_name)

            for r_idx, row in enumerate(dataframe_to_rows(table, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

            # Merging RIO cells for identical RIO values
            current_rio = None
            start_row = None
            for r_idx in range(2, len(table) + 2):  # Start from row 2 to skip header
                cell_value = ws.cell(row=r_idx, column=1).value  # RIO column index
                if cell_value == current_rio:
                    continue
                if current_rio is not None and start_row is not None:
                    ws.merge_cells(start_row=start_row, start_column=1, end_row=r_idx - 1, end_column=1)
                current_rio = cell_value
                start_row = r_idx
            if start_row is not None:
                ws.merge_cells(start_row=start_row, start_column=1, end_row=len(table) + 1, end_column=1)

            wb.save(output_file)
            st.success(f'Report saved successfully as {output_file}')

    else:
        st.error("'Cluster' and/or 'Zone' column(s) not found in the data!")

# Streamlit interface
st.title("Alarm Summary Report Generator")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
if uploaded_file:
    process_file(uploaded_file)
