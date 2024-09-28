import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
import io

# Function to process the Excel file and generate summary tables
def process_excel(file):
    try:
        # Load the Excel file and use row 3 (index 2) as the header
        df = pd.read_excel(file, header=2)
        st.subheader("📄 Original Data")
        st.dataframe(df)
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

    # Step 0: Identify leased sites (sites starting with 'L' in column 'Site')
    if 'Site' in df.columns:
        leased_sites = df[df['Site'].str.startswith('L', na=False)]
        leased_site_names = leased_sites['Site'].tolist()  # List of leased site names for filtering
        st.write(f"✅ Identified {len(leased_site_names)} leased sites.")
    else:
        st.error(f"'Site' column not found in the data! Columns available: {df.columns.tolist()}")
        return None

    # Step 1: Identify BANJO and Non BANJO sites in the 'Site Alias' column
    if 'Site Alias' in df.columns:
        df['Site Type'] = df['Site Alias'].apply(lambda x: 'BANJO' if '(BANJO)' in str(x) else 'Non BANJO')
        # Extract client information (BL, GP, Robi) from Site Alias
        df['Client'] = df['Site Alias'].apply(
            lambda x: 'BL' if '(BL)' in str(x)
            else 'GP' if '(GP)' in str(x)
            else 'Robi' if '(ROBI)' in str(x)
            else 'BANJO'
        )
        st.write("✅ Added 'Site Type' and 'Client' columns.")
    else:
        st.error(f"'Site Alias' column not found in the data! Columns available: {df.columns.tolist()}")
        return None

    # Step 2: Filter relevant alarm categories in the 'Alarm Name' column
    alarm_categories = ['Battery Low', 'Mains Fail', 'DCDB-01 Primary Disconnect', 'PG Run']
    if 'Alarm Name' in df.columns:
        filtered_df = df[df['Alarm Name'].isin(alarm_categories)]
        st.write(f"✅ Filtered data to {len(filtered_df)} records with specified alarm categories.")
    else:
        st.error(f"'Alarm Name' column not found in the data! Columns available: {df.columns.tolist()}")
        return None

    # Step 3: Exclude leased sites for "DCDB-01 Primary Disconnect" alarm
    dcdb_df = filtered_df[filtered_df['Alarm Name'] == 'DCDB-01 Primary Disconnect']
    non_leased_dcdb_df = dcdb_df[~dcdb_df['Site'].isin(leased_site_names)]  # Exclude leased sites

    # Step 4: Create a combined DataFrame with all alarms, replacing the filtered "DCDB-01 Primary Disconnect"
    filtered_df = pd.concat([
        filtered_df[filtered_df['Alarm Name'] != 'DCDB-01 Primary Disconnect'],
        non_leased_dcdb_df
    ])

    # Display the filtered DataFrame
    st.subheader("🔍 Filtered Data")
    st.dataframe(filtered_df)

    # Step 5: Generate a client-wise table for tenant counts for each alarm
    if 'Tenant' in df.columns:
        client_table = df.groupby(['Alarm Name', 'Client', 'Tenant']).size().reset_index(name='Count')  # Group by Alarm, Client, and Tenant and count

        # Pivot the table to get separate columns for each Tenant under each Alarm and Client
        pivot_table = client_table.pivot_table(index=['Alarm Name', 'Client'], columns='Tenant', values='Count', fill_value=0)
        pivot_table.reset_index(inplace=True)
        pivot_table.columns.name = None

        st.subheader("📊 Client-wise Tenant Counts")
        st.dataframe(pivot_table)
    else:
        st.error(f"'Tenant' column not found in the data! Columns available: {df.columns.tolist()}")
        return None

    # Check for RIO (Cluster) and Subcenter (Zone)
    if 'Cluster' in df.columns and 'Zone' in df.columns:

        # Step 7: Group data by Alarm Name, Cluster, Zone, Client, and Tenant and count occurrences
        summary_table = filtered_df.groupby(['Alarm Name', 'Cluster', 'Zone', 'Tenant']).size().reset_index(name='Alarm Count')

        # Pivot the summary table to have Tenant-wise counts in separate columns
        summary_pivot = summary_table.pivot_table(index=['Alarm Name', 'Cluster', 'Zone'], columns='Tenant', values='Alarm Count', fill_value=0)
        summary_pivot.reset_index(inplace=True)
        summary_pivot.columns.name = None

        # Calculate Zone Total
        tenant_columns = [col for col in summary_pivot.columns if col not in ['Alarm Name', 'Cluster', 'Zone']]
        summary_pivot['Zone Total'] = summary_pivot[tenant_columns].sum(axis=1)

        # Calculate Overall Total per Alarm
        overall_total = summary_pivot['Zone Total'].sum()

        # Sort the summary_pivot by Cluster and Zone
        summary_pivot = summary_pivot.sort_values(by=['Alarm Name', 'Cluster', 'Zone'])

        st.subheader("📈 Summary Table")
        st.dataframe(summary_pivot)

        # Define color fill and styles for Excel
        rios = summary_pivot['Cluster'].unique()
        colors = ['FFFF99', 'FF9999', 'D5A6F2', '99CCFF', 'FFCC99']  # Example color codes
        color_map = {rio: PatternFill(start_color=color, end_color=color, fill_type="solid") for rio, color in zip(rios, colors)}

        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        total_row_fill = PatternFill(start_color="EAD1DC", end_color="EAD1DC", fill_type="solid")
        overall_row_fill = PatternFill(start_color="E0EBEB", end_color="E0EBEB", fill_type="solid")
        bold_font = Font(bold=True)
        border_style = Border(left=Side(border_style="thin"),
                              right=Side(border_style="thin"),
                              top=Side(border_style="thin"),
                              bottom=Side(border_style="thin"))
        center_alignment = Alignment(horizontal='center', vertical='center')

        # Create a new Workbook
        wb = Workbook()
        ws_original = wb.active
        ws_original.title = "Original Data"
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws_original.cell(row=r_idx, column=c_idx, value=value)

        # Define alarm categories again to ensure consistency
        alarm_categories = alarm_categories  # ['Battery Low', 'Mains Fail', 'DCDB-01 Primary Disconnect', 'PG Run']

        # Function to apply borders
        def apply_borders(ws):
            thin_border = Border(left=Side(border_style="thin"),
                                 right=Side(border_style="thin"),
                                 top=Side(border_style="thin"),
                                 bottom=Side(border_style="thin"))
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = thin_border

        # Add each alarm-specific table to separate sheets
        for alarm in alarm_categories:
            alarm_data = summary_pivot[summary_pivot['Alarm Name'] == alarm]

            if alarm_data.empty:
                st.warning(f"No data found for alarm: **{alarm}**")
                continue

            # Create a new sheet for the alarm
            ws = wb.create_sheet(title=alarm)

            # Write the Alarm Name and Total Alarm Count at the top
            ws.cell(row=1, column=1, value="Alarm Name:")
            ws.cell(row=1, column=2, value=alarm).font = bold_font
            ws.cell(row=2, column=1, value="Total Alarm Count:")
            total_alarm_count = alarm_data['Zone Total'].sum()
            ws.cell(row=2, column=2, value=total_alarm_count).font = bold_font

            # Write the table headers
            headers = ['Cluster', 'Zone'] + tenant_columns + ['Zone Total']
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_idx, value=header)
                cell.font = bold_font
                cell.fill = header_fill
                cell.alignment = center_alignment

            # Write the data rows
            for r_idx, row in enumerate(alarm_data.itertuples(index=False), start=5):
                ws.cell(row=r_idx, column=1, value=row.Cluster)
                ws.cell(row=r_idx, column=2, value=row.Zone)
                for i, tenant in enumerate(tenant_columns, start=3):
                    ws.cell(row=r_idx, column=i, value=row._asdict()[tenant])
                ws.cell(row=r_idx, column=len(tenant_columns) + 3, value=row._asdict()['Zone Total'])

            # Add Total row
            total_row_idx = len(alarm_data) + 5
            ws.cell(row=total_row_idx, column=1, value="Total").font = bold_font
            for i, tenant in enumerate(tenant_columns, start=3):
                total = alarm_data[tenant].sum()
                ws.cell(row=total_row_idx, column=i, value=total).font = bold_font
            ws.cell(row=total_row_idx, column=len(tenant_columns) + 3, value=alarm_data['Zone Total'].sum()).font = bold_font

            # Add Overall row
            overall_row_idx = total_row_idx + 1
            ws.cell(row=overall_row_idx, column=1, value="Overall").font = bold_font
            ws.cell(row=overall_row_idx, column=1).fill = overall_row_fill
            ws.cell(row=overall_row_idx, column=1).alignment = center_alignment

            # Calculate overall total
            overall_total = alarm_data['Zone Total'].sum()
            ws.cell(row=overall_row_idx, column=2, value=overall_total).font = bold_font
            ws.cell(row=overall_row_idx, column=2).fill = overall_row_fill
            ws.cell(row=overall_row_idx, column=2).alignment = center_alignment

            # Apply borders
            apply_borders(ws)

            # Adjust column widths
            for column_cells in ws.columns:
                length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                adjusted_width = length + 2
                column_letter = column_cells[0].column_letter
                ws.column_dimensions[column_letter].width = adjusted_width

        # Save the workbook to an in-memory bytes buffer
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return output, summary_pivot
    else:
        st.error(f"'Cluster' and/or 'Zone' column(s) not found in the data! Columns available: {df.columns.tolist()}")
        return None

# Streamlit App
def main():
    st.set_page_config(page_title="🛠️ Alarm Summary Report Generator", layout="wide")
    st.title("🛠️ Alarm Summary Report Generator")

    st.write("""
    Upload your Excel file to generate a formatted Alarm Summary Report.
    The report includes separate summary tables for each alarm category, grouped by clusters and zones with tenant-wise counts.
    """)

    uploaded_file = st.file_uploader("📂 Choose an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        with st.spinner('🔄 Processing...'):
            result = process_excel(uploaded_file)
            if result:
                report, summary_table = result
                st.success("✅ Report generated successfully!")

                # Display separate summary tables for each alarm category
                st.subheader("📑 Detailed Summary Tables by Alarm")

                alarm_categories = summary_table['Alarm Name'].unique()

                for alarm in alarm_categories:
                    alarm_data = summary_table[summary_table['Alarm Name'] == alarm]
                    if alarm_data.empty:
                        st.warning(f"No data found for alarm: **{alarm}**")
                        continue

                    with st.expander(f"🔍 {alarm} Summary"):
                        # Create a copy to display
                        display_table = alarm_data.copy()

                        # Rename columns for better readability
                        display_table = display_table.rename(columns={
                            'Cluster': 'Cluster',
                            'Zone': 'Zone',
                            'Zone Total': 'Zone Total'
                        })

                        # Ensure tenant columns are ordered
                        tenant_columns = [col for col in display_table.columns if col not in ['Alarm Name', 'Cluster', 'Zone', 'Zone Total']]
                        display_table = display_table[['Cluster', 'Zone'] + tenant_columns + ['Zone Total']]

                        # Display the table
                        st.table(display_table)

                # Provide a download button for the report
                st.download_button(
                    label="💾 Download Excel Report",
                    data=report,
                    file_name="Alarm_Summary_Report_Formatted_By_Alarm.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
