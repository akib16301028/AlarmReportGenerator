import pandas as pd

# Load the Excel file
file_path = 'your_file_path.xlsx'  # Update with your file path
df = pd.read_excel(file_path, header=3)

# Define alarms and tenants
alarms = df['Alarm Name'].unique()
tenants = df['Tenant'].unique()

# Create an Excel writer
output_file = 'Alarm_Summary_Report.xlsx'
with pd.ExcelWriter(output_file) as writer:
    for alarm in alarms:
        alarm_data = df[df['Alarm Name'] == alarm]

        # Group by Cluster and Zone, counting each tenant
        summary_table = alarm_data.groupby(['Cluster', 'Zone', 'Tenant']).size().unstack(fill_value=0)

        # Add total count columns for each zone
        summary_table['Zone Total'] = summary_table.sum(axis=1)
        summary_table.loc['Total'] = summary_table.sum()

        # Add a total alarm count for the specific alarm
        total_count = len(alarm_data)

        # Create a formatted summary table
        formatted_summary = pd.DataFrame({
            'Total Alarm Count': [total_count],
            'Zone Wise Count': ['']  # Placeholder for clear separation
        })

        for cluster in summary_table.index:
            formatted_summary.loc[cluster] = [''] + summary_table.loc[cluster].tolist()

        # Write to the Excel file
        formatted_summary.to_excel(writer, sheet_name=alarm, index=True)

# Inform the user of successful completion
print(f"Summary report generated successfully in {output_file}.")
