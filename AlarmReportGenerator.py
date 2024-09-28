import pandas as pd
import streamlit as st
from openpyxl import Workbook
import io

# Function to load the Excel file
def load_data(file):
    """Load data from the uploaded Excel file, starting from the third row."""
    df = pd.read_excel(file, header=2)  # Adjust header row as needed
    return df

# Function to generate the summary tables for alarms
def generate_alarm_summaries(df):
    """Generate summary tables for each alarm."""
    alarm_categories = ['Battery Low', 'Mains Fail', 'DCDB-01 Primary Disconnect', 'PG Run']
    summaries = {}

    # Identify leased sites
    leased_sites = df[df['Site'].str.startswith('L', na=False)]['Site'].tolist()

    # Filter relevant alarms
    filtered_df = df[df['Alarm Name'].isin(alarm_categories)]

    # Exclude leased sites for "DCDB-01 Primary Disconnect"
    non_leased_dcdb_df = filtered_df[filtered_df['Alarm Name'] == 'DCDB-01 Primary Disconnect']
    non_leased_dcdb_df = non_leased_dcdb_df[~non_leased_dcdb_df['Site'].isin(leased_sites)]
    filtered_df = pd.concat([filtered_df[filtered_df['Alarm Name'] != 'DCDB-01 Primary Disconnect'], non_leased_dcdb_df])

    # Grouping data
    for alarm in alarm_categories:
        alarm_data = filtered_df[filtered_df['Alarm Name'] == alarm]
        if not alarm_data.empty:
            summary = (
                alarm_data.groupby(['Cluster', 'Zone', 'Tenant'])
                .size()
                .reset_index(name='Count')
            )

            # Create a pivot table for better visualization
            pivot_table = summary.pivot_table(index=['Cluster', 'Zone'], columns='Tenant', values='Count', fill_value=0)
            summaries[alarm] = pivot_table.reset_index()
    
    return summaries

# Streamlit interface
st.title("Alarm Report Analysis")

# Upload Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    st.subheader("Data Preview")
    st.dataframe(df)  # Display the loaded data

    # Generate alarm summaries
    alarm_summaries = generate_alarm_summaries(df)

    for alarm, summary in alarm_summaries.items():
        st.subheader(f"Summary for {alarm}")
        st.dataframe(summary)  # Display the summary table

    # Export the summaries to Excel
    if st.button("Download Summary as Excel"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for alarm, summary in alarm_summaries.items():
                summary.to_excel(writer, sheet_name=alarm, index=False)
            writer.save()
        output.seek(0)

        st.download_button(label="Download Excel", data=output, file_name="Alarm_Summaries.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
