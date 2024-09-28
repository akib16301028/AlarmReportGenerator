# Streamlit app
st.title("Alarm and Offline Data Pivot Table Generator")

# Upload Excel files for both Alarm Report and Offline Report
uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

# Define priority for alarms (you can customize this list as needed)
alarm_priority = {
    'DCDB-01 Primary Disconnect': 1,
    'Mains Fail': 2,  # replace with actual alarm names and priority
    'Battery Low': 3,
     'PG Run': 4,
'MDB Fault': 5,
'Door Open': 6
    # Add other alarms and their priorities
}

# Process both reports only after both files are uploaded
if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        # Read the Alarm Report file, assuming headers start from row 3 (0-indexed)
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        
        # Read the Offline Report file, assuming headers start from row 3 (0-indexed)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)
        
        # Extract date and time from the uploaded file names
        formatted_alarm_time = extract_timestamp(uploaded_alarm_file.name)
        formatted_offline_time = extract_timestamp(uploaded_offline_file.name)
        
        alarm_download_file_name = f"RMS Alarm Report {formatted_alarm_time}.xlsx"
        
        # Process the Offline Report first
        pivot_offline, total_offline_count = create_offline_pivot(offline_df)
        
        # Display header for Offline Report
        st.markdown("### Offline Report")
        st.markdown(f"<small><i>till {formatted_offline_time}</i></small>", unsafe_allow_html=True)
        st.markdown(f"**Total Offline Count:** {total_offline_count}")
        
        # Display the pivot table for Offline Report
        st.dataframe(pivot_offline)

        # Calculate days from Last Online Time
        days_offline_df = calculate_days_from_last_online(offline_df)
        
        # Create a summary table based on days offline
        summary_dict = {}
        for index, row in days_offline_df.iterrows():
            days = row['Days Offline']
            # Group days -1, 0, and 1 under "1 or Less than 1 day"
            if days in [-1, 0, 1]:
                key = "1 or Less than 1 day"
            else:
                key = days
            if key not in summary_dict:
                summary_dict[key] = []
            summary_dict[key].append(row)

        # Display the summary table
        for days, sites in summary_dict.items():
            st.markdown(f"### {days} Days")
            st.markdown("Site Name (Site Alias) Cluster Zone Last Online Time")
            for site in sites:
                st.markdown(f"{site['Site Alias']} {site['Cluster']} {site['Zone']} {site['Last Online Time']}")

        # Check if required columns exist for Alarm Report
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            # Add a new column for Client extracted from Site Alias
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            
            # Drop rows where Client extraction failed
            alarm_df = alarm_df[~alarm_df['Client'].isnull()]
            
            # Get the list of unique alarm names and sort them based on priority
            alarm_names = sorted(alarm_df['Alarm Name'].unique(), key=lambda x: alarm_priority.get(x, float('inf')))
            
            # Create a dictionary to store the pivot tables and total alarm counts
            alarm_data = {}
            for alarm_name in alarm_names:
                pivot_table, total_alarm_count = create_pivot_table(alarm_df, alarm_name)
                alarm_data[alarm_name] = (pivot_table, total_alarm_count)
                
            # Display each alarm report
            for alarm_name, (pivot_table, total_alarm_count) in alarm_data.items():
                st.markdown(f"### {alarm_name}")
                st.markdown(f"**Total Alarm Count:** {total_alarm_count}")
                st.dataframe(pivot_table)
                
            # Create a single download button for all alarms
            all_alarms_excel_data = to_excel(alarm_data)
            st.download_button(
                label="Download All Alarms Report as Excel",
                data=all_alarms_excel_data,
                file_name=alarm_download_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
