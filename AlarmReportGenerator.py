# Streamlit app
st.title("Alarm and Offline Data Pivot Table Generator")

uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)

        current_time = extract_timestamp(uploaded_alarm_file.name)
        offline_time = extract_timestamp(uploaded_offline_file.name)

        # Process the Offline Report
        pivot_offline, total_offline_count = create_offline_pivot(offline_df)

        st.markdown("### Offline Report")
        st.markdown(f"<small><i>till {offline_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)
        st.markdown(f"**Total Offline Count:** {total_offline_count}")
        st.dataframe(pivot_offline)

        # Calculate time offline smartly using the offline time
        time_offline_df = calculate_time_offline(offline_df, offline_time)

        # Create a summary table based on offline duration
        summary_dict = {}
        for index, row in time_offline_df.iterrows():
            duration = row['Offline Duration']
            if duration not in summary_dict:
                summary_dict[duration] = []
            summary_dict[duration].append(row)

        # Prepare DataFrame for display
        summary_data = []
        for duration, sites in summary_dict.items():
            for site in sites:
                summary_data.append([duration, site['Site Alias'], site['Cluster'], site['Zone'], site['Last Online Time']])
        
        # Display the summary table
        summary_df = pd.DataFrame(summary_data, columns=["Offline Duration", "Site Name (Site Alias)", "Cluster", "Zone", "Last Online Time"])
        st.markdown("### Summary of Offline Sites")
        st.dataframe(summary_df)

        # Check for required columns in Alarm Report
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name', 'Alarm Time']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            alarm_df = alarm_df[~alarm_df['Client'].isnull()]

            # Prepare download for Offline Report
            offline_report_data = {
                "Offline Summary": pivot_offline,
                "Offline Details": summary_df
            }
            offline_excel_data = to_excel(offline_report_data)

            st.download_button(
                label="Download Offline Report",
                data=offline_excel_data,
                file_name=f"Offline Report_{offline_time.strftime('%Y-%m-%d %H-%M-%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Add the current time to the alarm header
            st.markdown(f"### Current Alarms Report")
            st.markdown(f"<small><i>till {current_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)

            # Date range input for all alarms
            date_range = st.date_input("Select Date Range for Alarms", [])

            alarm_names = alarm_df['Alarm Name'].unique()

            # Define the priority order for the alarm names
            priority_order = [
                'Mains Fail',
                'Battery Low',
                'DCDB-01 Primary Disconnect',
                'PG Run',
                'MDB Fault',
                'Door Open'
            ]

            # Separate prioritized alarms from the rest
            prioritized_alarms = [name for name in priority_order if name in alarm_names]
            non_prioritized_alarms = [name for name in alarm_names if name not in priority_order]

            # Combine both lists to maintain the desired order
            ordered_alarm_names = prioritized_alarms + non_prioritized_alarms

            # Create a dictionary to store all pivot tables for current alarms
            alarm_data = {}

            for alarm_name in ordered_alarm_names:
                data = create_pivot_table(alarm_df, alarm_name)

                # Apply time filtering for all alarms
                if date_range:
                    # Filter the DataFrame based on the selected date range
                    filtered_data = alarm_df[
                        (alarm_df['Alarm Name'] == alarm_name) &
                        (pd.to_datetime(alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p').dt.date.isin(date_range))
                    ]
                    alarm_data[alarm_name] = create_pivot_table(filtered_data, alarm_name)
                else:
                    alarm_data[alarm_name] = data

            # Display each pivot table for the current alarms
            for alarm_name, (pivot, total_count) in alarm_data.items():
                st.markdown(f"### <b>{alarm_name}</b>", unsafe_allow_html=True)
                st.markdown(f"<small><i>till {current_time.strftime('%Y-%m-%d %H:%M:%S')}</i></small>", unsafe_allow_html=True)
                st.markdown(f"<small><i>Alarm Count: {total_count}</i></small>", unsafe_allow_html=True)
                st.dataframe(pivot)

            # Prepare download for Current Alarms Report only if there is data
            if alarm_data:
                current_alarm_excel_data = to_excel({alarm_name: data[0] for alarm_name, data in alarm_data.items()})
                st.download_button(
                    label="Download Current Alarms Report",
                    data=current_alarm_excel_data,
                    file_name=f"Current Alarms Report_{current_time.strftime('%Y-%m-%d %H-%M-%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No current alarm data available for export.")
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
