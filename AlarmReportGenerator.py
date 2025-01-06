import pandas as pd
import streamlit as st
from io import BytesIO
import re

# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            valid_sheet_name = re.sub(r'[\\/*?:[\]]', '_', sheet_name)[:31]
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    return output.getvalue()

# Function to create site-wise log table
def create_site_wise_log(df, selected_alarm):
    if selected_alarm == "All":
        filtered_df = df.copy()
    else:
        filtered_df = df[df['Alarm Name'] == selected_alarm].copy()
    filtered_df = filtered_df[['Site Alias', 'Cluster', 'Zone', 'Alarm Name', 'Alarm Time', 'Duration']]
    filtered_df = filtered_df.sort_values(by='Alarm Time', ascending=False)
    return filtered_df

def style_dataframe(df, duration_cols, is_dark_mode):
    # Create a copy for styling
    df_style = df.copy()

    # Define background colors
    cell_bg_color = '#f0f0f0'
    font_color = 'black' if not is_dark_mode else 'black'

    # Create a Styler object
    styler = df_style.style

    # Apply background color to cells with non-zero values
    def highlight_zero(val):
        if val != 0:
            return f'background-color: {cell_bg_color}; color: {font_color}'
        return ''

    styler = styler.applymap(highlight_zero)

    return styler

# Function to determine if the current theme is dark
def is_dark_mode():
    # Assume the theme is light by default
    return 'dark' in st.session_state.get('theme', 'light').lower()

# Streamlit app
st.title("StatusMatrix@STL")

# File Uploads
uploaded_alarm_file = st.file_uploader("Upload Current Alarms Report", type=["xlsx"])
uploaded_offline_file = st.file_uploader("Upload Offline Report", type=["xlsx"])

if uploaded_alarm_file is not None and uploaded_offline_file is not None:
    try:
        # Read Excel files starting from the third row (header=2)
        alarm_df = pd.read_excel(uploaded_alarm_file, header=2)
        offline_df = pd.read_excel(uploaded_offline_file, header=2)

        # Initialize Sidebar Filters
        st.sidebar.header("Filters")

        # Get unique clusters for filtering
        offline_clusters = sorted(offline_df['Cluster'].dropna().unique().tolist())
        offline_clusters.insert(0, "All")  # Add 'All' option
        selected_offline_cluster = st.sidebar.selectbox(
            "Select Cluster",
            options=offline_clusters,
            index=0
        )

        # === Current Alarms Filters ===
        st.sidebar.subheader("Current Alarms Filters")
        st.sidebar.text("[select alarm first]")
        # Get unique alarm names
        alarm_names = sorted(alarm_df['Alarm Name'].dropna().unique().tolist())
        alarm_names.insert(0, "All")  # Add 'All' option
        selected_alarm = st.sidebar.selectbox(
            "Select Alarm to Filter",
            options=alarm_names,
            index=0
        )

        # === Site-Wise Log Filters ===
        st.sidebar.subheader("Site-Wise Log Filters")
        view_site_wise = st.sidebar.checkbox("View Site-Wise Log")
        if view_site_wise:
            site_wise_alarms = st.sidebar.selectbox(
                "Select Alarm for Site-Wise Log",
                options=alarm_names,
                index=0
            )

        # Determine if dark mode is active
        dark_mode = is_dark_mode()

        # Process the Offline Report
        pivot_offline, total_offline_count = create_site_wise_log(offline_df, selected_alarm)

        # Apply Offline Cluster Filters
        if selected_offline_cluster != "All":
            filtered_pivot_offline = pivot_offline[pivot_offline['Cluster'] == selected_offline_cluster]
        else:
            filtered_pivot_offline = pivot_offline.copy()

        # Display the Offline Report
        st.markdown("### Offline Report")
        st.markdown(f"**Total Offline Count:** {total_offline_count}")

        # Apply styling
        styled_pivot_offline = style_dataframe(filtered_pivot_offline, [], dark_mode)
        st.dataframe(styled_pivot_offline)

        # === Current Alarms Report ===
        if selected_alarm != "All":
            st.markdown(f"### Current Alarms Report for: {selected_alarm}")
        else:
            st.markdown("### Current Alarms Report")

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

        # Process alarms based on selection
        for alarm_name in ordered_alarm_names:
            # Skip alarms if a specific alarm is selected and it's not the current one
            if selected_alarm != "All" and alarm_name != selected_alarm:
                continue

            # Initialize the filter criteria
            filtered_alarm_df = alarm_df.copy()

            if selected_alarm != "All":
                # Filter by selected alarm
                filtered_alarm_df = filtered_alarm_df[filtered_alarm_df['Alarm Name'] == alarm_name]

            # Create pivot table for the filtered data
            pivot = create_site_wise_log(filtered_alarm_df, alarm_name)
            alarm_data[alarm_name] = pivot

        # Display each pivot table for the current alarms with styling
        for alarm_name, pivot in alarm_data.items():
            st.markdown(f"### **{alarm_name}**")

            # Add "Alarm Count" from the data frame
            total_count = len(pivot)
            st.markdown(f"**Alarm Count:** {total_count}")

            # Extract the additional information (row 2, column 1)
            additional_header_info = alarm_df.iloc[1, 0]  # Extracts value from row 2, column 1
            st.markdown(f"#### {additional_header_info}")

            # Apply styling to the pivot table
            duration_cols = ['0+', '2+', '4+', '8+']
            styled_pivot = style_dataframe(pivot, duration_cols, dark_mode)

            # Display the styled DataFrame
            st.dataframe(styled_pivot)

        # Prepare download for Current Alarms Report only if there is data
        if alarm_data:
            # Create a dictionary with each alarm's pivot table
            current_alarm_excel_dict = {alarm_name: data for alarm_name, data in alarm_data.items()}
            current_alarm_excel_data = to_excel(current_alarm_excel_dict)

            st.download_button(
                label="Download Current Alarms Report",
                data=current_alarm_excel_data,
                file_name="Current_Alarms_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No current alarm data available for export.")

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
