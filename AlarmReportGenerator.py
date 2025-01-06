import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

# Function to extract client name from Site Alias
def extract_client(site_alias):
    match = re.search(r'\((.*?)\)', site_alias)
    return match.group(1) if match else None

# Function to categorize Duration Slot (Hours)
def categorize_duration(hours):
    if 0 <= hours < 2:
        return '0+'
    elif 2 <= hours < 4:
        return '2+'
    elif 4 <= hours < 8:
        return '4+'
    elif hours >= 8:
        return '8+'
    else:
        return 'Unknown'

# Function to create pivot table for a specific alarm
def create_pivot_table(df, alarm_name):
    alarm_df = df[df['Alarm Name'] == alarm_name].copy()
    
    if alarm_name == 'DCDB-01 Primary Disconnect':
        alarm_df = alarm_df[~alarm_df['RMS Station'].str.startswith('L')]
    
    # Categorize Duration Slot (Hours)
    alarm_df['Duration Category'] = alarm_df['Duration Slot (Hours)'].apply(categorize_duration)
    
    # Create pivot tables for each Duration Category
    pivot_total = pd.pivot_table(
        alarm_df,
        index=['Cluster', 'Zone'],
        columns='Client',
        values='Site Alias',
        aggfunc='count',
        fill_value=0
    ).reset_index()
    
    client_columns = [col for col in pivot_total.columns if col not in ['Cluster', 'Zone']]
    pivot_total['Total'] = pivot_total[client_columns].sum(axis=1)
    
    # Pivot for Duration Categories
    pivot_duration = pd.pivot_table(
        alarm_df,
        index=['Cluster', 'Zone'],
        columns='Duration Category',
        values='Site Alias',
        aggfunc='count',
        fill_value=0
    ).reset_index()
    
    # Merge the two pivot tables
    pivot = pd.merge(pivot_total, pivot_duration, on=['Cluster', 'Zone'], how='left')
    
    # Ensure all Duration Categories are present
    for cat in ['0+', '2+', '4+', '8+']:
        if cat not in pivot.columns:
            pivot[cat] = 0
    
    # Reorder columns: Original client columns + 'Total' + Duration Categories
    pivot = pivot[['Cluster', 'Zone'] + client_columns + ['Total', '0+', '2+', '4+', '8+']]
    
    # Add Total row
    numeric_cols = pivot.select_dtypes(include=['number']).columns
    total_row = pivot[numeric_cols].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    # Replace numeric columns in total_row with empty strings
    total_row[numeric_cols] = total_row[numeric_cols].replace(0, "0").astype(str)
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    total_alarm_count = pivot['Total'].iloc[-1]
    
    # Remove repeated Cluster names for better readability
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_alarm_count

import pandas as pd

# Function to create pivot table for offline report
def create_offline_pivot(df):
    df = df.drop_duplicates()
    
    # Create new columns for each duration category
    df['Less than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'Less than 24 hours' in x else 0)
    df['More than 24 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 24 hours' in x and '72' not in x else 0)
    df['More than 48 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 48 hours' in x else 0)
    df['More than 72 hours'] = df['Duration'].apply(lambda x: 1 if 'More than 72 hours' in x else 0)
    
    pivot = df.groupby(['Cluster', 'Zone']).agg({
        'Less than 24 hours': 'sum',
        'More than 24 hours': 'sum',
        'More than 48 hours': 'sum',  # Include the new column here
        'More than 72 hours': 'sum',
        'Site Alias': 'nunique'
    }).reset_index()

    pivot = pivot.rename(columns={'Site Alias': 'Total'})
    
    total_row = pivot[['Less than 24 hours', 'More than 24 hours', 'More than 48 hours', 'More than 72 hours', 'Total']].sum().to_frame().T
    total_row[['Cluster', 'Zone']] = ['Total', '']
    
    # Replace numeric columns in total_row with empty strings
    numeric_cols = ['Less than 24 hours', 'More than 24 hours', 'More than 48 hours', 'More than 72 hours', 'Total']
    total_row[numeric_cols] = total_row[numeric_cols].replace(0, "0").astype(str)
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    total_offline_count = int(pivot['Total'].iloc[-1])
    
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_offline_count

    
    # Replace numeric columns in total_row with empty strings
    numeric_cols = ['Less than 24 hours', 'More than 24 hours','More than 48 hours', 'More than 72 hours', 'Total']
    total_row[numeric_cols] = total_row[numeric_cols].replace(0, "0").astype(str)
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    total_offline_count = int(pivot['Total'].iloc[-1])
    
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_offline_count

    # Setting empty cells for the '48+' column if the value is zero
    pivot['More than 48 hours'] = pivot['More than 48 hours'].replace("", "0")  # This line ensures that if the value is 0, it stays empty

    return pivot, total_offline_count


    
    # Replace numeric columns in total_row with empty strings
    numeric_cols = ['Less than 24 hours', 'More than 24 hours','More than 48 hours', 'More than 72 hours', 'Total']
    total_row[numeric_cols] = total_row[numeric_cols].replace(0, "0").astype(str)
    
    pivot = pd.concat([pivot, total_row], ignore_index=True)
    
    total_offline_count = int(pivot['Total'].iloc[-1])
    
    last_cluster = None
    for i in range(len(pivot)):
        if pivot.at[i, 'Cluster'] == last_cluster:
            pivot.at[i, 'Cluster'] = ''
        else:
            last_cluster = pivot.at[i, 'Cluster']
    
    return pivot, total_offline_count

# Function to calculate duration for Offline Summary
def calculate_duration(df):
    current_time = datetime.now()
    df['Last Online Time'] = pd.to_datetime(df['Last Online Time'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce')
    df['Duration'] = (current_time - df['Last Online Time']).dt.days
    return df[['Site Alias', 'Zone', 'Cluster', 'Duration']]


# Function to convert multiple DataFrames to Excel with separate sheets
def to_excel(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            valid_sheet_name = re.sub(r'[\\/*?:[\]]', '_', sheet_name)[:31]
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    return output.getvalue()



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
    filtered_df = filtered_df[['Site Alias', 'Cluster', 'Zone', 'Alarm Name', 'Alarm Time','Duration']]
    filtered_df = filtered_df.sort_values(by='Alarm Time', ascending=False)
    return filtered_df

def style_dataframe(df, duration_cols, is_dark_mode):
    # Create a copy for styling
    df_style = df.copy()
    
    # Identify the total row based on 'Cluster' column
    total_row_mask = df_style['Cluster'] == 'Total'
    
    # Define background colors
    cell_bg_color = '#f0f0f0'
    font_color = 'black' if not is_dark_mode else 'black'
    
    # Create a Styler object
    styler = df_style.style
    
    # Apply background color to cells with non zero 
    def highlight_zero(val):
        if val != 0:
            return f'background-color: {cell_bg_color}; color: {font_color}'
        return ''
    
    styler = styler.applymap(highlight_zero)
    
    # Handle total row: set all cells to empty except 'Cluster' and 'Zone' if needed
    if total_row_mask.any():
        styler = styler.apply(
            lambda x: ['background-color: #f0f0f0; color: black' if total_row_mask.loc[x.name] else '' for _ in x],
            axis=1
        )
        # Optionally, you can set the 'Cluster' and 'Zone' cells to have a different style
        styler = styler.applymap(
            lambda x: f'background-color: {cell_bg_color}; color: {font_color}',
            subset=['Cluster', 'Zone']
        )
    
    # Optional: Remove borders for a cleaner look
    styler.set_table_styles(
        [{
            'selector': 'th',
            'props': [('border', '1px solid black')]
        },
        {
            'selector': 'td',
            'props': [('border', '1px solid black')]
        }]
    )
    
    return styler


# Function to determine if the current theme is dark
def is_dark_mode():
    # Streamlit provides theme options that can be accessed via st.get_option
    # As of Streamlit 1.10, you can access the theme via st.runtime
    # However, this may vary based on the Streamlit version
    # Here, we'll use st.session_state as a workaround

    # Check if 'theme' is in session_state
    if 'theme' in st.session_state:
        theme = st.session_state['theme']
    else:
        # Default to light mode if not set
        theme = 'light'

    # Assume 'dark' indicates dark mode
    return theme.lower() == 'dark'

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
        # Note: Streamlit does not provide a direct method to detect theme,
        # so this function is a placeholder and may need adjustment based on Streamlit version
        dark_mode = is_dark_mode()

        # Process the Offline Report
        pivot_offline, total_offline_count = create_offline_pivot(offline_df)

        # Apply Offline Cluster Filters
        if selected_offline_cluster != "All":
            if 'Total' in pivot_offline['Cluster'].values:
                filtered_pivot_offline = pivot_offline[
                    (pivot_offline['Cluster'] == selected_offline_cluster) | (pivot_offline['Cluster'] == 'Total')
                ]
            else:
                filtered_pivot_offline = pivot_offline[pivot_offline['Cluster'] == selected_offline_cluster]
        else:
            filtered_pivot_offline = pivot_offline.copy()

        # Display the Offline Report
        st.markdown("### Offline Report")
        st.markdown(f"**Total Offline Count:** {total_offline_count}")

        # Apply styling
        styled_pivot_offline = style_dataframe(filtered_pivot_offline, ['Less than 24 hours', 'More than 24 hours', 'More than 72 hours'], dark_mode)

        # Display styled DataFrame
        st.dataframe(styled_pivot_offline)

        # Generate Offline Summary Table
        st.markdown("### Offline Summary Table")
        offline_summary_df = calculate_duration(offline_df)
        st.dataframe(offline_summary_df)

        # Downloadable Offline Summary Table
        offline_summary_excel = to_excel({"Offline Summary": offline_summary_df})
        st.download_button(
            label="Download Offline Summary",
            data=offline_summary_excel,
            file_name="Offline_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )



        # === Site-Wise Log Display ===
        if view_site_wise:
            st.markdown("### Site-Wise Log")
            if site_wise_alarms != "All":
                site_wise_log_df = create_site_wise_log(alarm_df, site_wise_alarms)
                # Apply styling if needed
                styled_site_wise_log = style_dataframe(site_wise_log_df, [], dark_mode)
                st.dataframe(styled_site_wise_log)
            else:
                st.info("No specific alarm selected for Site-Wise Log.")

        # Check for required columns in Alarm Report
        alarm_required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name', 'Alarm Time', 'Duration Slot (Hours)']
        if not all(col in alarm_df.columns for col in alarm_required_columns):
            st.error(f"The uploaded Alarm Report file is missing one of the required columns: {alarm_required_columns}")
        else:
            # Extract client information
            alarm_df['Client'] = alarm_df['Site Alias'].apply(extract_client)
            alarm_df = alarm_df[~alarm_df['Client'].isnull()]

 
            # Add the current time to the alarm header
            st.markdown(f"### Current Alarms Report")

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
                    
                    # Apply cluster filter
                    if selected_offline_cluster != "All":
                        filtered_alarm_df = filtered_alarm_df[filtered_alarm_df['Cluster'] == selected_offline_cluster]
                    
                    # Apply date range filter
                    alarm_dates = pd.to_datetime(filtered_alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce')
                    min_date = alarm_dates.min().date()
                    max_date = alarm_dates.max().date()
                    selected_date_range = st.sidebar.date_input(
                        f"Select Date Range for {alarm_name}",
                        value=(min_date, max_date),
                        min_value=min_date,
                        max_value=max_date,
                        key=f"date_{alarm_name}"
                    )
                    # Ensure date range is a tuple of two dates
                    if isinstance(selected_date_range, tuple) and len(selected_date_range) == 2:
                        start_date, end_date = selected_date_range
                    else:
                        start_date, end_date = min_date, max_date

                    filtered_alarm_df['Alarm Time Parsed'] = pd.to_datetime(
                        filtered_alarm_df['Alarm Time'], 
                        format='%d/%m/%Y %I:%M:%S %p', 
                        errors='coerce'
                    )
                    filtered_alarm_df = filtered_alarm_df[
                        (filtered_alarm_df['Alarm Time Parsed'].dt.date >= start_date) &
                        (filtered_alarm_df['Alarm Time Parsed'].dt.date <= end_date)
                    ]

                # Special filter for "DCDB-01 Primary Disconnect"
                if alarm_name == 'DCDB-01 Primary Disconnect':
                    filtered_alarm_df = filtered_alarm_df[~filtered_alarm_df['RMS Station'].str.startswith('L')]

                # Create pivot table for the filtered data
                pivot, total_count = create_pivot_table(filtered_alarm_df, alarm_name)
                alarm_data[alarm_name] = (pivot, total_count)

                # Display each pivot table for the current alarms with styling
                try:
                # Your code that may raise an exception
                alarm_df = pd.read_excel(uploaded_alarm_file, header=2)

                # Extract the value from the second row, first column (this will be the new line to add under the header)
                additional_header_info = alarm_df.iloc[1, 0]  # Extracts value from row 2, column 1
                
                # Display the value under the current alarms report heading
                st.markdown(f"### Current Alarms Report")
                st.markdown(f"**Alarm Count:** {total_count}")
                st.markdown(f"#### {additional_header_info}")  # Display the additional header info

                # Proceed with your logic to process and display alarms
                # ...
    
                except Exception as e:
                # Handle exceptions here
                st.error(f"An error occurred while processing the Alarm Report: {e}")
                
                finally:
                # Optional: Any cleanup code can go here, if needed
                pass

                # Identify duration columns
                duration_cols = ['0+', '2+', '4+', '8+']

                # Apply styling
                styled_pivot = style_dataframe(pivot, duration_cols, dark_mode)

                # Display styled DataFrame
                st.dataframe(styled_pivot)

            # Prepare download for Current Alarms Report only if there is data
            if alarm_data:
                # Create a dictionary with each alarm's pivot table
                current_alarm_excel_dict = {alarm_name: data[0] for alarm_name, data in alarm_data.items()}
                current_alarm_excel_data = to_excel(current_alarm_excel_dict)
                st.download_button(
                    label="Download Current Alarms Report",
                    data=current_alarm_excel_data,
                    file_name=f"Current_Alarms_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No current alarm data available for export.")
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
