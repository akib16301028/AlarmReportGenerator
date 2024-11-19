import pandas as pd
import streamlit as st
from datetime import datetime

# Helper Functions
def create_pivot_table(filtered_df, alarm_name):
    """
    Creates a pivot table for a given filtered DataFrame and returns it along with the total count.
    """
    try:
        if filtered_df.empty:
            return pd.DataFrame(), 0
        pivot = pd.pivot_table(
            filtered_df,
            index=['Cluster'],
            values=['Alarm Count'],
            aggfunc='sum',
            fill_value=0
        )
        total_count = filtered_df['Alarm Count'].sum()
        return pivot, total_count
    except Exception as e:
        st.error(f"Error in creating pivot table for {alarm_name}: {e}")
        return pd.DataFrame(), 0

def style_dataframe(df, duration_cols, dark_mode):
    """
    Styles the pivot table DataFrame for better visualization.
    """
    if df.empty:
        return df
    # Add specific styling (e.g., color gradients for duration columns)
    styled = df.style.background_gradient(
        subset=duration_cols,
        cmap="coolwarm" if dark_mode else "viridis"
    )
    return styled

def to_excel(dataframes_dict):
    """
    Converts multiple DataFrames into an Excel file for download.
    """
    from io import BytesIO
    with pd.ExcelWriter(BytesIO(), engine="xlsxwriter") as writer:
        for sheet_name, df in dataframes_dict.items():
            df.to_excel(writer, sheet_name=sheet_name)
        writer.save()
        return writer.book.getvalue()

# Streamlit App
st.set_page_config(page_title="Alarm Monitoring", layout="wide")

st.title("Alarm Monitoring and Reporting Tool")

try:
    # Sidebar for File Upload
    uploaded_file = st.sidebar.file_uploader("Upload Alarm Data File (CSV/Excel)", type=["csv", "xlsx"])
    if not uploaded_file:
        st.warning("Please upload a valid data file to proceed.")
        st.stop()

    # Read the uploaded file
    if uploaded_file.name.endswith(".csv"):
        alarm_df = pd.read_csv(uploaded_file)
    else:
        alarm_df = pd.read_excel(uploaded_file)

    # Validate necessary columns
    required_columns = ['Alarm Name', 'Alarm Time', 'Cluster', 'RMS Station', 'Alarm Count']
    if not all(col in alarm_df.columns for col in required_columns):
        st.error(f"The uploaded file must contain the following columns: {', '.join(required_columns)}")
        st.stop()

    # Sidebar Filters
    selected_alarm = st.sidebar.selectbox("Select Alarm", options=["All"] + sorted(alarm_df['Alarm Name'].unique()))
    selected_offline_cluster = st.sidebar.selectbox(
        "Select Cluster",
        options=["All"] + sorted(alarm_df['Cluster'].unique())
    )
    
    dark_mode = st.sidebar.checkbox("Enable Dark Mode", value=False)

    # Priority order for alarms
    priority_order = [
        'Mains Fail',
        'Battery Low',
        'DCDB-01 Primary Disconnect',
        'PG Run',
        'MDB Fault',
        'Door Open'
    ]

    # Separate prioritized and non-prioritized alarms
    prioritized_alarms = [name for name in priority_order if name in alarm_df['Alarm Name'].unique()]
    non_prioritized_alarms = [name for name in alarm_df['Alarm Name'].unique() if name not in priority_order]
    ordered_alarm_names = prioritized_alarms + non_prioritized_alarms

    # Process Alarms
    alarm_data = {}
    for alarm_name in ordered_alarm_names:
        # Skip unselected alarms
        if selected_alarm != "All" and alarm_name != selected_alarm:
            continue

        filtered_alarm_df = alarm_df.copy()

        if selected_alarm != "All":
            filtered_alarm_df = filtered_alarm_df[filtered_alarm_df['Alarm Name'] == alarm_name]
            if selected_offline_cluster != "All":
                filtered_alarm_df = filtered_alarm_df[filtered_alarm_df['Cluster'] == selected_offline_cluster]

            # Parse and filter dates
filtered_alarm_df['Alarm Time Parsed'] = pd.to_datetime(
    filtered_alarm_df['Alarm Time'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce'
)

# Handle empty or invalid dates gracefully
if filtered_alarm_df['Alarm Time Parsed'].isna().all():
    st.warning(f"No valid dates found for alarm: {alarm_name}")
    continue  # Skip to the next alarm if all dates are invalid

# Drop rows with NaT in parsed dates
filtered_alarm_df = filtered_alarm_df.dropna(subset=['Alarm Time Parsed'])

# Define minimum and maximum dates for the filter
min_date = filtered_alarm_df['Alarm Time Parsed'].min().date()
max_date = filtered_alarm_df['Alarm Time Parsed'].max().date()

# Sidebar for date range input
selected_date_range = st.sidebar.date_input(
    f"Date Range for {alarm_name}",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
    key=f"date_{alarm_name}"
)

# Ensure date range is a tuple of two dates
if isinstance(selected_date_range, tuple) and len(selected_date_range) == 2:
    start_date, end_date = selected_date_range
    filtered_alarm_df = filtered_alarm_df[
        (filtered_alarm_df['Alarm Time Parsed'].dt.date >= start_date) &
        (filtered_alarm_df['Alarm Time Parsed'].dt.date <= end_date)
    ]
else:
    st.error(f"Invalid date range selected for alarm: {alarm_name}")
            if isinstance(selected_date_range, tuple) and len(selected_date_range) == 2:
                start_date, end_date = selected_date_range
                filtered_alarm_df = filtered_alarm_df[
                    (filtered_alarm_df['Alarm Time Parsed'].dt.date >= start_date) &
                    (filtered_alarm_df['Alarm Time Parsed'].dt.date <= end_date)
                ]

        # Special filter for "DCDB-01 Primary Disconnect"
        if alarm_name == 'DCDB-01 Primary Disconnect':
            filtered_alarm_df = filtered_alarm_df[~filtered_alarm_df['RMS Station'].str.startswith('L')]

        # Create pivot table
        pivot, total_count = create_pivot_table(filtered_alarm_df, alarm_name)
        alarm_data[alarm_name] = (pivot, total_count)

    # Display Pivot Tables
    for alarm_name, (pivot, total_count) in alarm_data.items():
        st.markdown(f"### **{alarm_name}**")
        st.markdown(f"**Total Alarm Count:** {total_count}")
        duration_cols = ['0+', '2+', '4+', '8+']  # Example duration columns
        styled_pivot = style_dataframe(pivot, duration_cols, dark_mode)
        st.dataframe(styled_pivot)

    # Export Report
    if alarm_data:
        excel_data = to_excel({alarm: data[0] for alarm, data in alarm_data.items()})
        st.download_button(
            label="Download Current Alarms Report",
            data=excel_data,
            file_name=f"Current_Alarms_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No data available for export.")
except Exception as e:
    st.error(f"An error occurred: {e}")
