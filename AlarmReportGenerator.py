import pandas as pd
import streamlit as st
from io import BytesIO

# Helper function to create pivot table
def create_pivot_table(df, alarm_name):
    if df.empty:
        return pd.DataFrame(), 0

    pivot = pd.pivot_table(
        df,
        index='Cluster',
        values='Alarm Count',
        aggfunc='sum',
        fill_value=0
    )
    total_count = df['Alarm Count'].sum()
    return pivot, total_count

# Helper function to style the DataFrame
def style_dataframe(df, duration_cols, dark_mode):
    if df.empty:
        return df

    # Apply custom styles
    styled_df = df.style
    for col in duration_cols:
        if col in df.columns:
            styled_df = styled_df.background_gradient(subset=col, cmap='Reds' if dark_mode else 'Blues')
    return styled_df

# Function to export data to Excel
def to_excel(data_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            df.to_excel(writer, index=True, sheet_name=sheet_name)
    return output.getvalue()

# Streamlit app starts here
try:
    # Sidebar inputs
    st.sidebar.title("Filter Options")
    selected_alarm = st.sidebar.selectbox(
        "Select Alarm",
        options=["All", "Mains Fail", "Battery Low", "DCDB-01 Primary Disconnect", "PG Run", "MDB Fault", "Door Open"]
    )
    selected_offline_cluster = st.sidebar.selectbox("Select Cluster", options=["All", "Cluster1", "Cluster2", "Cluster3"])

    # File upload
    uploaded_file = st.file_uploader("Upload Alarm Data File", type=["csv", "xlsx"])
    if uploaded_file:
        # Read the data
        if uploaded_file.name.endswith('.csv'):
            alarm_df = pd.read_csv(uploaded_file)
        else:
            alarm_df = pd.read_excel(uploaded_file)

        # Ensure required columns are present
        required_columns = ['Alarm Name', 'Cluster', 'Alarm Count', 'RMS Station']
        if not all(col in alarm_df.columns for col in required_columns):
            st.error(f"Uploaded file must contain the following columns: {', '.join(required_columns)}")
        else:
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
            prioritized_alarms = [name for name in priority_order if name in alarm_df['Alarm Name'].unique()]
            non_prioritized_alarms = [name for name in alarm_df['Alarm Name'].unique() if name not in priority_order]

            # Combine both lists to maintain the desired order
            ordered_alarm_names = prioritized_alarms + non_prioritized_alarms

            # Create a dictionary to store all pivot tables for current alarms
            alarm_data = {}

            # Process alarms based on selection
            for alarm_name in ordered_alarm_names:
                if selected_alarm != "All" and alarm_name != selected_alarm:
                    continue

                # Filter the DataFrame
                filtered_alarm_df = alarm_df.copy()
                filtered_alarm_df = filtered_alarm_df[filtered_alarm_df['Alarm Name'] == alarm_name]

                if selected_offline_cluster != "All":
                    filtered_alarm_df = filtered_alarm_df[filtered_alarm_df['Cluster'] == selected_offline_cluster]

                # Special filter for "DCDB-01 Primary Disconnect"
                if alarm_name == 'DCDB-01 Primary Disconnect':
                    filtered_alarm_df = filtered_alarm_df[~filtered_alarm_df['RMS Station'].str.startswith('L')]

                # Create pivot table
                pivot, total_count = create_pivot_table(filtered_alarm_df, alarm_name)
                alarm_data[alarm_name] = (pivot, total_count)

            # Display each pivot table
            for alarm_name, (pivot, total_count) in alarm_data.items():
                st.markdown(f"### **{alarm_name}**")
                st.markdown(f"**Alarm Count:** {total_count}")

                duration_cols = ['0+', '2+', '4+', '8+']
                styled_pivot = style_dataframe(pivot, duration_cols, dark_mode=False)
                st.dataframe(styled_pivot)

            # Export to Excel
            if alarm_data:
                current_alarm_excel_dict = {alarm_name: data[0] for alarm_name, data in alarm_data.items()}
                current_alarm_excel_data = to_excel(current_alarm_excel_dict)
                st.download_button(
                    label="Download Current Alarms Report",
                    data=current_alarm_excel_data,
                    file_name="Current_Alarms_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No current alarm data available for export.")
    else:
        st.info("Please upload a file to get started.")

except Exception as e:
    st.error(f"An error occurred while processing the files: {e}")
