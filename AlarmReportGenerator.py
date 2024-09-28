# Streamlit app
st.title("Alarm Data Pivot Table Generator")

# Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file is not None:
    try:
        # Read the uploaded Excel file, assuming headers start from row 3 (0-indexed)
        df = pd.read_excel(uploaded_file, header=2)
        
        # Check if required columns exist
        required_columns = ['RMS Station', 'Cluster', 'Zone', 'Site Alias', 'Alarm Name']
        if not all(col in df.columns for col in required_columns):
            st.error(f"The uploaded file is missing one of the required columns: {required_columns}")
        else:
            # Add a new column for Client extracted from Site Alias
            df['Client'] = df['Site Alias'].apply(extract_client)
            
            # Drop rows where Client extraction failed
            df = df.dropna(subset=['Client'])
            
            # Define priority alarms
            priority_alarms = [
                'Mains Fail', 
                'Battery Low', 
                'DCDB-01 Primary Disconnect', 
                'MDB Fault', 
                'PG Run', 
                'Door Open'
            ]
            
            # Get unique alarms, including non-priority ones
            all_alarms = list(df['Alarm Name'].unique())
            non_priority_alarms = [alarm for alarm in all_alarms if alarm not in priority_alarms]
            ordered_alarms = priority_alarms + sorted(non_priority_alarms)
            
            # Dictionary to store pivot tables and total counts
            pivot_tables = {}
            
            for alarm in ordered_alarms:
                pivot, total_count = create_pivot_table(df, alarm)
                pivot_tables[alarm] = (pivot, total_count)
                
                # Use an expander to keep headers and alarm count visible
                with st.expander(f"{alarm} (Total Count: {int(total_count)})", expanded=False):
                    st.dataframe(pivot)  # Display the pivot table
            
            # Create download button
            excel_data = to_excel(pivot_tables)
            st.download_button(
                label="Download All Pivot Tables as Excel",
                data=excel_data,
                file_name="pivot_tables_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
