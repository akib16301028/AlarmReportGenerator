import streamlit as st
import pandas as pd

# Assuming other necessary imports and helper functions like `style_dataframe`, `to_excel` are already defined

try:
    # Display each pivot table for the current alarms with styling
    for alarm_name, (pivot, total_count) in alarm_data.items():
        st.markdown(f"### **{alarm_name}**")
        st.markdown(f"**Alarm Count:** {total_count}")

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
        
        # Download button with simple file name (no timestamp)
        st.download_button(
            label="Download Current Alarms Report",
            data=current_alarm_excel_data,
            file_name="Current_Alarms_Report.xlsx",  # No timestamp in the file name
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No current alarm data available for export.")

except Exception as e:
    st.error(f"An error occurred while processing the files: {e}")
