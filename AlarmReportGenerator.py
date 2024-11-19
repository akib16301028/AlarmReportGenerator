import pandas as pd
import streamlit as st

def load_alarm_data(file):
    try:
        df = pd.read_excel(file)
        st.write("Alarm data loaded successfully!")
        return df
    except Exception as e:
        st.error(f"Error loading alarm data: {e}")
        return pd.DataFrame()

def create_pivot_table(df, alarm_name):
    try:
        filtered_df = df[df['Alarm Name'] == alarm_name]
        pivot = pd.pivot_table(
            filtered_df,
            values='Count',
            index='Zone',
            aggfunc='sum',
            fill_value=0
        )
        total_count = pivot['Count'].sum()
        return pivot, total_count
    except Exception as e:
        st.error(f"Error creating pivot table: {e}")
        return pd.DataFrame(), 0

def to_excel(data):
    try:
        output = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
        for sheet_name, df in data.items():
            df.to_excel(output, sheet_name=sheet_name)
        output.close()
        st.write("Excel file created successfully!")
    except Exception as e:
        st.error(f"Error writing to Excel: {e}")

# Streamlit App
st.title("Alarm Report Generator")

# File Upload
uploaded_file = st.file_uploader("Upload Alarm Data", type=["xlsx"])
if uploaded_file:
    alarm_df = load_alarm_data(uploaded_file)
    if not alarm_df.empty:
        st.write("First 5 rows of Alarm Data:")
        st.write(alarm_df.head())

        # Debug: Check data types
        st.write("Data Types:")
        st.write(alarm_df.dtypes)

        # Select Alarm Names
        alarm_names = alarm_df['Alarm Name'].unique()
        st.write("Available Alarm Names:", alarm_names)

        # Select Alarm Name for Report
        selected_alarm = st.selectbox("Select an Alarm Name", alarm_names)
        if selected_alarm:
            st.write(f"Generating report for: {selected_alarm}")

            try:
                pivot_table, total = create_pivot_table(alarm_df, selected_alarm)
                st.write(f"Pivot Table for {selected_alarm}:")
                st.write(pivot_table)
                st.write(f"Total Count: {total}")
            except Exception as e:
                st.error(f"Error processing alarm data: {e}")

            # Export to Excel
            if st.button("Download Report"):
                try:
                    excel_data = {'Alarm Report': pivot_table}
                    to_excel(excel_data)
                    st.success("Report ready for download!")
                except Exception as e:
                    st.error(f"Error generating Excel report: {e}")
else:
    st.info("Please upload a valid Excel file.")
