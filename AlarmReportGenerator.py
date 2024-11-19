import pandas as pd
import streamlit as st

# Function to sanitize Duration Slot (Hours)
def sanitize_duration(df, duration_col='Duration Slot (Hours)', max_duration=24):
    """
    Cleans and caps values in the Duration Slot column.
    """
    df[duration_col] = pd.to_numeric(df[duration_col], errors='coerce')  # Convert to numeric, NaN for invalid
    df[duration_col] = df[duration_col].fillna(0)  # Replace NaN with 0
    df[duration_col] = df[duration_col].clip(upper=max_duration)  # Cap values at max_duration
    return df

# Function to normalize timestamps
def normalize_timestamps(df, time_col='Alarm Time', datetime_format='%d/%m/%Y %I:%M:%S %p'):
    """
    Ensures that all timestamps are in a consistent format.
    Invalid timestamps are replaced with NaT.
    """
    df[time_col] = pd.to_datetime(df[time_col], format=datetime_format, errors='coerce')  # Parse dates
    invalid_rows = df[time_col].isna().sum()  # Check for invalid dates
    if invalid_rows > 0:
        st.warning(f"Warning: {invalid_rows} invalid timestamps detected and replaced with NaT.")
    return df

# Main processing function
def process_alarm_file(file_path):
    """
    Main function to process the alarm file.
    Applies sanitization and timestamp normalization.
    """
    try:
        # Load Excel file
        df = pd.read_excel(file_path, sheet_name='RMS Station Current Alarm Repor')

        # Step 1: Sanitize Duration Slot (Hours)
        df = sanitize_duration(df)

        # Step 2: Normalize Timestamps
        df = normalize_timestamps(df)

        # Additional processing if required
        st.success("File processed successfully!")
        return df

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
        return None

# Streamlit UI
st.title("Alarm Report Processor")

uploaded_file = st.file_uploader("Upload Alarm Report", type=["xlsx"])

if uploaded_file:
    st.write("Processing file:", uploaded_file.name)
    processed_df = process_alarm_file(uploaded_file)

    if processed_df is not None:
        st.write("Processed Data:")
        st.dataframe(processed_df)

        # Option to download the processed file
        csv_data = processed_df.to_csv(index=False)
        st.download_button(
            label="Download Processed Data",
            data=csv_data,
            file_name="Processed_Alarm_Report.csv",
            mime="text/csv"
        )
