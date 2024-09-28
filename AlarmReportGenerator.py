import pandas as pd
import streamlit as st

# Function to load the Excel file
def load_data(file):
    """Load data from the uploaded Excel file, starting from the third row."""
    df = pd.read_excel(file, header=2)  # Adjust header row as needed
    return df

# Function to generate the pivot table
def generate_pivot_table(df):
    """Generate a pivot table based on the specified columns."""
    pivot = df.pivot_table(
        index=['Cluster', 'Zone'],  # Grouping by Cluster and Zone
        columns='Duration',          # Pivoting on Duration
        values='Site Alias',         # Values to count
        aggfunc='count',             # Count occurrences
        fill_value=0                 # Fill empty values with 0
    )
    return pivot

# Streamlit interface
st.title("Alarm Report Analysis")

# Upload Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    st.subheader("Data Preview")
    st.dataframe(df)  # Display the loaded data

    # Generate the pivot table
    pivot_table = generate_pivot_table(df)

    st.subheader("Pivot Table")
    st.dataframe(pivot_table)  # Display the pivot table

    # Export button for the pivot table
    if st.button("Download Pivot Table"):
        pivot_file = "pivot_table.xlsx"
        pivot_table.to_excel(pivot_file, index=True)  # Save pivot table to Excel
        st.success(f"Pivot table exported as {pivot_file}.")
