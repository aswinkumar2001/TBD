import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
import numpy as np

# Title of the app
st.title("Excel Date-Time Value Converter")

# File uploader with acceptance message
uploaded_file = st.file_uploader("Upload your Excel file (xlsx/xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    errors = []
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Check if "Date" column exists
        if "Date" not in df.columns:
            errors.append("Error: 'Date' column not found in the uploaded file.")
        else:
            # Define expected time fractions (48 half-hour increments, ending with 0.0)
            time_fractions = [i / 48 for i in range(1, 48)] + [0.0]  # 0.020833333 to 0.979166667, then 0.0
            column_headers = df.columns[1:]  # Exclude "Date" column

            # Convert column headers to numeric, handling strings
            try:
                column_headers_numeric = pd.to_numeric(column_headers, errors='coerce')
                if column_headers_numeric.isna().any():
                    errors.append("Error: Some time fraction columns could not be converted to numeric values.")
            except Exception as e:
                errors.append(f"Error converting column headers to numeric: {str(e)}.")

            # Debug: Print column headers for inspection
            st.write("Column headers (excluding Date):", list(column_headers))
            st.write("Numeric column headers:", list(column_headers_numeric))

            # Check if headers match expected time fractions with tolerance
            if not errors:
                tolerance = 1e-6  # Small tolerance for floating-point comparison
                expected_set = set(np.round(time_fractions, 8))
                actual_set = set(np.round(column_headers_numeric, 8))
                if actual_set != expected_set or len(column_headers) != 48:
                    errors.append("Error: Mismatch in expected time fraction columns. Expected exactly 48 fractions (0.020833333 to 0.979166667, 0.0).")
                    st.write("Expected fractions:", sorted(expected_set))
                    st.write("Actual fractions:", sorted(actual_set))
                    st.write(f"Expected number of columns: 48, Actual number: {len(column_headers)}")
            
            # Process the data
            if not errors:
                try:
                    # Melt the dataframe
                    melted_df = pd.melt(df, 
                                        id_vars=["Date"], 
                                        var_name="Time_Fraction", 
                                        value_name="Value")
                    
                    # Convert Time_Fraction to numeric
                    melted_df['Time_Fraction'] = pd.to_numeric(melted_df['Time_Fraction'], errors='coerce')
                    
                    # Handle Date column flexibly (text or Excel date format)
                    try:
                        # Try parsing as text date first (e.g., "Thursday, March 27, 2025")
                        melted_df['Date'] = pd.to_datetime(melted_df['Date'], 
                                                           format="%A, %B %d, %Y", 
                                                           errors='coerce')
                        # If still NaN, try Excel serial date
                        mask = melted_df['Date'].isna()
                        melted_df.loc[mask, 'Date'] = pd.to_datetime(
                            pd.to_numeric(melted_df.loc[mask, 'Date'], errors='coerce'), 
                            unit='D', 
                            origin='1899-12-30', 
                            errors='coerce'
                        )
                        if melted_df['Date'].isna().all():
                            errors.append("Error: Unable to parse dates. Ensure 'Date' column contains valid date strings or Excel serial dates.")
                    except ValueError as e:
                        errors.append(f"Error parsing dates: {str(e)}. Ensure valid date formats.")
                    
                    # Construct Timestamp
                    # Handle 0.0 as 00:00 of the same day
                    melted_df['Timestamp'] = (
                        melted_df['Date'] + pd.to_timedelta(
                            melted_df['Time_Fraction'].apply(lambda x: 1.0 if x == 0.0 else x), unit='D')
                    ).dt.strftime("%d/%m/%y %H:%M")
                    
                    melted_df = melted_df.drop(columns=["Date", "Time_Fraction"])
                    melted_df = melted_df.dropna(subset=["Timestamp", "Value"])
                    
                    # Validate converted data
                    if melted_df.empty:
                        errors.append("Warning: No valid data after conversion. Check your input file.")
                except Exception as e:
                    errors.append(f"Error during data processing: {str(e)}.")
            
            # Display errors if any
            if errors:
                st.error("The following issues were encountered:")
                for error in errors:
                    st.write(error)
            
            # Display the converted data if no critical errors
            if not errors or all("Warning" in e for e in errors):
                st.write("Converted Data Preview:", melted_df)
                
                # Offer download as XLSX using BytesIO
                try:
                    output = io.BytesIO()
                    melted_df.to_excel(output, index=False, engine='openpyxl')
                    output.seek(0)
                    st.download_button(
                        label="Download Processed File as XLSX",
                        data=output,
                        file_name="processed_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    errors.append(f"Error generating XLSX file: {str(e)}.")

    except Exception as e:
        errors.append(f"Unexpected error reading file: {str(e)}.")
        st.error("The following issues were encountered:")
        for error in errors:
            st.write(error)
