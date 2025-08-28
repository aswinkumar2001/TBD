import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

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
            # Define expected time fractions (48 half-hour increments)
            time_fractions = [i / 48 for i in range(1, 49)] + [0.0]  # 0.020833333 to 0.979166667, then 0.0
            column_headers = df.columns[1:]  # Exclude "Date" column
            if not all(col in column_headers for col in time_fractions):
                errors.append("Error: Mismatch in expected time fraction columns. Expected 48 fractions plus 0.0.")
            
            # Process the data
            try:
                # Melt the dataframe
                melted_df = pd.melt(df, 
                                  id_vars=["Date"], 
                                  var_name="Time_Fraction", 
                                  value_name="Value")
                
                # Handle Date column flexibly (text or Excel date format)
                try:
                    # Try parsing as text date first (e.g., "Thursday, March 27, 2025")
                    melted_df['Date'] = pd.to_datetime(melted_df['Date'], 
                                                     format="%A, %B %d, %Y", 
                                                     errors='coerce')
                    # If still NaN, try Excel serial date (e.g., 44962)
                    mask = melted_df['Date'].isna()
                    melted_df.loc[mask, 'Date'] = pd.to_datetime(melted_df.loc[mask, 'Date'], 
                                                               unit='D', 
                                                               origin='1899-12-30', 
                                                               errors='coerce')
                    if melted_df['Date'].isna().all():
                        errors.append("Error: Unable to parse dates. Ensure 'Date' column contains valid date strings or Excel serial dates.")
                except ValueError as e:
                    errors.append(f"Error parsing dates: {str(e)}. Ensure valid date formats.")
                
                # Map time fractions to time strings
                time_mapping = {frac: (datetime(1900, 1, 1) + timedelta(days=frac)).strftime("%H:%M") 
                              for frac in time_fractions}
                melted_df['Time'] = melted_df['Time_Fraction'].map(time_mapping)
                
                # Construct Timestamp
                melted_df['Timestamp'] = melted_df.apply(
                    lambda row: row['Date'].strftime("%d/%m/%y") + " " + row['Time'] if pd.notna(row['Date']) else None,
                    axis=1
                )
                melted_df = melted_df.drop(columns=["Date", "Time_Fraction", "Time"])
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
                
                # Offer download as XLSX
                try:
                    output = melted_df.to_excel("processed_data.xlsx", index=False, engine='openpyxl')
                    with open("processed_data.xlsx", "rb") as file:
                        st.download_button(
                            label="Download Processed File as XLSX",
                            data=file,
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
