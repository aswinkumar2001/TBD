import streamlit as st
import pandas as pd
from datetime import datetime

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
            # Validate time columns (0:30 to 23:30, 0:00)
            expected_time_columns = [f"{h:02d}:{m:02d}" for h in range(24) for m in (0, 30)] + ["0:00"]
            missing_columns = [col for col in expected_time_columns if col not in df.columns]
            if missing_columns:
                errors.append(f"Error: Missing expected time columns: {', '.join(missing_columns)}")
            
            # Process the data
            try:
                # Melt the dataframe
                melted_df = pd.melt(df, 
                                  id_vars=["Date"], 
                                  var_name="Time", 
                                  value_name="Value")
                
                # Convert Date to datetime and combine with Time
                melted_df['Date'] = pd.to_datetime(melted_df['Date'], errors='coerce')
                if melted_df['Date'].isna().all():
                    errors.append("Error: Unable to parse dates. Ensure 'Date' column contains valid date strings.")
                else:
                    melted_df['Timestamp'] = melted_df.apply(
                        lambda row: row['Date'].strftime("%d/%m/%y") + " " + row['Time'] if pd.notna(row['Date']) else None,
                        axis=1
                    )
                    melted_df = melted_df.drop(columns=["Date", "Time"])
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
