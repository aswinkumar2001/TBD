import streamlit as st
import pandas as pd
import io
import numpy as np
from datetime import timedelta

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
            # Define standard time fractions (48 half-hour increments, ending with 0.0)
            standard_fractions = [i / 48 for i in range(1, 48)] + [0.0]  # 0.020833333 to 0.979166667, then 0.0
            # Map fractions to HH:MM times
            time_mappings = {}
            for i in range(1, 48):
                hours = (i * 30) // 60
                minutes = (i * 30) % 60
                time_mappings[i / 48] = f"{hours:02d}:{minutes:02d}"
            time_mappings[0.0] = "00:00"
            
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
                expected_set = set(np.round(standard_fractions, 8))
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
                    
                    # Map Time_Fraction to nearest standard fraction
                    def map_to_nearest_fraction(fraction):
                        if pd.isna(fraction):
                            return fraction
                        # Find the standard fraction with minimum absolute difference
                        differences = [abs(fraction - sf) for sf in standard_fractions]
                        min_index = np.argmin(differences)
                        return standard_fractions[min_index]
                    
                    melted_df['Mapped_Fraction'] = melted_df['Time_Fraction'].apply(map_to_nearest_fraction)
                    
                    # Handle Date column flexibly (DD/MM/YY or Excel serial date)
                    try:
                        # Try parsing as DD/MM/YY (e.g., "13/03/25")
                        melted_df['Date'] = pd.to_datetime(melted_df['Date'], 
                                                           format="%d/%m/%Y", 
                                                           errors='coerce')
                        # If still NaN, try Excel serial date (e.g., 45841)
                        mask = melted_df['Date'].isna()
                        melted_df.loc[mask, 'Date'] = pd.to_datetime(
                            pd.to_numeric(melted_df.loc[mask, 'Date'], errors='coerce'), 
                            unit='D', 
                            origin='1899-12-30', 
                            errors='coerce'
                        )
                        if melted_df['Date'].isna().all():
                            errors.append("Error: Unable to parse dates. Ensure 'Date' column contains valid DD/MM/YY strings or Excel serial dates.")
                    except ValueError as e:
                        errors.append(f"Error parsing dates: {str(e)}. Ensure valid DD/MM/YY or serial date formats.")
                    
                    # Construct Timestamp using mapped fractions with exact HH:MM
                    melted_df['Timestamp'] = melted_df.apply(
                        lambda row: pd.NA if pd.isna(row['Date']) or pd.isna(row['Mapped_Fraction']) 
                        else row['Date'].strftime("%d/%m/%y ") + time_mappings[row['Mapped_Fraction']],
                        axis=1
                    )
                    
                    melted_df = melted_df.drop(columns=["Date", "Time_Fraction", "Mapped_Fraction"])
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
                
                # Offer download as XLSX using BytesIO, forcing Timestamp column as text
                try:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        melted_df.to_excel(writer, index=False, sheet_name='Sheet1')
                        worksheet = writer.sheets['Sheet1']
                        # Set Timestamp column (column A, index 1) to text format '@'
                        for row in range(2, worksheet.max_row + 1):  # Data rows start at 2
                            worksheet.cell(row=row, column=1).number_format = '@'
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
