import streamlit as st
import pandas as pd
import io

# Set page configuration
st.title("Duplicate Fed To Analyzer")
st.write("Upload an Excel file with columns 'Fed To' and 'Fed From'. The app will extract rows where 'Fed To' has duplicates and provide a download option for the resulting Excel.")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        # Ensure columns exist
        if 'Fed To' not in df.columns or 'Fed From' not in df.columns:
            st.error("The uploaded file must contain 'Fed To' and 'Fed From' columns.")
        else:
            # Handle empty DataFrame
            if df.empty:
                st.warning("The uploaded file is empty. No data to process.")
            else:
                # Find duplicate 'Fed To'
                duplicates = df['Fed To'].value_counts()
                duplicate_values = duplicates[duplicates > 1].index
                
                # Filter the DataFrame to include only rows with duplicate 'Fed To'
                result_df = df[df['Fed To'].isin(duplicate_values)].sort_values(by='Fed To')
                
                # Display preview
                st.write("Preview of Resulting Data (Duplicates in 'Fed To'):")
                st.dataframe(result_df.head(10))  # Show first 10 rows
                
                if result_df.empty:
                    st.info("No duplicates found in 'Fed To'.")
                
                # Function to convert DataFrame to Excel
                def to_excel(df):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Duplicates')
                    return output.getvalue()
                
                # Download button
                excel_data = to_excel(result_df)
                st.download_button(
                    label="Download Resulting Excel",
                    data=excel_data,
                    file_name="duplicates_fed_to.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"An error occurred while processing the file: {str(e)}")