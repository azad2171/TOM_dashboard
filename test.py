import streamlit as st
import time
import os
import xlwings as xw
import pandas as pd

def get_summary_from_sheet(sheet):
    # Implement your logic to extract data from the sheet and return a Pandas DataFrame
    # For example, assuming the sheet is a list of lists:
    df = pd.DataFrame(sheet, columns=["Column1", "Column2", "Column3"])
    return df

# Path to your Excel workbook
excel_path = 'Autospreader_monitor_non_macro.xlsx'

# Connect to the Excel workbook
wb = xw.Book(excel_path)


products = ['BRN', 'BOX', 'HO', 'RB', 'ZL', 'ZS', 'ZM', 'ZW', 'ZC', 'NGHH']

# Get the initial modification time
last_modified_time = os.path.getmtime(excel_path)

# Set up the Streamlit app
st.title("Test")

try:
    while True:
        current_modified_time = os.path.getmtime(excel_path)
        
        # Check if the workbook has been updated
        if current_modified_time > last_modified_time:
            main_out_df = pd.DataFrame()
            
            
            for sheet in wb.sheets:
                if sheet.name in products:
                    product_sheet = sheet.used_range.value
                    df = get_summary_from_sheet(product_sheet)
                    st.dataframe(df)
            
            # Update the last modification time
            last_modified_time = current_modified_time
        
        # Sleep for 5 seconds before the next check
        time.sleep(1)

except KeyboardInterrupt:
    # Close the workbook when the Streamlit app is closed
    # wb.close()
    st.write("Excel workbook closed.")
