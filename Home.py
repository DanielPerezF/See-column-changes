import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


# Function to convert dataframe to Excel in memory
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Access the workbook and worksheet objects
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Define a format for the colored cells (set your preferred color)
    cell_format = workbook.add_format({'bg_color': '#B8CCE4'})#, 'font_color': '#9C0006'})
    # Apply formats
    for i in range(1,old.columns.size+1):
        worksheet.write(0, i, df.columns[i], cell_format)  # First cell in first row

    # Columns selected part of new dataframe
    cell_format = workbook.add_format({'bg_color': '#FCD5B4'})#, 'font_color': '#9C0006'})
    # Apply formats
    for i in range(old.columns.size+1, old.columns.size+1+len(cols)):
        worksheet.write(0, i, df.columns[i], cell_format)  # First cell in first row

    # Autoscale the column widths
    for idx, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2  # Add some padding
        worksheet.set_column(idx, idx, max_len)
    
    # Conditional formatting for the last n columns (TRUE values highlighted)
    start_col_index = len(df.columns) - len(cols)  # Starting index for the last n columns
    for col_idx in range(start_col_index, len(df.columns)):
        worksheet.conditional_format(1, col_idx, len(df), col_idx,  # From row 1 to the last row in this column
                                     {'type': 'cell',
                                      'criteria': '==',
                                      'value': True,  # Condition: TRUE
                                      'format': workbook.add_format({'bg_color': '#FFCCCC', 'font_color': '#9C0006'})  # Green highlight
                                     })

    writer.close()  # Use close() instead of save()
    processed_data = output.getvalue()
    return processed_data




st.set_page_config(layout="wide")
st.title('Compare Excel files')
st.write('Compare two Excel files and highlights the changes between them. Please upload the two files below.')
st.write('Then select the column to be used as unique ID for each row')
st.write('Finally, select the columns to be compared for changes')
st.write('\nDeveloped by Daniel Perez')

st.subheader('Upload files')

col1, col2 = st.columns(2)
old_file = col1.file_uploader("Old file")
new_file = col2.file_uploader("New file")

if old_file is None or new_file is None:
    st.warning('Please upload both files')
    st.stop()
else:
    old = pd.read_excel(old_file)
    new = pd.read_excel(new_file)
    id_var = st.selectbox('Choose the key variable', old.columns)
    old = old.set_index(id_var)
    new = new.set_index(id_var)
    cols = st.multiselect('Columns to check for changes', old.columns, default=old.columns[0])


    merged = old.merge(new[cols], on=id_var, how='outer', suffixes=('_old', '_new'))
    for col in cols:
        changed = merged[col+'_old'] != merged[col+'_new']
        merged[col+'_new'] = merged[col+'_new'].where(changed, np.nan)
        merged[col+'_changed'] = changed

    merged = merged.sort_values(cols[0]+'_old')
    with st.expander('Show changes'):
        st.write(merged)
    #st.write(merged)


    # Call the function to create the Excel file in memory
    excel_data = to_excel(merged.reset_index())

    # Provide the download button in the Streamlit app
    st.download_button(
        label="Download Excel file",
        data=excel_data,
        file_name='Changes.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
