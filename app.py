import pandas as pd
import streamlit as st
from datetime import datetime
import xlrd
import openpyxl
import io

# --- UI Header ---
st.set_page_config(page_title="Emp Termination Formatter", layout="centered")
st.title("Employee Termination Excel Formatter")
st.markdown("Upload your `.xls` file to extract and format 'Emp Term- SAP Accounts' rows.")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload Excel (.xls) file", type=["xls"])

if uploaded_file:
    try:
        # Convert XLS to XLSX in memory
        xls_book = xlrd.open_workbook(file_contents=uploaded_file.read())
        sheet = xls_book.sheet_by_index(0)

        xlsx_io = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active

        for row_idx in range(sheet.nrows):
            for col_idx in range(sheet.ncols):
                ws.cell(row=row_idx + 1, column=col_idx + 1, value=sheet.cell_value(row_idx, col_idx))

        wb.save(xlsx_io)
        xlsx_io.seek(0)

        # Load as DataFrame
        df = pd.read_excel(xlsx_io, header=None)

        # Column mapping
        headers = {
            'HD ID': 'A',
            'Task ID': 'C',
            'Task Desc': 'E',
            'Task Tech': 'J',
            'Task Create': 'L',
            'Task Status': 'N',
            'Task Comp Date': 'Q',
            'Task Group': 'U'
        }

        def col_letter_to_index(letter):
            return openpyxl.utils.column_index_from_string(letter) - 1

        output_columns = list(headers.keys())
        data = []
        last_known = {}

        for i in range(7, df.shape[0]):
            row_data = {}

            for col_name in output_columns:
                col_idx = col_letter_to_index(headers[col_name])
                value = df.iloc[i, col_idx] if col_idx < df.shape[1] else ""

                if col_name in ['Task Create', 'Task Comp Date']:
                    if isinstance(value, (int, float)):
                        try:
                            value = datetime(1899, 12, 30) + pd.to_timedelta(value, unit='d')
                            value = value.strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            pass

                if col_name == 'HD ID':
                    if pd.notna(value) and str(value).strip() != '':
                        last_known['HD ID'] = value
                    else:
                        value = last_known.get('HD ID', '')

                row_data[col_name] = value

            if str(row_data['Task Desc']).strip() == "Emp Term- SAP Accounts":
                data.append(row_data)

        # Show and offer download
        if data:
            result_df = pd.DataFrame(data, columns=output_columns)
            st.success("Processed successfully!")
            st.dataframe(result_df)

            output_io = io.BytesIO()
            result_df.to_excel(output_io, index=False)
            output_io.seek(0)

            today_str = datetime.today().strftime('%Y-%m-%d')
            filename = f"Emp_Termination_{today_str}.xlsx"

            st.download_button(
                label="Download Excel File",
                data=output_io,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No matching rows found with 'Emp Term- SAP Accounts'.")

    except Exception as e:
        st.error(f"Error: {str(e)}")
