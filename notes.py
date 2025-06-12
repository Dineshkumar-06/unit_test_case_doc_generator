"""  
function part for finding the last non empty column and pasting the test case column:

row = 1
    while row <= combined_sheet.max_row:
        found_col = None
        for col in range(1, combined_sheet.max_column + 1):
            cell_value = str(combined_sheet.cell(row=row, column=col).value).strip().lower()
            if "relaxation years" in cell_value:
                found_col = col
                break

        if found_col:
            # Find last non-empty column in this row after the found_col
            last_data_col = found_col
            for col in range(found_col, combined_sheet.max_column + 1):
                if combined_sheet.cell(row=row, column=col).value not in [None, ""]:
                    last_data_col = col

            start_col = last_data_col + 1  # Start after the last filled column

            start_row = row + 1
            end_row = start_row

            while end_row <= combined_sheet.max_row:
                if all(combined_sheet.cell(row=end_row, column=col).value in [None, ""] for col in range(1, 6)):
                    break
                end_row += 1

            add_testcase_columns_func(combined_sheet, start_row, end_row - 1, start_col)

            row = end_row
        else:
            row += 1
 """


""" 
pyinstaller --onefile --icon=icon.ico testcase_doc_generator.py

pyinstaller --onefile --windowed --name=MyTool your_script.py

pyinstaller --onefile --windowed --name=UT Document Generator --icon=icon.ico testcase_doc_generator.py 
"""