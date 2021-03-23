import openpyxl
import re

def value_get(base_sheet, search_sheet, id, search_str):
    for row_data in base_sheet.iter_rows(min_row=4):
        if str(row_data[2].value) == str(id):
            base_value = row_data[43].value
            break

    for row_data in search_sheet.iter_rows(min_row=2):
        for cell in row_data:
            if cell.value == None:
                continue
            elif re.search(r'2021-01-01',str(cell.value)):
                start_col = cell.column
            elif search_str == cell.value:
                end_col = cell.column
                break
        break

    for row_data in search_sheet.iter_rows(min_row=3):
        if str(row_data[1].value) == str(id):
            flag = False
            dec = 0
            for cell in row_data:

                if cell.column == start_col:
                    flag = True
                if flag == True:
                    dec += int(cell.value)
                    if cell.column == end_col:
                        break

    result = int(base_value) + dec
    return result


def main():
    target_file = 'FILENAME.xlsx'
    result_file = 'RESULT.xlsx'

    base_value_sheet = 'BASE'
    decrement_value_sheet  = 'SEARCH'

    search_wb = openpyxl.load_workbook(result_file)
    search_wb.save(result_file)

    sheet_list = search_wb.sheetnames
    ws = search_wb[sheet_list[0]]
    search_date = ws['C2'].value

    wb = openpyxl.load_workbook(target_file)

    for row_data in ws.iter_rows(min_row=5,min_col=2):
        row_data[1].value = value_get(wb[base_value_sheet], wb[decrement_value_sheet], row_data[0].value, search_date)


    search_wb.save(result_file)

if __name__ == '__main__':
    main()
