from openpyxl.worksheet import worksheet
from code___.common_def import open_sheet, create_classes


def get_orn_obj() -> list:
    sheet_ranges = open_sheet(filename = '../daily_raports/orn.xlsx', sheetname = 'бурение')
    dzo = 'ООО "ННК-Оренбургнефтегаз'
    wellname_list = get_orn_wellname(sheet_ranges)
    fieldname_list = get_orn_fieldname(sheet_ranges)
    padname_list = ['-'] * len(wellname_list)
    orn_obj = create_classes(dzo, fieldname_list, padname_list, wellname_list)

    return orn_obj

def get_orn_wellname(sheet_ranges: worksheet) -> list:
    column_b_value_list = []
    wellname_list = []

    for i in range(6, 250):
        cell_name = 'b' + str(i)
        value = sheet_ranges[cell_name].value
        if value is not None:
            column_b_value_list.append(value)

    for i in range(0, len(column_b_value_list), 3):

        wellname_list.append(column_b_value_list[i])
    return wellname_list

def get_orn_fieldname(sheet_ranges: worksheet) -> list:
    fieldname_list = []
    for i in range(6, 250):
        cell_name = 'a' + str(i)
        value = sheet_ranges[cell_name].value
        if value is not None:
            fieldname_list.append(value)
    return fieldname_list


print(get_orn_obj())