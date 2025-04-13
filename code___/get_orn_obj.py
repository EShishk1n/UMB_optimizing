from openpyxl.worksheet import worksheet
from code___.common_def import open_sheet, create_classes


def get_orn_obj() -> list:
    sheet_ranges = open_sheet(filename = '../daily_raports/orn.xlsx', sheetname = 'бурение')
    dzo = 'ООО "ННК-Оренбургнефтегаз'
    wellname_list = get_orn_data(sheet_ranges)[0][::3]
    fieldname_list = get_orn_data(sheet_ranges)[1]
    padname_list = ['-'] * len(wellname_list)
    orn_obj = create_classes(dzo, fieldname_list, padname_list, wellname_list)

    return orn_obj

def get_orn_data(sheet_ranges: worksheet) -> tuple[list, list]:
    wellname_list = []
    fieldname_list = []

    for row in sheet_ranges.iter_rows(min_row = 6, max_row = 250, min_col = 1, max_col = 2):
        a_value = row[0].value
        b_value = row[1].value

        if a_value is not None:
            fieldname_list.append(a_value)

        if b_value is not None:
            wellname_list.append(b_value)

    return wellname_list, fieldname_list


print(get_orn_obj())