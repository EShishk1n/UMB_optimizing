from openpyxl.worksheet import worksheet
from code___.common_def import open_sheet, create_classes


def get_nng_obj() -> list:
    sheet_ranges = open_sheet(filename = '../daily_raports/nng.xlsx', sheetname = 'ЭБ и ЗБС ')
    dzo = 'ООО "ННК-Няганьнефтегаз'
    wellname_list = get_nng_wellname(sheet_ranges)
    fieldname_list = ['Красноленинское'] * len(wellname_list)
    padname_list = get_nng_padname(sheet_ranges)
    nng_obj = create_classes(dzo, fieldname_list, padname_list, wellname_list)

    return nng_obj

def get_nng_wellname(sheet_ranges: worksheet) -> list:
    column_c_value_list = []
    wellname_list = []

    for i in range(7, 350):
        cell_name = 'c' + str(i)
        value = sheet_ranges[cell_name].value
        if value is not None:
            column_c_value_list.append(value)

    for i in range(0, len(column_c_value_list), 2):
        value_to_append = column_c_value_list[i].split('\n')[0]
        value_to_append = value_to_append.split(' ')[0]
        wellname_list.append(value_to_append)
    return wellname_list

def get_nng_padname(sheet_ranges: worksheet) -> list:
    padname_list = []
    for i in range(7, 350):
        cell_name = 'b' + str(i)
        value = sheet_ranges[cell_name].value
        if value is not None:
            value_to_append = value.split('\n')[0]
            value_to_append = value_to_append.split(' ')[0]
            padname_list.append(value_to_append)
    return padname_list


print(get_nng_obj())