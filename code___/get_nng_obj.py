from openpyxl.worksheet import worksheet
from code___.common_def import open_sheet, create_classes


def get_nng_obj() -> list:
    sheet_ranges = open_sheet(filename = '../daily_raports/nng.xlsx', sheetname = 'ЭБ и ЗБС ')
    dzo = 'ООО "ННК-Няганьнефтегаз'
    wellname_list = get_nng_data(sheet_ranges)[0][::2]
    fieldname_list = ['Красноленинское'] * len(wellname_list)
    padname_list = get_nng_data(sheet_ranges)[1]
    nng_obj = create_classes(dzo, fieldname_list, padname_list, wellname_list)

    return nng_obj


def get_nng_data(sheet_ranges: worksheet) -> tuple[list, list]:
    wellname_list = []
    padname_list = []

    for row in sheet_ranges.iter_rows(min_row = 7, max_row = 350, min_col = 2, max_col = 3):
        b_value = row[0].value
        c_value = row[1].value

        if b_value is not None:
            value_to_append = b_value.split('\n')[0].split(' ')[0]
            padname_list.append(value_to_append)

        if c_value is not None:
            value_to_append = c_value.split('\n')[0].split(' ')[0]
            wellname_list.append(value_to_append)

    return wellname_list, padname_list

print(get_nng_obj())