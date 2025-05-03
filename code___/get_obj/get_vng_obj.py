from openpyxl.worksheet import worksheet
from code___.get_obj.common_def import open_sheet, create_classes


def get_vng_obj() -> list:
    sheet_ranges_BUR = open_sheet(filename = '../../daily_raports/vng.xlsx', sheetname = 'SMR_ бурения')
    sheet_ranges_ZBS = open_sheet(filename = '../../daily_raports/vng.xlsx', sheetname = 'SMR_ЗБС')
    dzo = 'ПАО "ННК-Варьеганнефтегаз'
    wellname_list = get_vng_data(sheet_ranges_BUR)[0] + get_vng_data(sheet_ranges_ZBS)[0]
    padname_list = get_vng_data(sheet_ranges_BUR)[1] + get_vng_data(sheet_ranges_ZBS)[1]
    fieldname_list = get_vng_data(sheet_ranges_BUR)[2] + get_vng_data(sheet_ranges_ZBS)[2]
    orn_obj = create_classes(dzo, fieldname_list, padname_list, wellname_list)

    return orn_obj

def get_vng_data(sheet_ranges: worksheet) -> tuple[list, list, list]:
    wellname_list = []
    padname_list = []
    fieldname_list = []

    for row in sheet_ranges.iter_rows(min_row = 4, max_row = 50, min_col = 1, max_col = 4):
        b_value = row[1].value
        c_value = row[2].value
        d_value = row[3].value

        if str(row[0].value).startswith('Выполнил'):
            break

        if row[2].fill.start_color.rgb in ['FFFFFF00', 'FFFFC000']:
            continue

        if b_value is not None:
            fieldname_list.append(b_value)
        if b_value is None:
            fieldname_list.append(fieldname_list[-1])

        if c_value is not None:
            padname_list.append(c_value)

        if d_value is not None:
            wellname_list.append(d_value)

    return wellname_list, padname_list, fieldname_list
