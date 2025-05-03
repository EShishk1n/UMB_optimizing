from openpyxl.worksheet import worksheet
from code___.get_obj.common_def import open_sheet, create_classes


def get_smng_obj():
    sheet_ranges = open_sheet(filename = '../../daily_raports/smng.xlsx', sheetname = 'Север')
    dzo = 'ООО "ННК-Самаранефтегаз'
    wellname_list = get_smng_data(sheet_ranges)[::8]
    fieldname_list = get_smng_data(sheet_ranges)[5::8]
    padname_list = ['-'] * len(wellname_list)
    smng_obj = create_classes(dzo, fieldname_list, padname_list, wellname_list)
    return smng_obj

def get_smng_data(sheet_ranges: worksheet) -> list:
    data_without_hidden_rows = []

    for row in sheet_ranges.iter_rows(min_row=8, max_row=250, min_col=2, max_col=2):
        row_num = row[0].row
        if str(row[0].value).startswith("ИТОГО"):
            break

        if not sheet_ranges.row_dimensions[row_num].hidden:
            data_without_hidden_rows.append(row[0].value)

    return data_without_hidden_rows
