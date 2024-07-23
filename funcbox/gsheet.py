import numpy as np
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials


SHEET_RETURN_TYPES = ('list_of_lists', 'list_of_dicts', 'array', 'dataframe')


def get_column_name(index):
    column_name = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        column_name = chr(65 + remainder) + column_name
    return column_name


def get_worksheet_data_as_list(worksheet_data):
    if isinstance(worksheet_data, list) and all(isinstance(datum, list) for datum in worksheet_data):
        return worksheet_data

    elif isinstance(worksheet_data, list) and all(isinstance(datum, dict) for datum in worksheet_data):
        header = set().union(*(d.keys() for d in worksheet_data))
        return [list(header)] + [[d.get(key, None) for key in header] for d in worksheet_data]

    elif isinstance(worksheet_data, pd.DataFrame):
        return [worksheet_data.columns.tolist()] + worksheet_data.values.tolist()

    elif isinstance(worksheet_data, np.ndarray):
        return worksheet_data.tolist()

    else:
        raise ValueError(
            f"Datatype '{type(worksheet_data)}' for data is not supported. Supported datatypes: {SHEET_RETURN_TYPES}")


def get_worksheet_as(worksheet, return_type: str = 'list_of_lists'):
    if return_type == 'list_of_lists':
        return worksheet.get_all_values()
    elif return_type == 'list_of_dicts':
        return worksheet.get_all_records()
    elif return_type == 'array':
        return np.array(worksheet.get_all_values())
    elif return_type == 'dataframe':
        sheet_list = get_worksheet_as(worksheet)
        return pd.DataFrame(sheet_list[1:], columns=sheet_list[0])
    else:
        raise ValueError(f"Return type '{return_type}' is not supported. Supported return types: {SHEET_RETURN_TYPES}")


def append_rows(worksheet, row_data):
    row_data = get_worksheet_data_as_list(row_data)
    n_rows = worksheet.row_count
    worksheet.update(f'A{n_rows + 1}', row_data)


def append_cols(worksheet, col_data, header: bool = True):
    col_data = get_worksheet_data_as_list(col_data)
    col_data = [[row[i] for row in col_data] for i in range(len(col_data[0]))]
    n_cols = worksheet.col_count
    worksheet.update(f'{get_column_name(n_cols + 1)}{2 if header else 1}', col_data)


class Workbook:
    def __init__(self, creds_filepath: str, workbook_id: str):
        self.workbook_id = workbook_id
        self.creds = Credentials.from_service_account_file(
            filename=creds_filepath,
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )

        self.client = gspread.authorize(self.creds)
        self.workbook = self.client.open_by_key(self.workbook_id)

    def __getitem__(self, key):
        if isinstance(key, int):
            return self.workbook.get_worksheet(key)
        elif isinstance(key, str):
            return self.workbook.worksheet(key)
        else:
            raise ValueError(f"Datatype for key in __getitem__ cannot be '{type(key)}'. Supported datatypes: (str, int)")

    def __setitem__(self, key, data):
        worksheet = self.__getitem__(key)
        data = get_worksheet_data_as_list(data)

        n_rows, n_cols = len(data), len(data[0])
        worksheet.clear()
        worksheet.update(f"A1:{chr(ord('A') + n_cols - 1)}{n_rows}", data)


if __name__ == '__main__':
    print(get_column_name(703))

