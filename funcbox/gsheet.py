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


class Workbook:
    def __init__(self, creds_filepath: str, workbook_id: str):
        self.workbook_id = workbook_id
        self.creds = Credentials.from_service_account_file(
            filename=creds_filepath,
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )

        self.client = gspread.authorize(self.creds)
        self.workbook = self.client.open_by_key(self.workbook_id)
        self.worksheets = self.workbook.worksheets()
        self.worksheet = self.worksheets[0]
        self.__truncate_empty_worksheets()

    def create_worksheet(self, name: str, headers: list, data=None, select: bool = True):
        if data:
            data = get_worksheet_data_as_list(data)
        self.workbook.add_worksheet(name, rows=1 + (len(data) if data else 0), cols=len(headers))
        worksheet = self.workbook.worksheet(name)
        self.worksheets.append(worksheet)
        worksheet.update(range_name="A1", values=[headers] + (data if data else []))

        if select:
            self.worksheet = worksheet

    def remove_worksheet(self, select=0):
        self.workbook.del_worksheet(self.worksheet)
        self.worksheet = self.select_worksheet(select)

    def select_worksheet(self, identifier):
        if isinstance(identifier, int):
            self.worksheet = self.workbook.get_worksheet(identifier)
        elif isinstance(identifier, str):
            self.worksheet = self.workbook.worksheet(identifier)
        else:
            raise ValueError(f"Datatype for key in __getitem__ cannot be '{type(identifier)}'. Supported datatypes: (str, int)")

    def get_worksheet_data(self, return_type: str = 'list_of_lists'):
        if return_type == 'list_of_lists':
            return self.worksheet.get_all_values()
        elif return_type == 'list_of_dicts':
            return self.worksheet.get_all_records()
        elif return_type == 'array':
            return np.array(self.worksheet.get_all_values())
        elif return_type == 'dataframe':
            sheet_list = self.get_worksheet_data(self.worksheet)
            return pd.DataFrame(sheet_list[1:], columns=sheet_list[0])
        else:
            raise ValueError(
                f"Return type '{return_type}' is not supported. Supported return types: {SHEET_RETURN_TYPES}")

    def set_worksheet(self, data):
        data = get_worksheet_data_as_list(data)
        n_rows, n_cols = len(data), len(data[0])

        self.worksheet.resize(rows=n_rows, cols=n_cols)
        self.worksheet.update("A1", data)

    def append_rows(self, row_data):
        row_data = get_worksheet_data_as_list(row_data)
        n_rows = self.worksheet.row_count
        n_cols = self.worksheet.col_count
        self.worksheet.resize(rows=n_rows + len(row_data), cols=n_cols)
        self.worksheet.update(f'A{n_rows + 1}', row_data)

    def append_cols(self, col_data, header: bool = True):
        col_data = get_worksheet_data_as_list(col_data)
        col_data = [[row[i] for row in col_data] for i in range(len(col_data[0]))]
        n_rows = self.worksheet.row_count
        n_cols = self.worksheet.col_count
        self.worksheet.resize(rows=n_rows, cols=n_cols + len(col_data))
        self.worksheet.update(f'{get_column_name(n_cols + 1)}{2 if header else 1}', col_data)

    def __truncate_empty_worksheets(self):
        for worksheet in self.worksheets:
            if worksheet.acell('A1').value is None:
                worksheet.resize(rows=1, cols=1)

    def __str__(self):
        return f'<Workbook "{self.workbook.title}" total_worksheets={len(self.worksheets)}>'

    def __repr__(self):
        return self.__str__()

    def __getattr__(self, item):
        return getattr(self.workbook, item)
