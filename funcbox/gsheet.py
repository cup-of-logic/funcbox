from typing import List, Dict

import numpy as np
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials


SH_RETURN_TYPES = ('list_of_lists', 'list_of_dicts', 'array', 'dataframe')


def get_worksheet(creds_filepath: str, sheet_id: str, worksheet_index: int = 0, worksheet_name: str = None):

    creds = Credentials.from_service_account_file(
        filename=creds_filepath,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )

    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id)

    if worksheet_name:
        worksheet = sheet.worksheet(worksheet_name)
    else:
        worksheet = sheet.get_worksheet(worksheet_index)

    return worksheet


def format_worksheet(worksheet, return_type: str = 'list_of_lists'):
    if return_type == 'list_of_lists':
        return worksheet.get_all_values()
    elif return_type == 'list_of_dicts':
        return worksheet.get_all_records()
    elif return_type == 'array':
        return np.array(worksheet.get_all_values())
    elif return_type == 'dataframe':
        sheet_list = format_worksheet(worksheet)
        return pd.DataFrame(sheet_list[1:], columns=sheet_list[0])
    else:
        raise ValueError(f"Return type '{return_type}' is not supported. Supported return types: {SH_RETURN_TYPES}")


def update_entire_worksheet(worksheet, data):
    if isinstance(data, list) and all(isinstance(datum, list) for datum in data):
        n_rows, n_cols = len(data), len(data[0])
        worksheet.update(f"A1:{chr(ord('A') + n_cols - 1)}{n_rows}", data)

    elif isinstance(data, list) and all(isinstance(datum, dict) for datum in data):
        header = set().union(*(d.keys() for d in data))
        update_entire_worksheet(worksheet, data=[list(header)] + [[d.get(key, None) for key in header] for d in data])

    elif isinstance(data, pd.DataFrame):
        update_entire_worksheet(worksheet, data=[data.columns.tolist()] + data.values.tolist())

    elif isinstance(data, np.ndarray):
        update_entire_worksheet(worksheet, data=data.tolist())

    else:
        raise ValueError(f"Datatype '{type(data)}' for data is not supported. Supported datatypes: {SH_RETURN_TYPES}")

