import gspread
from oauth2client.service_account import ServiceAccountCredentials
import warnings

def write_fast_gsheet(data, sheetname="Sheet1",
                             scope='https://spreadsheets.google.com/feeds',
                             json_filename=r'/data/MatlabCode/PBLabToolkit/CalciumDataAnalysis/client_secret.json',
                             spreadsheet_name=r"FAST Multiscaler Data Registry", row=2):
    """
    Write the data dict into a Google sheet
    :return: None
    """
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_filename, scope)
    client = gspread.authorize(creds)
    try:  # if the worksheet already exists
        sheet = client.open(spreadsheet_name).add_worksheet(sheetname, rows=1, cols=30)
    except:
        sheet = client.open(spreadsheet_name).worksheet(sheetname)

    col_heads = ['filename', 'frequency', 'tot_strict', 'length_strict', 'pct_strict', 'tot_loose', 'length_loose', 'pct_loose', 'theo_delay', 'meas_delay']
    for key, item in data.items():
        try:
            col_idx = sheet.find(key)
        except CellNotFound:
            warnings.warn(f"Key {key} was missing.")
        sheet.update_cell(row, col_idx.col, item)


def write_dict_to_gsheet(data, sheet, row: str):
    """
    Write the data dict into a Google sheet
    :return: None
    """
    headers = sheet.row_values(1)
    if [] == headers:
        raise UserWarning("Missing header row")

    try:
        row_idx = sheet.find(row)
    except CellNotFound:
        raise UserWarning(f"Missing row key {row}.")

    print(data)
    for key, item in data.items():
        try:
            col_idx = sheet.find(key)
        except CellNotFound:
            warnings.warn(f"Key {key} was missing.")
        sheet.append_row(item)