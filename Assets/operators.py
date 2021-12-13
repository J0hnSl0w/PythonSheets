import json
import gspread

from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials


def open_spread_sheet(spreadsheet_name: str, credentials_file: str):
    scope = ['https://www.googleapis.com/auth/drive',
             'https://www.googleapis.com/auth/drive.file']

    file_name = credentials_file
    creds = ServiceAccountCredentials.from_json_keyfile_name(file_name, scope)
    client = gspread.authorize(creds)

    spreadsheet = client.open(spreadsheet_name)
    print(f'Táblázat megnyitva: {spreadsheet.title}')

    return spreadsheet, creds


def load_json(json_file):
    f = open(json_file, 'r')
    dict = json.load(f)

    return dict


def calc_date(sheets_dict):
    date_dict = {'prev_month': '',
                 'this_month': ''}

    for key in sheets_dict['month']:
        date_dict['prev_month'] = key
        splitted = key.split('.')
        splitted_year = int(splitted[0])
        splitted_month = int(splitted[1])

        if splitted_month < 9:
            date_dict['this_month'] = splitted[0] + '.0' + str(splitted_month + 1) + '.'
        elif splitted_month == 12:
            date_dict['this_month'] = str(splitted_year + 1) + '.0' + str(splitted_month - 11) + '.'
        else:
            date_dict['this_month'] = splitted[0] + '.' + str(splitted_month + 1) + '.'
        break

    tm = date_dict['this_month']
    iftm = find_sheet_by_name(tm, sheets_dict['month'])
    if len(iftm) != 0:
        exeption = 'Kérlek, rendezd a napi kiadásokat tartalmazó lapokat balról jobbra növekvő sorrendbe, ' \
                   'indítsd el a programot és válaszd az 5-ös menüpontot!'
        raise Exception(exeption)
    return date_dict


def find_sheet_by_name(name: str, sheet_dict: dict):
    sheet = []

    for key in sheet_dict:
        if str(key) == name:
            sheet.append(name)
            sheet.append(sheet_dict[key][0])
            sheet.append(sheet_dict[key][1])
            break

    return sheet


def repeat_formula_over_range(creds, wbook_id, sheet_id, formula, start_row, start_col, end_row, end_col):
    service = discovery.build('sheets', 'v4', credentials=creds)

    requests = {'repeatCell': {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": start_row,
            "endRowIndex": end_row,
            "startColumnIndex": start_col,
            "endColumnIndex": end_col},
        "cell": {
            "userEnteredValue": {
                "formulaValue": formula}},
        "fields": 'userEnteredValue'}}

    body = {'requests': [requests]}

    request = service.spreadsheets().batchUpdate(spreadsheetId=wbook_id, body=body)
    response = request.execute()

    return response


def show_hide_wsheets(creds, wbook_id, sheet_id, is_hidden):
    service = discovery.build('sheets', 'v4', credentials=creds)

    requests = {'updateSheetProperties': {
        "properties": {
            "sheetId": sheet_id,
            "hidden": is_hidden},
        "fields": 'hidden'}}

    body = {'requests': [requests]}

    request = service.spreadsheets().batchUpdate(spreadsheetId=wbook_id, body=body)
    response = request.execute()

    return response


def show_hide_cols(creds, wbook_id, sheet_id, is_hidden, start, end):
    service = discovery.build('sheets', 'v4', credentials=creds)

    requests = {'updateDimensionProperties': {
        "range": {
            "sheetId": sheet_id,
            "dimension": 'COLUMNS',
            "startIndex": start,
            "endIndex": end},
        "properties": {
            "hiddenByUser": is_hidden},
        "fields": 'hiddenByUser'}}

    body = {'requests': [requests]}

    request = service.spreadsheets().batchUpdate(spreadsheetId=wbook_id, body=body)
    response = request.execute()

    return response


def insert_row(creds, wbook_id, sheet_id, start, end, inherit_from_before):
    service = discovery.build('sheets', 'v4', credentials=creds)

    requests = {'insertDimension': {
        "range": {
            "sheetId": sheet_id,
            "dimension": 'ROWS',
            "startIndex": start,
            "endIndex": end},
        "inheritFromBefore": inherit_from_before}}

    body = {'requests': [requests]}

    request = service.spreadsheets().batchUpdate(spreadsheetId=wbook_id, body=body)
    response = request.execute()

    return response