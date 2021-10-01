import gspread
import json

from googleapiclient import discovery
from enumerator_dickts import *
from datetime import date
from oauth2client.service_account import ServiceAccountCredentials


def open_spread_sheet(spreadsheet_name: str):
    scope = ['https://www.googleapis.com/auth/drive',
             'https://www.googleapis.com/auth/drive.file']

    file_name = 'credentials.json'
    creds = ServiceAccountCredentials.from_json_keyfile_name(file_name, scope)
    client = gspread.authorize(creds)

    spreadsheet = client.open(spreadsheet_name)
    print(f'Spreadsheet opened: {spreadsheet.title}')

    return spreadsheet, creds


def update_sheet_list(spreadsheet):
    ssheet = {}

    print('Updating sheet list...')

    l = len(spreadsheet.worksheets())
    k = 0

    for i in range(l):
        wsheet = spreadsheet.get_worksheet(i)
        index = i + 1

        ssheet[f'{wsheet.title}'] = [wsheet.id, index]
        percent = (i/l)*100

        if k == 9:
            print(f'done: {round(percent)}%')
            k = 0
        k += 1

    with open('sheets.jason', 'w') as file:
        json.dump(ssheet, file)

    print(f'done: 100%')


def load_json(json_file):
    f = open(json_file, 'r')
    dict = json.load(f)

    return dict


def find_sheet_by_name(name: str, sheet_dict: dict):
    sheet = []

    for key in sheet_dict:
        if key == name:
            sheet.append(name)
            sheet.append(sheet_dict[key][0])
            sheet.append(sheet_dict[key][1])

    return sheet


def calc_date():
    today = date.today()
    date_dict = {}

    if today.month == 1:
        date_dict = {'prev_month': f'{today.year - 1}.0{today.month + 11}.',
                     'this_month': f'{today.year}.0{today.month}.',
                     'next_month': f'{today.year}.0{today.month + 1}.'}
    elif today.month < 9:
        date_dict = {'prev_month': f'{today.year}.0{today.month - 1}.',
                     'this_month': f'{today.year}.0{today.month}.',
                     'next_month': f'{today.year}.0{today.month + 1}.'}
    elif today.month == 9:
        date_dict = {'prev_month': f'{today.year}.0{today.month - 1}.',
                     'this_month': f'{today.year}.0{today.month}.',
                     'next_month': f'{today.year}.{today.month + 1}.'}
    elif today.month == 10:
        date_dict = {'prev_month': f'{today.year}.0{today.month - 1}.',
                     'this_month': f'{today.year}.{today.month}.',
                     'next_month': f'{today.year}.{today.month + 1}.'}
    elif today.month == 11:
        date_dict = {'prev_month': f'{today.year}.{today.month - 1}.',
                     'this_month': f'{today.year}.{today.month}.',
                     'next_month': f'{today.year}.{today.month + 1}.'}
    elif today.month == 12:
        date_dict = {'prev_month': f'{today.year}.{today.month - 1}.',
                     'this_month': f'{today.year}.{today.month}.',
                     'next_month': f'{today.year + 1}.{today.month - 11}.'}

    return date_dict


def add_new_month(wbook, creds, is_new_month=False):
    dates = calc_date()
    template_sheet = wbook.worksheet('Template')

    if is_new_month:
        new_sheet = wbook.duplicate_sheet(source_sheet_id=template_sheet.id,
                                          insert_sheet_index=11,
                                          new_sheet_name=dates['this_month'])

        cell = new_sheet.acell('A3', 'FORMULA')
        new_value = cell.value[:2] + dates['prev_month'] + cell.value[-4:]
        new_sheet.update('A3', new_value, raw=False)

        for i in range(13):
            new_sheet.update(f'E{i+3}', [[f"{dates['this_month']}01."]], raw=False)

    else:
        new_sheet = wbook.duplicate_sheet(source_sheet_id=template_sheet.id,
                                          insert_sheet_index=11,
                                          new_sheet_name=dates['next_month'])

        cell = new_sheet.acell('A3', 'FORMULA')
        new_value = cell.value[:2] + dates['this_month'] + cell.value[-4:]
        new_sheet.update('A3', new_value, raw=False)

        for i in range(13):
            new_sheet.update(f'E{i+3}', [[f"{dates['next_month']}01."]], raw=False)

    print('New sheet added succesfully!')

    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=new_sheet.id, is_hidden=False)
    edit_mounthly_costs_sheet(wbook=wbook, creds=creds, is_new_month=is_new_month, dates=dates)


def edit_mounthly_costs_sheet(wbook, creds, is_new_month=False, dates=calc_date()):
    year = None

    if is_new_month:
        year = dates['this_month'][:-4]
    else:
        year = dates['next_month'][:-4]

    sheet_name = f'Havi Kiadások {year}'
    cost_sheet = wbook.worksheet(sheet_name)
    count = cost_sheet.acell('A1').numeric_value

    print(f'Modifying monthly costs table ({cost_sheet.title})  ...', end='  ')

    if is_new_month:
        for i in range(count):
            value = [f"=SUM(SUMIF('{dates['this_month']}'!$C$4:$C$100;'{sheet_name}'!$A{i + 2};'{dates['this_month']}'!$A$4:$A$100))"]
            col = None
            key_value = dates['this_month'][-3:]

            for key in mounthly_costs_columns:
                if key == key_value:
                    col = mounthly_costs_columns[key]
                    cost_sheet.update(f'{col[0]}{i + 2}', [value], raw=False)

        show_hide_cols(creds=creds, wbook_id=wbook.id, sheet_id=cost_sheet.id, is_hidden=False, start=col[1]-1, end=col[1])

    else:
        for j in range(count):
            value = [f"=SUM(SUMIF('{dates['next_month']}'!$C$4:$C$100;'{sheet_name}'!$A{j + 2};'{dates['next_month']}'!$A$4:$A$100))"]
            col = None
            key_value = dates['next_month'][-3:]

            for key in mounthly_costs_columns:
                if key == key_value:
                    col = mounthly_costs_columns[key]
                    cost_sheet.update(f'{col[0]}{j + 2}', [value], raw=False)

        show_hide_cols(creds=creds, wbook_id=wbook.id, sheet_id=cost_sheet.id, is_hidden=False, start=col[1] - 1, end=col[1])

    print('Done!')


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


if __name__ == '__main__':
    workb, creds = open_spread_sheet('Pénz másolata')
    date = calc_date()

    edit_mounthly_costs_sheet(wbook=workb, creds=creds, is_new_month=True, dates=date)
