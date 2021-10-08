import datetime
import json
import gspread

from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
from sty import *
from Assets.enumerator_dickts import *


def init(spreadsheet_name, credentials_file, sheets_file):
    workb, creds = open_spread_sheet(spreadsheet_name, credentials_file)
    sheets = load_json(sheets_file)
    date = calc_date(sheets)

    return workb, creds, sheets, date


def load_json(json_file):
    f = open(json_file, 'r')
    dict = json.load(f)

    return dict


def open_spread_sheet(spreadsheet_name: str, credentials_file: str):
    scope = ['https://www.googleapis.com/auth/drive',
             'https://www.googleapis.com/auth/drive.file']

    file_name = credentials_file
    creds = ServiceAccountCredentials.from_json_keyfile_name(file_name, scope)
    client = gspread.authorize(creds)

    spreadsheet = client.open(spreadsheet_name)
    print(f'Táblázat megnyitva: {spreadsheet.title}')

    return spreadsheet, creds


def update_sheet_list(spreadsheet_name, credentials_file, sheets_json_file):
    spreadsheet, creds = open_spread_sheet(spreadsheet_name, credentials_file)

    ssheet = {'templates': {},
              'sums': {},
              'month': {}}

    print(ef.italic + ef.bold + fg.li_blue + 'A lapok adatait tartalmazó fájl frissítése, kérlek várj  ...' + ef.rs + fg.rs)
    print(ef.italic + ef.bold + fg.yellow +
          '    Ha ezen folyamat közben történik valami, indítsd el újra a programot, és válaszd az 5-ös menüpontot\n'
          '    Lehetséges, hogy csak 1 perc várakozás után fog újra működni a program.' + ef.rs + fg.rs)

    l = len(spreadsheet.worksheets())
    k = 0

    for i in range(l):
        wsheet = spreadsheet.get_worksheet(i)
        index = i + 1

        if 'template'.lower() in wsheet.title.lower():
            ssheet['templates'][f'{wsheet.title}'] = [wsheet.id, index]
        elif 'Kiadások' in wsheet.title or 'Megtakarítás' in wsheet.title or \
                'összesítő' in wsheet.title or 'Rendszeres' in wsheet.title or 'Gyűjtés' in wsheet.title:
            ssheet['sums'][f'{wsheet.title}'] = [wsheet.id, index]
        else:
            ssheet['month'][f'{wsheet.title}'] = [wsheet.id, index]

        percent = (i / l) * 100

        if k == 9:
            print(fg.li_blue + f'kész: {round(percent)}%' + fg.rs)
            k = 0
        k += 1

    with open(sheets_json_file, 'w') as file:
        json.dump(ssheet, file)

    print(ef.italic + ef.bold + fg.li_blue + f'kész: 100%' + ef.rs + fg.rs)


def find_sheet_by_name(name: str, sheet_dict: dict):
    sheet = []

    for key in sheet_dict:
        if str(key) == name:
            sheet.append(name)
            sheet.append(sheet_dict[key][0])
            sheet.append(sheet_dict[key][1])
            break

    return sheet


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
        exeption = ef.italic + ef.bold + fg.yellow +\
                   'Kérlek, rendezd a napi kiadásokat tartalmazó lapokat balról jobbra növekvő sorrendbe,' \
                   'indítsd el a programot és válaszd az 5-ös menüpontot!' + fg.rs + ef.rs
        raise Exception(exeption)
    return date_dict


def edit_mounthly_costs_sheet(wbook, creds, dates):
    year = dates['this_month'][:-4]
    sheet_name = f'Havi Kiadások {year}'
    cost_sheet = wbook.worksheet(sheet_name)
    count = cost_sheet.acell('A1').numeric_value

    print(f'A ({cost_sheet.title}) táblázat frissítése  ...')
    print(ef.italic + ef.bold + fg.yellow +
          '    Ha ezen folyamat közben történik valami, indítsd el újra a programot, és válaszd az 3-as, majd a 4-es menüpontot\n'
          '    Lehetséges, hogy csak 1 perc várakozás után fog újra működni a program.' + ef.rs + fg.rs + '  ... ', end='  ')

    col = None
    key_value = dates['this_month'][-3:]

    for key in mounthly_columns:
        if key == key_value:
            col = mounthly_columns[key]
            break

    start_row = 1
    end_row = count+1
    start_col = col[1]-1
    end_col = col[1]

    value = f"=SUM(SUMIF('{dates['this_month']}'!$C$4:$C$100;'{sheet_name}'!$A2;'{dates['this_month']}'!$A$4:$A$100))"

    repeat_formula_over_range(creds=creds, wbook_id=wbook.id, sheet_id=cost_sheet.id, formula=value,
                              start_col=start_col, end_col=end_col, start_row=start_row, end_row=end_row)

    show_hide_cols(creds=creds, wbook_id=wbook.id, sheet_id=cost_sheet.id, is_hidden=False, start=col[1] - 1,
                   end=col[1])

    print('Kész!')


def edit_savings_sheet(wbook, creds, dates):
    year = dates['this_month'][:-4]
    sheet_name = f'Megtakarítás részletező {year}'
    saving_sheet = wbook.worksheet(sheet_name)
    count = saving_sheet.acell('A1').numeric_value

    print(f'A ({saving_sheet.title}) táblázat szerkesztése  ...')
    print(ef.italic + ef.bold + fg.yellow +
          '    Ha ezen folyamat közben történik valami, indítsd el újra a programot, és válaszd az 4-es menüpontot\n'
          '    Lehetséges, hogy csak 1 perc várakozás után fog újra működni a program.' + ef.rs + fg.rs + '  ... ', end='  ')

    col = None
    key_value = dates['this_month'][-3:]

    for key in mounthly_columns:
        if key == key_value:
            col = mounthly_columns[key]
            break

    start_row = 1
    end_row = count + 1
    start_col = col[1] - 1
    end_col = col[1]

    value = f"=SUMIF('{dates['this_month']}'!$D$4:$D$100;'{sheet_name}'!$A2;'{dates['this_month']}'!$A$4:$A$100)*-1"

    repeat_formula_over_range(creds=creds, wbook_id=wbook.id, sheet_id=saving_sheet.id, formula=value,
                              start_col=start_col, end_col=end_col, start_row=start_row, end_row=end_row)

    show_hide_cols(creds=creds, wbook_id=wbook.id, sheet_id=saving_sheet.id, is_hidden=False, start=col[1] - 1,
                   end=col[1])

    print('Kész!')


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


def add_new_month(wbook, creds, sheets_dict, dates):
    print(f'Új hónap hozzádása folyamatban  ...', end='  ')

    template_sheet = wbook.worksheet('Template')
    is_new_year = False

    prev_sheet = wbook.worksheet(dates['prev_month'])
    dict = find_sheet_by_name(prev_sheet.title, sheets_dict['month'])

    new_sheet = wbook.duplicate_sheet(source_sheet_id=template_sheet.id,
                                      insert_sheet_index=dict[2] - 1,
                                      new_sheet_name=dates['this_month'])

    cell = new_sheet.acell('A3', 'FORMULA')
    new_value = cell.value[:2] + dates['prev_month'] + cell.value[-4:]
    new_sheet.update('A3', new_value, raw=False)

    cell2 = new_sheet.acell('H2', 'FORMULA')
    new_value2 = cell2.value[:31] + dates['this_month'][:-4] + cell2.value[-5:]
    new_sheet.update('H2', new_value2, raw=False)

    for i in range(13):
        new_sheet.update(f'E{i + 3}', [[f"{dates['this_month']}01."]], raw=False)

    past_year = int(dates['prev_month'][:-4])
    year = int(dates['this_month'][:-4])

    if past_year < year:
        is_new_year = True

    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=prev_sheet.id, is_hidden=True)

    print(f'Kész!  Az új lap neve: {new_sheet.title}')

    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=new_sheet.id, is_hidden=False)

    if is_new_year:
        add_new_year(wbook=wbook, creds=creds, sheets_dict=sheets_dict, dates=dates)

    else:
        edit_savings_sheet(wbook=wbook, creds=creds, dates=dates)
        edit_mounthly_costs_sheet(wbook=wbook, creds=creds, dates=dates)


def add_new_year(wbook, creds, sheets_dict, dates):
    print(f'Új év hozzáadása  ...')

    template_sheet_1 = wbook.worksheet('Megtakarítás részletező template')
    template_sheet_2 = wbook.worksheet('Havi Kiadások Template')

    prev_sheet_1 = wbook.worksheet(f"Megtakarítás részletező {dates['prev_month'][:-4]}")
    prev_sheet_2 = wbook.worksheet(f"Havi Kiadások {dates['prev_month'][:-4]}")

    dict = find_sheet_by_name(dates['prev_month'], sheets_dict['month'])

    new_sheet_1 = wbook.duplicate_sheet(source_sheet_id=template_sheet_1.id,
                                        insert_sheet_index=dict[2] - 3,
                                        new_sheet_name=f"Megtakarítás részletező {dates['this_month'][:-4]}")

    new_sheet_2 = wbook.duplicate_sheet(source_sheet_id=template_sheet_2.id,
                                        insert_sheet_index=dict[2] - 2,
                                        new_sheet_name=f"Havi Kiadások {dates['this_month'][:-4]}")

    edit_savings_sheet(wbook=wbook, creds=creds, dates=dates)
    edit_mounthly_costs_sheet(wbook=wbook, creds=creds, dates=dates)

    print('Kész!')

    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=new_sheet_1.id, is_hidden=False)
    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=new_sheet_2.id, is_hidden=False)
    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=prev_sheet_1.id, is_hidden=True)
    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=prev_sheet_2.id, is_hidden=True)


def add_new_category():
    # TODO implement
    print('Még nincs kész.... :(')
    pass


if __name__ == '__main__':
    spreadsheet_name = 'Pénz másolata'
    credentials_file = 'credentials.json'
    sheets_file = 'sheets.jason'
    workb, creds, sheets, date = init(spreadsheet_name, credentials_file, sheets_file)

    print('gg')
