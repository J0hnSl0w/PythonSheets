import datetime
import json
from datetime import date

import gspread
from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
from sty import *

from enumerator_dickts import *


def open_spread_sheet(spreadsheet_name: str):
    scope = ['https://www.googleapis.com/auth/drive',
             'https://www.googleapis.com/auth/drive.file']

    file_name = 'credentials.json'
    creds = ServiceAccountCredentials.from_json_keyfile_name(file_name, scope)
    client = gspread.authorize(creds)

    spreadsheet = client.open(spreadsheet_name)
    print(f'Spreadsheet opened: {spreadsheet.title}')

    return spreadsheet, creds


def update_sheet_list(spreadsheet_name):
    spreadsheet, credentials = open_spread_sheet(spreadsheet_name)

    ssheet = {'templates': {},
              'sums': {},
              'month': {}}

    print(ef.italic + ef.bold + fg.yellow + 'Updating sheet list, please wait...' + ef.rs + fg.rs)

    l = len(spreadsheet.worksheets())
    k = 0

    for i in range(l):
        wsheet = spreadsheet.get_worksheet(i)
        index = i + 1

        if 'template'.lower() in wsheet.title.lower():  # or 'Template' in wsheet.title:
            ssheet['templates'][f'{wsheet.title}'] = [wsheet.id, index]
        elif 'Kiadások' in wsheet.title or 'Megtakarítás' in wsheet.title or \
                'összesítő' in wsheet.title or 'Rendszeres' in wsheet.title or 'Gyűjtés' in wsheet.title:
            ssheet['sums'][f'{wsheet.title}'] = [wsheet.id, index]
        else:
            ssheet['month'][f'{wsheet.title}'] = [wsheet.id, index]

        percent = (i / l) * 100

        if k == 9:
            print(fg.yellow + f'done: {round(percent)}%' + fg.rs)
            k = 0
        k += 1

    with open('sheets.jason', 'w') as file:
        json.dump(ssheet, file)

    print(ef.italic + ef.bold + fg.yellow + f'done: 100%' + ef.rs + fg.rs)


def load_json(json_file):
    f = open(json_file, 'r')
    dict = json.load(f)

    return dict


def find_sheet_by_name(name: str, sheet_dict: dict):
    sheet = []

    for key in sheet_dict:
        if str(key) == name:
            sheet.append(name)
            sheet.append(sheet_dict[key][0])
            sheet.append(sheet_dict[key][1])
            break

    return sheet


def calc_date():
    today = date.today()
    defoult_dict = {}
    date_dict = {'prev_month': '',
                 'this_month': '',}

    if today.month == 1:
        defoult_dict = {'prev_month': f'{today.year - 1}.{today.month + 11}.',
                        'this_month': f'{today.year}.0{today.month}.',
                        'next_month': f'{today.year}.0{today.month + 1}.'}

    elif today.month < 9:
        defoult_dict = {'prev_month': f'{today.year}.0{today.month - 1}.',
                        'this_month': f'{today.year}.0{today.month}.',
                        'next_month': f'{today.year}.0{today.month + 1}.'}

    elif today.month == 9:
        defoult_dict = {'prev_month': f'{today.year}.0{today.month - 1}.',
                        'this_month': f'{today.year}.0{today.month}.',
                        'next_month': f'{today.year}.{today.month + 1}.'}

    elif today.month == 10:
        defoult_dict = {'prev_month': f'{today.year}.0{today.month - 1}.',
                        'this_month': f'{today.year}.{today.month}.',
                        'next_month': f'{today.year}.{today.month + 1}.'}

    elif today.month == 11:
        defoult_dict = {'prev_month': f'{today.year}.{today.month - 1}.',
                        'this_month': f'{today.year}.{today.month}.',
                        'next_month': f'{today.year}.{today.month + 1}.'}

    elif today.month == 12:
        defoult_dict = {'prev_month': f'{today.year}.{today.month - 1}.',
                        'this_month': f'{today.year}.{today.month}.',
                        'next_month': f'{today.year + 1}.{today.month - 11}.'}

    sn = defoult_dict['this_month']
    k = find_sheet_by_name(sn, load_json('sheets.jason')['month'])
    if len(k) != 0 and k[0 == sn]:
        date_dict['this_month'] = defoult_dict['next_month']
        date_dict['prev_month'] = defoult_dict['this_month']

    else:
        sn_prev = defoult_dict['prev_month']
        l = find_sheet_by_name(sn, load_json('sheets.jason')['month'])
        if len(l) != 0 and k[0 == sn_prev]:
            date_dict['this_month'] = defoult_dict['this_month']
            date_dict['prev_month'] = defoult_dict['prev_month']
        else:
            print("You can't create a spreadsheet for this month yet")

    return date_dict


def edit_mounthly_costs_sheet(wbook, creds, is_new_month=False, dates=calc_date()):
    if is_new_month:
        year = dates['this_month'][:-4]
    else:
        year = dates['next_month'][:-4]

    sheet_name = f'Havi Kiadások {year}'
    cost_sheet = wbook.worksheet(sheet_name)
    count = cost_sheet.acell('A1').numeric_value

    print(f'Modifying monthly costs table ({cost_sheet.title})  ...', end='  ')

    if is_new_month:
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

    else:
        col = None
        key_value = dates['next_month'][-3:]

        for key in mounthly_columns:
            if key == key_value:
                col = mounthly_columns[key]
                break

        start_row = 1
        end_row = count+1
        start_col = col[1]-1
        end_col = col[1]

        value = f"=SUM(SUMIF('{dates['next_month']}'!$C$4:$C$100;'{sheet_name}'!$A2;'{dates['next_month']}'!$A$4:$A$100))"

        repeat_formula_over_range(creds=creds, wbook_id=wbook.id, sheet_id=cost_sheet.id, formula=value,
                                  start_col=start_col, end_col=end_col, start_row=start_row, end_row=end_row)

    print('Done!')


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


def edit_savings_sheet(wbook, creds, is_new_month=False, dates=calc_date()):
    if is_new_month:
        year = dates['this_month'][:-4]
    else:
        year = dates['next_month'][:-4]

    sheet_name = f'Megtakarítás részletező {year}'
    saving_sheet = wbook.worksheet(sheet_name)
    count = saving_sheet.acell('A1').numeric_value

    print(f'Modifying savings table ({saving_sheet.title})  ...', end='  ')

    if is_new_month:
        for i in range(count):
            value = [
                f"=SUMIF('{dates['this_month']}'!$D$4:$D$100;'{sheet_name}'!$A{i + 2};'{dates['this_month']}'!$A$4:$A$100)*-1"]
            col = None
            key_value = dates['this_month'][-3:]

            for key in mounthly_columns:
                if key == key_value:
                    col = mounthly_columns[key]
                    saving_sheet.update(f'{col[0]}{i + 2}', [value], raw=False)

        show_hide_cols(creds=creds, wbook_id=wbook.id, sheet_id=saving_sheet.id, is_hidden=False, start=col[1] - 1,
                       end=col[1])

    else:
        for j in range(count):
            value = [
                f"=SUMIF('{dates['next_month']}'!$D$4:$D$100;'{sheet_name}'!$A{j + 2};'{dates['next_month']}'!$A$4:$A$100)*-1"]
            col = None
            key_value = dates['next_month'][-3:]

            for key in mounthly_columns:
                if key == key_value:
                    col = mounthly_columns[key]
                    saving_sheet.update(f'{col[0]}{j + 2}', [value], raw=False)

        show_hide_cols(creds=creds, wbook_id=wbook.id, sheet_id=saving_sheet.id, is_hidden=False, start=col[1] - 1,
                       end=col[1])

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


def add_new_month(wbook, creds, is_new_month=False):
    dates = calc_date()
    template_sheet = wbook.worksheet('Template')
    is_new_year = False

    if is_new_month:
        prev_sheet = wbook.worksheet(dates['prev_month'])
        dict = find_sheet_by_name(dates['prev_month'], load_json('sheets.jason')['month'])

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

    else:
        prev_sheet = wbook.worksheet(dates['prev_month'])
        dict = find_sheet_by_name(dates['this_month'], load_json('sheets.jason')['month'])

        new_sheet = wbook.duplicate_sheet(source_sheet_id=template_sheet.id,
                                          insert_sheet_index=dict[2] - 1,
                                          new_sheet_name=dates['next_month'])

        cell = new_sheet.acell('A3', 'FORMULA')
        new_value = cell.value[:2] + dates['this_month'] + cell.value[-4:]
        new_sheet.update('A3', new_value, raw=False)

        cell2 = new_sheet.acell('H2', 'FORMULA')
        new_value2 = cell2.value[:31] + dates['next_month'][:-4] + cell2.value[-5:]
        new_sheet.update('H2', new_value2, raw=False)

        for i in range(13):
            new_sheet.update(f'E{i + 3}', [[f"{dates['next_month']}01."]], raw=False)

        year = int(dates['this_month'][:-4])
        next_year = int(dates['next_month'][:-4])

        if year < next_year:
            is_new_year = True

    print('New sheet added succesfully!')

    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=new_sheet.id, is_hidden=False)

    edit_savings_sheet(wbook=wbook, creds=creds, is_new_month=is_new_month, dates=dates)
    edit_mounthly_costs_sheet(wbook=wbook, creds=creds, is_new_month=is_new_month, dates=dates)

    if is_new_year:
        add_new_year(wbook=wbook, creds=creds, is_new_month=is_new_month)


def add_new_year(wbook, creds, is_new_month):
    dates = calc_date()
    template_sheet_1 = wbook.worksheet('Megtakarítás részletező template')
    template_sheet_2 = wbook.worksheet('Havi Kiadások Template')

    if is_new_month:
        prev_sheet_1 = wbook.worksheet(f"Megtakarítás részletező {dates['prev_month'][:-4]}")
        prev_sheet_2 = wbook.worksheet(f"Havi Kiadások {dates['prev_month'][:-4]}")

        dict = find_sheet_by_name(dates['prev_month'], load_json('sheets.jason')['month'])

        new_sheet_1 = wbook.duplicate_sheet(source_sheet_id=template_sheet_1.id,
                                            insert_sheet_index=dict[2],
                                            new_sheet_name=f"Megtakarítás részletező {dates['this_month'][:-4]}")

        new_sheet_2 = wbook.duplicate_sheet(source_sheet_id=template_sheet_2.id,
                                            insert_sheet_index=dict[2],
                                            new_sheet_name=f"Havi Kiadások {dates['this_month'][:-4]}")

        edit_savings_sheet(wbook=wbook, creds=creds, is_new_month=is_new_month, dates=dates)
        edit_mounthly_costs_sheet(wbook=wbook, creds=creds, is_new_month=is_new_month, dates=dates)

    else:
        prev_sheet_1 = wbook.worksheet(f"Megtakarítás részletező {dates['this_month'][:-4]}")
        prev_sheet_2 = wbook.worksheet(f"Havi Kiadások {dates['this_month'][:-4]}")

        dict = find_sheet_by_name(dates['this_month'], load_json('sheets.jason')['month'])

        new_sheet_1 = wbook.duplicate_sheet(source_sheet_id=template_sheet_1.id,
                                            insert_sheet_index=dict[2] - 1,
                                            new_sheet_name=f"Megtakarítás részletező {dates['next_month'][:-4]}")

        new_sheet_2 = wbook.duplicate_sheet(source_sheet_id=template_sheet_2.id,
                                            insert_sheet_index=dict[2] - 1,
                                            new_sheet_name=f"Megtakarítás részletező {dates['next_month'][:-4]}")

        edit_savings_sheet(wbook=wbook, creds=creds, is_new_month=is_new_month, dates=dates)
        edit_mounthly_costs_sheet(wbook=wbook, creds=creds, is_new_month=is_new_month, dates=dates)

    print('New year added succesfully!')

    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=new_sheet_1.id, is_hidden=False)
    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=new_sheet_2.id, is_hidden=False)
    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=prev_sheet_1.id, is_hidden=True)
    show_hide_wsheets(creds=creds, wbook_id=wbook.id, sheet_id=prev_sheet_2.id, is_hidden=True)


def add_new_category():
    # TODO implement
    pass


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



if __name__ == '__main__':
    workb, creds = open_spread_sheet('Pénz másolata')

    # update_sheet_list('Pénz másolata')
    kefir = load_json('sheets.jason')['month']

    # add_new_month(wbook=workb, creds=creds, is_new_month=True)
    # calc_date()

    # sheet = workb.worksheet('Havi Kiadások 2021')
    # cell = sheet.acell('H2', 'FORMULA')

    # insert_row(creds=creds, wbook_id=workb.id, sheet_id=sheet.id, start=5, end=6, inherit_from_before=True)
    # edit_mounthly_costs_sheet(wbook=workb, creds=creds, is_new_month=False, dates=calc_date())

    # calc_date()
    print('gg')
