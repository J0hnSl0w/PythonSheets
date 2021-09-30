import gspread
import json
import pandas

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

    return spreadsheet


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

def add_new_month(workb, sheetname, new_sheetname):
    sheet = workb.worksheet(sheetname)
    main_costs = workb.worksheet(title='Rendszeres kiadások')
    cost_values = []
    saving_values = []

    new_sheet = workb.duplicate_sheet(source_sheet_id=sheet.id, insert_sheet_index=4, new_sheet_name=new_sheetname)

    workb.values_clear(f"{new_sheet.title}!F4:J100")

    cell = new_sheet.acell('F3', 'FORMULA')
    n = cell.value[2:-5].split('.')
    new_value = cell.value[:2] + n[0] + '.' + f'0{int(n[1]) + 1}' + cell.value[-5:]
    new_sheet.update('F3', new_value, raw=False)

    cost_names = main_costs.get('A2:A8')
    saving_names = main_costs.get('G2:G6')

    for i in range(len(cost_names)):
        cell = main_costs.acell(f'B{i+2}', 'FORMULA')
        value = cell.numeric_value
        cost_values.append(value)

    for j in range(len(saving_names)):
        cell = main_costs.acell(f'H{j+2}', 'FORMULA')
        value = cell.numeric_value
        saving_values.append(value)

    new_sheet.update('G4:G10', cost_names, raw=True)
    new_sheet.update('G11:G15', saving_names, raw=True)
    # new_sheet.update('F4:F10', [cost_values], raw=False)
    # new_sheet.update('F11:F15', [saving_values], raw=False)

    print('fdg')


def add_new_month_v2(wbook, is_new_month=False):
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

        edit_mounthly_costs_sheet(wbook=wbook, is_new_month=is_new_month, dates=dates)

    else:
        new_sheet = wbook.duplicate_sheet(source_sheet_id=template_sheet.id,
                                          insert_sheet_index=11,
                                          new_sheet_name=dates['next_month'])

        cell = new_sheet.acell('A3', 'FORMULA')
        new_value = cell.value[:2] + dates['this_month'] + cell.value[-4:]
        new_sheet.update('A3', new_value, raw=False)

        for i in range(13):
            new_sheet.update(f'E{i+3}', [[f"{dates['next_month']}01."]], raw=False)

        edit_mounthly_costs_sheet(wbook=wbook, is_new_month=is_new_month, dates=dates)
        print('gd')


def edit_mounthly_costs_sheet(wbook,  is_new_month=False, dates=calc_date()):
    cost_sheet = wbook.worksheet('Havi Kiadások 2021')
    count = cost_sheet.acell('A1').numeric_value

    for i in range(count + 1):
        value = f"=SUM(SUMIF('{dates['next_month']}'!$C$4:$C$100;'Havi Kiadások 2021'!$A{i + 2};'{dates['next_month']}'!$A$4:$A$100);SUMIF('{dates['next_month']}'!$H$4:$H$100;'Havi Kiadások 2021'!$A{i + 2};'{dates['next_month']}'!$F$4:$F$100))"
        col = None
        key_value = dates['next_month'][-3:]

        for key in mounthly_costs_columns:
            if key == key_value:
                col = mounthly_costs_columns[key]
                cost_sheet.update(f'{col}{i + 2}', [value], raw=True)
                print('dfdf')

