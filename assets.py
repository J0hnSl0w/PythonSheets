import gspread
import json

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


def add_new_month(workb, sheetname, new_sheetname):
    sheet_dict = load_json('sheets.jason')
    sheet_data = find_sheet_by_name(sheetname, sheet_dict)
    new_sheet = workb.duplicate_sheet(source_sheet_id=sheet_data[1], insert_sheet_index=sheet_data[2] - 1,
                                      new_sheet_name=new_sheetname)
    workb.values_clear(f"{new_sheet.title}!F4:J100")
    cell = new_sheet.acell('F3', 'FORMULA')

    n = cell.value[2:-5].split('.')
    new_value = cell.value[:2] + n[0] + '.' + f'0{int(n[1]) + 1}' + cell.value[-5:]

    new_sheet.update('F3', new_value, raw=False)

