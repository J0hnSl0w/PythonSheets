from Assets.operators import *
from Assets.enumerator_dickts import *


def init(spreadsheet_name, credentials_file, sheets_file):
    workb, creds = open_spread_sheet(spreadsheet_name, credentials_file)
    sheets = load_json(sheets_file)
    date = calc_date(sheets)

    return workb, creds, sheets, date


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
        elif '.' in wsheet.title:
            ssheet['month'][f'{wsheet.title}'] = [wsheet.id, index]
        else:
            ssheet['sums'][f'{wsheet.title}'] = [wsheet.id, index]

        percent = (i / l) * 100

        if k == 9:
            print(fg.li_blue + f'    kész: {round(percent)}%' + fg.rs)
            k = 0
        k += 1

    with open(sheets_json_file, 'w') as file:
        json.dump(ssheet, file)

    print(ef.italic + ef.bold + fg.li_blue + f'    kész: 100%' + ef.rs + fg.rs)


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
    # TODO kivitelezni, hogy univerzális legyen

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
    # TODO kivitelezni, hogy univerzális legyen

    repeat_formula_over_range(creds=creds, wbook_id=wbook.id, sheet_id=saving_sheet.id, formula=value,
                              start_col=start_col, end_col=end_col, start_row=start_row, end_row=end_row)

    show_hide_cols(creds=creds, wbook_id=wbook.id, sheet_id=saving_sheet.id, is_hidden=False, start=col[1] - 1,
                   end=col[1])

    print('Kész!')


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
    # workb, creds, sheets, date = init(spreadsheet_name, credentials_file, sheets_file)
    update_sheet_list(spreadsheet_name, credentials_file, sheets_file)

    print('ggg')
