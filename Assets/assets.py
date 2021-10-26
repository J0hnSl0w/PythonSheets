from Assets.operators import *
from Assets.enumerator_dickts import *


class EditSpreadsheet:
    def __init__(self, spreadsheet_name, credentials_file, sheets_file, cell_values_dict, date_row_range):
        self.wbook, self.creds = open_spread_sheet(spreadsheet_name, credentials_file)
        self.sheets_file = sheets_file
        self.sheets = load_json(self.sheets_file)
        self.date = calc_date(self.sheets)
        self.cell_values_dict = cell_values_dict
        self.date_row_range = date_row_range

    def update_sheet_list(self):
        spreadsheet = self.wbook

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

        with open(self.sheets_file, 'w') as file:
            json.dump(ssheet, file)

        print(ef.italic + ef.bold + fg.li_blue + f'    kész: 100%' + ef.rs + fg.rs)

    def edit_mounthly_costs_sheet(self):
        year = self.date['this_month'][:-4]
        sheet_name = f'Havi Kiadások {year}'
        cost_sheet = self.wbook.worksheet(sheet_name)
        count = cost_sheet.acell('A1').numeric_value

        print(f'A ({cost_sheet.title}) táblázat frissítése  ...')
        print(ef.italic + ef.bold + fg.yellow +
              '    Ha ezen folyamat közben történik valami, indítsd el újra a programot, és válaszd az 3-as, majd a 4-es menüpontot\n'
              '    Lehetséges, hogy csak 1 perc várakozás után fog újra működni a program.' + ef.rs + fg.rs + '  ... ', end='  ')

        col = None
        key_value = self.date['this_month'][-3:]

        for key in mounthly_columns:
            if key == key_value:
                col = mounthly_columns[key]
                break

        start_row = 1
        end_row = count+1
        start_col = col[1]-1
        end_col = col[1]

        formula = self.cell_values_dict['mc_value'].format(date=self.date['this_month'], sheet_name=sheet_name)

        repeat_formula_over_range(creds=self.creds, wbook_id=self.wbook.id, sheet_id=cost_sheet.id,
                                  formula=formula, start_col=start_col, end_col=end_col,
                                  start_row=start_row, end_row=end_row)

        show_hide_cols(creds=self.creds, wbook_id=self.wbook.id, sheet_id=cost_sheet.id, is_hidden=False,
                       start=col[1] - 1, end=col[1])

        print('Kész!')

    def edit_savings_sheet(self):
        year = self.date['this_month'][:-4]
        sheet_name = f'Megtakarítás részletező {year}'
        saving_sheet = self.wbook.worksheet(sheet_name)
        count = saving_sheet.acell('A1').numeric_value

        print(f'A ({saving_sheet.title}) táblázat szerkesztése  ...')
        print(ef.italic + ef.bold + fg.yellow +
              '    Ha ezen folyamat közben történik valami, indítsd el újra a programot, és válaszd az 4-es menüpontot\n'
              '    Lehetséges, hogy csak 1 perc várakozás után fog újra működni a program.' + ef.rs + fg.rs + '  ... ', end='  ')

        col = None
        key_value = self.date['this_month'][-3:]

        for key in mounthly_columns:
            if key == key_value:
                col = mounthly_columns[key]
                break

        start_row = 1
        end_row = count + 1
        start_col = col[1] - 1
        end_col = col[1]

        formula = self.cell_values_dict['sv_value'].format(date=self.date['this_month'], sheet_name=sheet_name)

        repeat_formula_over_range(creds=self.creds, wbook_id=self.wbook.id, sheet_id=saving_sheet.id,
                                  formula=formula, start_col=start_col, end_col=end_col,
                                  start_row=start_row, end_row=end_row)

        show_hide_cols(creds=self.creds, wbook_id=self.wbook.id, sheet_id=saving_sheet.id, is_hidden=False,
                       start=col[1] - 1, end=col[1])

        print('Kész!')

    def add_new_month(self):
        print(f'Új hónap hozzádása folyamatban  ...')
        print(ef.italic + ef.bold + fg.yellow +
              '    Ha a program sikeresen hozzáadta az új hónapot, de utána leáll,\n'
              '    akkor indítsd újra az alkalmazást és válaszd a 3-as, majd a 4-es menüpontot!\n'
              '    Ha nem tudta végigcsinálni ezt a folyamatot, nézd meg az interneten, hogy létre jött-e új hónap.\n'
              '    Ha igen, töröld ki és indítsd el újra a programot. Lehet hogy várni kell 1 percet.' + ef.rs + fg.rs + '  ... ', end='  ')

        template_sheet = self.wbook.worksheet('Template')
        is_new_year = False

        prev_sheet = self.wbook.worksheet(self.date['prev_month'])
        dict = find_sheet_by_name(prev_sheet.title, self.sheets['month'])

        new_sheet = self.wbook.duplicate_sheet(source_sheet_id=template_sheet.id,
                                               insert_sheet_index=dict[2] - 1,
                                               new_sheet_name=self.date['this_month'])

        cell = new_sheet.acell('A3', 'FORMULA')
        new_value = cell.value[:2] + self.date['prev_month'] + cell.value[-4:]
        new_sheet.update('A3', new_value, raw=False)

        cell2 = new_sheet.acell('H2', 'FORMULA')
        new_value2 = cell2.value[:31] + self.date['this_month'][:-4] + cell2.value[-5:]
        new_sheet.update('H2', new_value2, raw=False)

        for i in range(self.date_row_range):
            new_sheet.update(f'E{i + 3}', [[f"{self.date['this_month']}01."]], raw=False)

        past_year = int(self.date['prev_month'][:-4])
        year = int(self.date['this_month'][:-4])

        if past_year < year:
            is_new_year = True

        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=prev_sheet.id, is_hidden=True)

        print(f'Kész!  Az új lap neve: {new_sheet.title}')

        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=new_sheet.id, is_hidden=False)

        if is_new_year:
            EditSpreadsheet.add_new_year(self)

        else:
            EditSpreadsheet.edit_savings_sheet(self)
            EditSpreadsheet.edit_mounthly_costs_sheet(self)

    def add_new_year(self):
        print(f'Új év hozzáadása  ...')

        template_sheet_1 = self.wbook.worksheet('Megtakarítás részletező template')
        template_sheet_2 = self.wbook.worksheet('Havi Kiadások Template')

        prev_sheet_1 = self.wbook.worksheet(f"Megtakarítás részletező {self.date['prev_month'][:-4]}")
        prev_sheet_2 = self.wbook.worksheet(f"Havi Kiadások {self.date['prev_month'][:-4]}")

        dict = find_sheet_by_name(self.date['prev_month'], self.sheets['month'])

        new_sheet_1 = self.wbook.duplicate_sheet(source_sheet_id=template_sheet_1.id,
                                                 insert_sheet_index=dict[2] - 3,
                                                 new_sheet_name=f"Megtakarítás részletező {self.date['this_month'][:-4]}")

        new_sheet_2 = self.wbook.duplicate_sheet(source_sheet_id=template_sheet_2.id,
                                                 insert_sheet_index=dict[2] - 2,
                                                 new_sheet_name=f"Havi Kiadások {self.date['this_month'][:-4]}")

        EditSpreadsheet.edit_savings_sheet(self)
        EditSpreadsheet.edit_mounthly_costs_sheet(self)

        print('Kész!')

        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=new_sheet_1.id, is_hidden=False)
        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=new_sheet_2.id, is_hidden=False)
        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=prev_sheet_1.id, is_hidden=True)
        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=prev_sheet_2.id, is_hidden=True)

    def add_new_category(self):
        # TODO implement
        print('Még nincs kész.... :(')
        pass


if __name__ == '__main__':
    spreadsheet_name = 'Pénz másolata'
    credentials_file = 'credentials.json'
    sheets_file = 'sheets.jason'
    cell_values_dict = {'mc_value': "=SUM(SUMIF('{date}'!$C$4:$C$100;'{sheet_name}'!$A2;'{date}'!$A$4:$A$100))",
                        'sv_value': "=SUMIF('{date}'!$D$4:$D$100;'{sheet_name}'!$A2;'{date}'!$A$4:$A$100)*-1"}
    date_row_range = 13

    main_sheet = EditSpreadsheet(spreadsheet_name=spreadsheet_name, credentials_file=credentials_file,
                                 sheets_file=sheets_file, cell_values_dict=cell_values_dict,
                                 date_row_range=date_row_range)

    # workb, creds, sheets, date = init(spreadsheet_name, credentials_file, sheets_file)
    # update_sheet_list(spreadsheet_name, credentials_file, sheets_file

    # main_sheet.update_sheet_list()
    # main_sheet.add_new_month()
    # main_sheet.edit_savings_sheet()
    main_sheet.edit_mounthly_costs_sheet()

    print('ggg')
