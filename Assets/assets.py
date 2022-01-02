from Assets.operators import *
from Assets.enumerator_dickts import *
from Assets.logger import *


class EditSpreadsheet:
    def __init__(self, spreadsheet_name, credentials_file, sheets_file):
        self.wbook, self.creds, self.feedback = open_spread_sheet(spreadsheet_name, credentials_file)
        self.sheets_file = sheets_file
        self.sheets = load_json(self.sheets_file)
        self.date = calc_date(self.sheets)
        self.values = self.wbook.worksheet(title='Formula').batch_get(['B1:E12'])

    def update_sheet_list(self):
        ssheet = {'templates': {},
                  'sums': {},
                  'month': {}}

        module_logger.info(prints['update_sheet_list'])

        l = len(self.wbook.worksheets())
        k = 0

        for i in range(l):
            wsheet = self.wbook.get_worksheet(i)
            index = i + 1

            if 'template'.lower() in wsheet.title.lower():
                ssheet['templates'][f'{wsheet.title}'] = [wsheet.id, index]
            elif '.' in wsheet.title:
                ssheet['month'][f'{wsheet.title}'] = [wsheet.id, index]
            else:
                ssheet['sums'][f'{wsheet.title}'] = [wsheet.id, index]

            percent = (i / l) * 100

            if k == 9:
                module_logger.info(f'    {round(percent)}%')
                k = 0
            k += 1

        with open(self.sheets_file, 'w') as file:
            json.dump(ssheet, file)

        module_logger.info('    100% -- Kész!\n')

    def edit_mounthly_costs_sheet(self):
        year = self.date['this_month'][:-4]
        sheet_name = f'Havi Kiadások {year}'
        cost_sheet = self.wbook.worksheet(sheet_name)

        count = int(self.values[0][0][0])
        mc_value = "=" + f"{self.values[0][6][0]}"

        module_logger.info(f'A ({cost_sheet.title}) táblázat frissítése  ...')
        module_logger.info(prints['edit_mounthly_costs_sheet'])

        col = mounthly_columns[self.date['this_month'][-3:]]

        start_row = 1
        end_row = count+1
        start_col = col[1]-1
        end_col = col[1]

        formula = mc_value.format(date=self.date['this_month'], sheet_name=sheet_name)

        repeat_formula_over_range(creds=self.creds, wbook_id=self.wbook.id, sheet_id=cost_sheet.id,
                                  formula=formula, start_col=start_col, end_col=end_col,
                                  start_row=start_row, end_row=end_row)

        show_hide_cols(creds=self.creds, wbook_id=self.wbook.id, sheet_id=cost_sheet.id, is_hidden=False,
                       start=col[1] - 1, end=col[1])

        module_logger.info('Havi összesítő frissítve!\n')

    def edit_savings_sheet(self):
        year = self.date['this_month'][:-4]
        sheet_name = f'Megtakarítás részletező {year}'
        saving_sheet = self.wbook.worksheet(sheet_name)

        count = int(self.values[0][1][0])
        sv_value = "=" + f"{self.values[0][7][0]}"

        module_logger.info(f'A ({saving_sheet.title}) táblázat szerkesztése  ...')
        module_logger.info(prints['edit_savings_sheet'])

        col = mounthly_columns[self.date['this_month'][-3:]]

        start_row = 1
        end_row = count + 1
        start_col = col[1] - 1
        end_col = col[1]

        formula = sv_value.format(date=self.date['this_month'], sheet_name=sheet_name)

        repeat_formula_over_range(creds=self.creds, wbook_id=self.wbook.id, sheet_id=saving_sheet.id,
                                  formula=formula, start_col=start_col, end_col=end_col,
                                  start_row=start_row, end_row=end_row)

        show_hide_cols(creds=self.creds, wbook_id=self.wbook.id, sheet_id=saving_sheet.id, is_hidden=False,
                       start=col[1] - 1, end=col[1])

        module_logger.info('Megtakarítás összesítő frissítve!\n')

    def add_new_month(self):
        starting = [self.values[0][5][0], self.values[0][5][1]]
        bank_account = [self.values[0][4][0], self.values[0][4][1]]
        date = [int(self.values[0][3][0]), self.values[0][3][1], int(self.values[0][3][2])]
        module_logger.info(prints['edit_new_month'])

        template_sheet = self.wbook.worksheet('Template')
        is_new_year = False

        prev_sheet = self.wbook.worksheet(self.date['prev_month'])
        dict = find_sheet_by_name(prev_sheet.title, self.sheets['month'])

        new_sheet = self.wbook.duplicate_sheet(source_sheet_id=template_sheet.id,
                                               insert_sheet_index=dict[2] - 1,
                                               new_sheet_name=self.date['this_month'])

        new_sheet.update(starting[0], "=" + starting[1].format(date=self.date['prev_month']), raw=False)
        new_sheet.update(bank_account[0], "=" + bank_account[1].format(date=self.date['this_month'][:-4]), raw=False)

        for i in range(date[0]):
            new_sheet.update(f"{date[1]}{i + date[2]}", [[f"{self.date['this_month']}01."]], raw=False)

        past_year = int(self.date['prev_month'][:-4])
        year = int(self.date['this_month'][:-4])

        if past_year < year:
            is_new_year = True

        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=prev_sheet.id, is_hidden=True)

        module_logger.info(f'Kész!  Az új lap neve: {new_sheet.title}\n')

        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=new_sheet.id, is_hidden=False)

        if is_new_year:
            EditSpreadsheet.add_new_year(self)

        else:
            EditSpreadsheet.edit_savings_sheet(self)
            EditSpreadsheet.edit_mounthly_costs_sheet(self)

        EditSpreadsheet.update_sheet_list(self)

    def add_new_year(self):
        hitelkeret_range = [int(self.values[0][2][0]), int(self.values[0][2][1])]
        hitelkeret = [self.values[0][8][0], self.values[0][8][1]]
        other = [self.values[0][9][0], int(self.values[0][9][1]), int(self.values[0][9][2]), int(self.values[0][9][3])]
        lakaskassza = [self.values[0][10][0], self.values[0][10][1]]
        nyugdijj = [self.values[0][11][0], self.values[0][11][1]]
        year = int(self.date['prev_month'][:-4])

        prev_sheets = True

        module_logger.info(f'Új év hozzáadása  ...')

        template_sheet_1 = self.wbook.worksheet('Megtakarítás részletező template')
        template_sheet_2 = self.wbook.worksheet('Havi Kiadások template')

        if len(find_sheet_by_name(f"Megtakarítás részletező {self.date['prev_month'][:-4]}", self.sheets)) != 0:
            prev_sheet_1 = self.wbook.worksheet(f"Megtakarítás részletező {self.date['prev_month'][:-4]}")
            prev_sheet_2 = self.wbook.worksheet(f"Havi Kiadások {self.date['prev_month'][:-4]}")
        else:
            prev_sheets = False

        dict = find_sheet_by_name(self.date['prev_month'], self.sheets['month'])

        new_sheet_1 = self.wbook.duplicate_sheet(source_sheet_id=template_sheet_1.id,
                                                 insert_sheet_index=dict[2] - 3,
                                                 new_sheet_name=f"Megtakarítás részletező {self.date['this_month'][:-4]}")

        new_sheet_2 = self.wbook.duplicate_sheet(source_sheet_id=template_sheet_2.id,
                                                 insert_sheet_index=dict[2] - 2,
                                                 new_sheet_name=f"Havi Kiadások {self.date['this_month'][:-4]}")

        new_sheet_1.update(hitelkeret[1], "=" + hitelkeret[0].format(range1=hitelkeret_range[0],
                                                                     range2=hitelkeret_range[1],
                                                                     year=year), raw=False)
        new_sheet_1.update(lakaskassza[1], "=" + lakaskassza[0].format(year=year), raw=False)
        new_sheet_1.update(nyugdijj[1], "=" + nyugdijj[0].format(year=year), raw=False)

        repeat_formula_over_range(creds=self.creds, wbook_id=self.wbook.id, sheet_id=new_sheet_1.id,
                                  formula="=" + other[0].format(year=year),
                                  start_col=other[1]-1, end_col=other[1],
                                  start_row=other[2], end_row=other[3])

        EditSpreadsheet.edit_savings_sheet(self)
        EditSpreadsheet.edit_mounthly_costs_sheet(self)

        module_logger.info('Új év hozzáadva!\n')

        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=new_sheet_1.id, is_hidden=False)
        show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=new_sheet_2.id, is_hidden=False)

        if prev_sheets:
            show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=prev_sheet_1.id, is_hidden=True)
            show_hide_wsheets(creds=self.creds, wbook_id=self.wbook.id, sheet_id=prev_sheet_2.id, is_hidden=True)

    def add_new_category(self):
        # TODO implement
        module_logger.info('Még nincs kész.... :(\n')


if __name__ == '__main__':
    spreadsheet = 'Pénz másolata'
    credentials = 'credentials.json'
    sheets = 'sheets_pm.jason'

    main_sheet = EditSpreadsheet(spreadsheet_name=spreadsheet, credentials_file=credentials,
                                 sheets_file=sheets)

    # main_sheet.update_sheet_list()
    main_sheet.add_new_month()
    # main_sheet.edit_savings_sheet()
    # main_sheet.edit_mounthly_costs_sheet()

    print('ggg')
