from Assets.assets import *
from Assets.gui import *

if __name__ == '__main__':
    spreadsheet_name = 'Pénz'
    # spreadsheet_name = 'Pénz másolata'
    credentials_file = r'Assets/credentials.json'
    sheets_file = r'Assets/sheets_p.jason'

    table = EditSpreadsheet(spreadsheet_name=spreadsheet_name, credentials_file=credentials_file,
                                 sheets_file=sheets_file)

    window = start_gui(table)
    print('Viszlát!')

