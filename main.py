from time import sleep

from Assets.gui import *
from Assets.logger import *

if __name__ == '__main__':
    # spreadsheet_name = 'Pénz'
    spreadsheet_name = 'Pénz másolata'
    credentials_file = r'Assets/credentials.json'
    sheets_file = r'Assets/sheets_pm.jason'

    # TODO printeket megcsinálni külön
    # TODO guit kiszépíteni

    table = EditSpreadsheet(spreadsheet_name=spreadsheet_name,
                            credentials_file=credentials_file,
                            sheets_file=sheets_file)

    app = MainWindow(None, table)
    app.title('Táblázat szerkesztő')
    app.configure(bg='black')
    app.resizable(width=False, height=False)

    stderrHandler = logging.StreamHandler()
    module_logger.addHandler(stderrHandler)
    guiHandler = HandlerFeedback(app.logger)
    module_logger.addHandler(guiHandler)
    module_logger.setLevel(logging.INFO)

    module_logger.info(table.feedback)

    app.mainloop()

