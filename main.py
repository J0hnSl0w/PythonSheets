from Assets.gui import *
from Assets.logger import *

if __name__ == '__main__':
    table_info = {"spreadsheet_name": 'Sir Manó Pénze másolata',
                  "credentials_file": r'Assets/credentials.json',
                  "sheets_file": r'Assets/sheets_pm.jason'}

    app = MainWindow(None, table_info)
    app.title('Táblázat szerkesztő')
    app.configure(bg='black')
    app.resizable(width=False, height=False)

    stderrHandler = logging.StreamHandler()
    module_logger.addHandler(stderrHandler)
    # guiHandler = HandlerFeedback(app.logger)
    # module_logger.addHandler(guiHandler)
    module_logger.setLevel(logging.INFO)

    app.mainloop()
