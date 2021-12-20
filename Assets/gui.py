import tkinter as tk
from Assets.assets import *
from gspread.exceptions import APIError
import time


class MainWindow(tk.Tk):
    def __init__(self, parent, table_data):
        self.table_data = table_data

        tk.Tk.__init__(self, parent)
        self.parent = parent

        self.grid()

        # Def widgets
        self.frame_1 = tk.Frame(parent, bg=frame_spec['bg'])
        self.frame_2 = tk.Frame(parent, bg=frame_spec['bg'])
        self.frame_3 = tk.Frame(parent, bg=frame_spec['bg'])
        self.frame_4 = tk.Frame(parent, bg=frame_spec['bg'])

        self.greetings_0 = tk.Label(master=self.frame_1,
                                    text='Üdvözöllek a Táblázat szerkesztőben Sir Manó!',
                                    fg=text_spec['fg'],
                                    bg=text_spec['bg'],
                                    width=text_spec['width'],
                                    height=text_spec['height'],
                                    font=text_spec['font_c1'])

        self.greetings_1 = tk.Label(master=self.frame_1,
                                    text='Válaszd ki, mit szeretnél csinálni!',
                                    fg=text_spec['fg'],
                                    bg=text_spec['bg'],
                                    width=text_spec['width'],
                                    height=text_spec['height'],
                                    font=text_spec['font_c2'])

        self.text_0 = tk.Label(master=self.frame_3,
                               text='Az alábbiak közül akkor kell választani, ha valamilyen hiba történt:',
                               fg=text_spec['fg'],
                               bg=text_spec['bg'],
                               width=text_spec['width'],
                               height=text_spec['height'],
                               font=text_spec['font_c2'])

        self.button_nm = tk.Button(master=self.frame_2,
                                   text="Új hónap hozzáadása",
                                   width=buttons_spec['widht'],
                                   height=buttons_spec['height'],
                                   bg=buttons_spec['bg'],
                                   fg=buttons_spec['fg'],
                                   font=buttons_spec['font'])

        self.button_nc = tk.Button(master=self.frame_2,
                                   text="Új kategória hozzáadása",
                                   width=buttons_spec['widht'],
                                   height=buttons_spec['height'],
                                   bg=buttons_spec['bg'],
                                   fg=buttons_spec['fg'],
                                   font=buttons_spec['font'])

        self.button_ecs = tk.Button(master=self.frame_3,
                                    text="Kiadás összesítő frissítése",
                                    width=buttons_spec['widht'],
                                    height=buttons_spec['height'],
                                    bg=buttons_spec['bg'],
                                    fg=buttons_spec['fg'],
                                    font=buttons_spec['font'])

        # self.button_ess = tk.Button(master=self.frame_3,
        #                             text="Megtakarítás összesítő frissítése",
        #                             width=buttons_spec['widht'],
        #                             height=buttons_spec['height'],
        #                             bg=buttons_spec['bg'],
        #                             fg=buttons_spec['fg'],
        #                             font=buttons_spec['font'])

        self.button_ut = tk.Button(master=self.frame_3,
                                   text="Táblázat adatok frissítése",
                                   width=buttons_spec['widht'],
                                   height=buttons_spec['height'],
                                   bg=buttons_spec['bg'],
                                   fg=buttons_spec['fg'],
                                   font=buttons_spec['font'])

        self.button_quit = tk.Button(master=self.frame_4,
                                     text="Kilépés",
                                     width=buttons_spec['widht'],
                                     height=buttons_spec['height'],
                                     bg=buttons_spec['bg'],
                                     fg=buttons_spec['fg'],
                                     font=buttons_spec['font'])

        # self.logger = tk.Text(self,
        #                       state=logger_spec['state'],
        #                       width=logger_spec['width'],
        #                       height=logger_spec['height'])

        # Frames
        self.frame_1.grid(row=0, column=0)
        self.frame_2.grid(row=2, column=0, pady=10)
        self.frame_3.grid(row=3, column=0, pady=10)
        self.frame_4.grid(row=4, column=0, pady=30)
        # self.logger.grid(row=5, column=0, padx=30, pady=30)

        # Placement on frames
        # frame_1
        self.greetings_0.grid(row=0, column=0)
        self.greetings_1.grid(row=1, column=0, sticky='s')

        # frame_2
        self.button_nm.grid(row=0, column=0, pady=5)
        self.button_nc.grid(row=1, column=0, pady=5)

        # frame_3
        self.text_0.grid(row=0, column=0, pady=5)
        self.button_ecs.grid(row=1, column=0, pady=5)
        # self.button_ess.grid(row=2, column=0, pady=5)
        self.button_ut.grid(row=3, column=0, pady=5)

        # frame_4
        self.button_quit.grid(row=0, column=0, pady=5)

        # Map event handlers
        self.button_nm.bind("<ButtonRelease>", self.nm_handler)
        self.button_nc.bind("<ButtonRelease>", self.nc_handler)
        self.button_ecs.bind("<ButtonRelease>", self.ecs_handler)
        # self.button_ess.bind("<ButtonRelease>", self.ess_handler)
        self.button_ut.bind("<ButtonRelease>", self.ut_handler)
        self.button_quit.bind("<ButtonRelease>", self.quit_handler)

    @staticmethod
    def __open_table(table_info):
        table = EditSpreadsheet(spreadsheet_name=table_info['spreadsheet_name'],
                                credentials_file=table_info['credentials_file'],
                                sheets_file=table_info['sheets_file'])

        module_logger.info(table.feedback)
        return table

    @staticmethod
    def __sleeper(time_in_sec):
        module_logger.info('Sikertelen, próbáld újra a legutóbbi leírás alapján.\n'
                           'Ennyi idő múlva:\n')
        for i in range(time_in_sec, 0, -1):
            time.sleep(1)
            module_logger.info(i)
        module_logger.info('Most!\n')

    def nm_handler(self, event):
        try:
            opened_table = self.__open_table(self.table_data)
            opened_table.add_new_month()
        except APIError as err:
            if err.response.status_code == 429:
                self.__sleeper(10)
            else:
                raise

    def nc_handler(self, event):
        try:
            opened_table = self.__open_table(self.table_data)
            opened_table.add_new_category()
        except APIError as err:
            if err.response.status_code == 429:
                self.__sleeper(10)
            else:
                raise

    def ecs_handler(self, event):
        try:
            opened_table = self.__open_table(self.table_data)
            opened_table.edit_mounthly_costs_sheet()
        except APIError as err:
            if err.response.status_code == 429:
                self.__sleeper(10)
            else:
                raise

    def ess_handler(self, event):
        try:
            opened_table = self.__open_table(self.table_data)
            opened_table.edit_savings_sheet()

        except APIError as err:
            if err.response.status_code == 429:
                self.__sleeper(10)
            else:
                raise

    def ut_handler(self, event):
        try:
            opened_table = self.__open_table(self.table_data)
            opened_table.update_sheet_list()

        except APIError as err:
            if err.response.status_code == 429:
                self.__sleeper(10)
            else:
                raise

    def quit_handler(self, event):
        MainWindow.destroy(self)
