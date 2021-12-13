import tkinter as tk


def start_gui(table):
    text_spec = {'bg': 'black',
                 'fg': 'white',
                 'width': 50,
                 'height': 2,
                 'font_c1': 'Sans 15 bold',
                 'font_c2': 'Sans 12'}
    buttons_spec = {'bg': 'grey',
                    'fg': 'black',
                    'widht': 30,
                    'height': 2,
                    'font': 'Sans 10'}

    main_window = tk.Tk(className='Táblázat szerkesztő')
    main_window.configure(bg='black')

    main_window.rowconfigure([0, 1, 2, 3, 4, 5, 6], minsize=10, weight=1)
    main_window.columnconfigure(0, minsize=200, weight=1)

    greetings_0 = tk.Label(text='Üdvözöllek a Táblázat szerkesztőben JohnSlow!',
                           fg=text_spec['fg'],
                           bg=text_spec['bg'],
                           width=text_spec['width'],
                           height=text_spec['height'],
                           font=text_spec['font_c1'])

    greetings_1 = tk.Label(text='Válaszd ki, mit szeretnél csinálni!',
                           fg=text_spec['fg'],
                           bg=text_spec['bg'],
                           width=text_spec['width'],
                           height=text_spec['height'],
                           font=text_spec['font_c2'])

    button_0 = tk.Button(text="Új hónap hozzáadása",
                         width=buttons_spec['widht'],
                         height=buttons_spec['height'],
                         bg=buttons_spec['bg'],
                         fg=buttons_spec['fg'],
                         font=buttons_spec['font'],
                         command=table.add_new_month)

    button_1 = tk.Button(text="Új kategória hozzáadása",
                         width=buttons_spec['widht'],
                         height=buttons_spec['height'],
                         bg=buttons_spec['bg'],
                         fg=buttons_spec['fg'],
                         font=buttons_spec['font'],
                         command=table.add_new_category)

    button_2 = tk.Button(text="Kiadás összesítő frissítése",
                         width=buttons_spec['widht'],
                         height=buttons_spec['height'],
                         bg=buttons_spec['bg'],
                         fg=buttons_spec['fg'],
                         font=buttons_spec['font'],
                         command=table.edit_mounthly_costs_sheet)

    button_3 = tk.Button(text="Megtakarítás összesítő frissítése",
                         width=buttons_spec['widht'],
                         height=buttons_spec['height'],
                         bg=buttons_spec['bg'],
                         fg=buttons_spec['fg'],
                         font=buttons_spec['font'],
                         command=table.edit_savings_sheet)

    button_4 = tk.Button(text="Táblázat adatok frissítése",
                         width=buttons_spec['widht'],
                         height=buttons_spec['height'],
                         bg=buttons_spec['bg'],
                         fg=buttons_spec['fg'],
                         font=buttons_spec['font'],
                         command=table.update_sheet_list)

    button_quit = tk.Button(text="Kilépés",
                            width=buttons_spec['widht'],
                            height=buttons_spec['height'],
                            bg=buttons_spec['bg'],
                            fg=buttons_spec['fg'],
                            font=buttons_spec['font'],
                            command=main_window.destroy)

    greetings_0.grid(row=0, column=0)
    greetings_1.grid(row=1, column=0, sticky='s')
    button_0.grid(row=2, column=0, pady=5)
    button_1.grid(row=3, column=0, pady=5)
    button_2.grid(row=4, column=0, pady=5)
    button_3.grid(row=5, column=0, pady=5)
    button_4.grid(row=6, column=0, pady=5)
    button_quit.grid(row=7, column=0, pady=5)

    main_window.mainloop()
