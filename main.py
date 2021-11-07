from Assets.assets import *


def choose_from_menu():
    while True:
        x = input(fg.li_green + '\nVálassz egy számot a fentiek közül: ' + fg.rs)
        if x in menu:
            break

    sheets_updated, month_added = main_menu(x)
    return sheets_updated, month_added


def main_menu(x):
    sheets_updated = False
    month_added = False
    main_sheet = EditSpreadsheet(spreadsheet_name=spreadsheet_name, credentials_file=credentials_file,
                                 sheets_file=sheets_file)

    if x == '1':
        main_sheet.add_new_month()
        month_added = True

    elif x == '2':
        main_sheet.add_new_category()

    elif x == '3':
        main_sheet.edit_mounthly_costs_sheet()

    # elif x == '4':
    #     main_sheet.edit_savings_sheet()

    elif x == '4':
        main_sheet.update_sheet_list()
        sheets_updated = True

    return sheets_updated, month_added




if __name__ == '__main__':
    spreadsheet_name = 'Klárka Pénze'
    credentials_file = r'Assets/credentials.json'
    sheets_file = r'Assets/sheets_p.jason'

    menu = ['1', '2', '3', '4']
    yn = ['i', 'n']

    print(ef.bold + fg.cyan + '~'*25, ' Üdvözöllek a Táblázat-szerkesztőben Klárka! ', '~'*25 + fg.rs + ef.rs)
    print('Mit szeretnél csinálni?')
    print(ef.italic + 'Alapműveletek:' + ef.rs)
    print('  1 - Új hónap hozzáadása')
    print('  2 - Új kategória hozzáadása\n')
    print(ef.italic + 'Hiba elhárításához szükséges műveletek:' + ef.rs)
    print('  3 - A kiadások összesítő táblázatának frissítése')
    # print('  4 - A megtakarításokat összesítő táblázat frissítése')
    print('  4 - A lapok adatait tartalmazó fájl frissítése')
    print(ef.italic + ef.bold + fg.yellow +
          '    Ezeket csak akkor kell használni, ha a program futása közben valamilyen hibaüzenet jön elő.\n'
          '    A műveletek során hiba esetén kövesd, az ilyen betűtípussal írt javaslatokat!\n'
          '    Ha a javaslatok ellenére sem stimmel valami, szólj nekem, kitalálunk valamit! :)' + ef.rs + fg.rs)

    sheets_updated, month_added = choose_from_menu()

    while True:
        while True:
            print(ef.italic + ef.bold + fg.yellow +
                  '\n    Mielőtt erre válaszolsz, érdemes várni egy kicsit. :)' + ef.rs + fg.rs)
            y = str(input(fg.li_green + 'Szeretnél még csinálni valamit? (i/n) ' + fg.rs))
            if y in yn:
                break

        if month_added and not sheets_updated:
            main_menu('5')

        if y == 'i':
            sheets_updated, month_added = choose_from_menu()

        elif y == 'n':
            print(fg.cyan + 'Viszlát!' + fg.rs)
            break

