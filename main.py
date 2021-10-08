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

    if x == '5':
        update_sheet_list(spreadsheet_name=spreadsheet_name,
                          credentials_file=credentials_file,
                          sheets_json_file=sheets_file)
        sheets_updated = True

    else:
        workb, creds, sheets, date = init(spreadsheet_name=spreadsheet_name, credentials_file=credentials_file,
                                          sheets_file=sheets_file)

        if x == '1':
            print(ef.italic + ef.bold + fg.yellow +
                  '    Ha a program sikeresen hozzáadta az új hónapot, de utána leáll,\n'
                  '    akkor indítsd újra az alkalmazást és válaszd a 3-as, majd a 4-es menüpontot!\n'
                  '    Ha nem tudta végigcsinálni ezt a folyamatot, nézd meg az interneten, hogy létre jött-e új hónap.\n'
                  '    Ha igen, töröld ki és indítsd el újra a programot. Lehet hogy várni kell 1 percet.' + ef.rs + fg.rs)
            add_new_month(wbook=workb, creds=creds, sheets_dict=sheets, dates=date)
            month_added = True

        elif x == '2':
            add_new_category()

        elif x == '3':
            edit_mounthly_costs_sheet(wbook=workb, creds=creds, dates=date)

        elif x == '4':
            edit_savings_sheet(wbook=workb, creds=creds, dates=date)

    return sheets_updated, month_added




if __name__ == '__main__':
    # spreadsheet_name = 'Pénz'
    spreadsheet_name = 'Pénz másolata'
    credentials_file = r'Assets/credentials.json'
    sheets_file = r'Assets/sheets.jason'

    menu = ['1', '2', '3', '4', '5']
    yn = ['i', 'n']

    print(ef.bold + fg.cyan + '~'*25, ' Üdvözöllek a Táblázat-szerkesztőben JohnSlow! ', '~'*25 + fg.rs + ef.rs)
    print('Mit szeretnél csinálni?')
    print(ef.italic + 'Alapműveletek:' + ef.rs)
    print('  1 - Új hónap hozzáadása')
    print('  2 - Új kategória hozzáadása\n')
    print(ef.italic + 'Hiba elhárításához szükséges műveletek:' + ef.rs)
    print('  3 - A kiadások összesítő táblázatának frissítése')
    print('  4 - A megtakarításokat összesítő táblázat frissítése')
    print('  5 - A lapok adatait tartalmazó fájl frissítése')
    print(ef.italic + ef.bold + fg.yellow +
          '    Ezeket csak akkor kell használni, ha a program futása közben valamilyen hibaüzenet jön elő.\n'
          '    A műveletek során hiba esetén kövesd, az ilyen betűtípussal írt javaslatokat!\n'
          '    Ha a javaslatok ellenére sem stimmel valami, szólj nekem, kitalálunk valamit! :)' + ef.rs + fg.rs)


    # print(ef.italic + ef.bold + fg.yellow + '- If you want to add a new month, but the program sopped during\n'
    #                                         '  the procedure, please re-run the program and choose 3 and than 4.' + ef.rs + fg.rs)
    # print(ef.italic + ef.bold + fg.yellow + '- Choose 5, if the program crases, and there is a "list index out of range"\n'
    #                                         '  line in the error message.\n' + ef.rs + fg.rs)

    sheets_updated, month_added = choose_from_menu()

    while True:
        while True:
            print(ef.italic + ef.bold + fg.yellow +
                  '\n    Mielőtt erre válaszolsz, érdemes várni egy kicsit. :)' + ef.rs + fg.rs)
            y = str(input(fg.li_green + 'Szeretnél még csinálni valamit? (i/n) ' + fg.rs))
            if y in yn:
                break

        if month_added and not sheets_updated:
            update_sheet_list(spreadsheet_name=spreadsheet_name,
                              credentials_file=credentials_file,
                              sheets_json_file=sheets_file)

        if y == 'i':
            sheets_updated, month_added = choose_from_menu()

        elif y == 'n':
            print(fg.cyan + 'Viszlát!' + fg.rs)
            break

