from Assets.assets import *


def choose_from_menu():
    while True:
        x = input(fg.li_green + 'Choose a number from above: ' + fg.rs)
        if x in menu:
            break

    main_menu(x)

def is_new_month():
    truth = None

    while True:
        value = input(fg.li_green + 'Is today in the new month/year you want to add? (y/n) ' + fg.rs)
        if value in yn:
            break

    if value == 'y':
        truth = True
    elif value == 'n':
        truth = False

    return truth

def main_menu(x):
    workb, creds = open_spread_sheet(sheetname)

    if x == '1':
        truth = is_new_month()
        add_new_month(wbook=workb, creds=creds, is_new_month=truth)

    elif x == '2':
        add_new_category()

    elif x == '3':
        dates = calc_date()
        truth = is_new_month()
        edit_mounthly_costs_sheet(wbook=workb, creds=creds, is_new_month=truth, dates=dates)

    elif x == '4':
        dates = calc_date()
        truth = is_new_month()
        edit_savings_sheet(wbook=workb, creds=creds, is_new_month=truth, dates=dates)

    elif x == '5':
        update_sheet_list(sheetname)


if __name__ == '__main__':
    sheetname = 'Pénz'
    # sheetname = 'Pénz másolata'

    menu = ['1', '2', '3', '4', '5']
    yn = ['y', 'n']

    print(ef.bold + fg.cyan + '~'*25, ' Welcome to Walet editor JohnSlow! ', '~'*25 + fg.rs + ef.rs)
    print('What would you like to do?')
    print('  1 - Add new month')
    print('  2 - Add new category')
    print('  3 - Refresh costs table')
    print('  4 - Refresh savings table')
    print('  5 - Refresh sheet list')
    print(ef.italic + ef.bold + fg.yellow + '- If you want to add a new month, but the program sopped during\n'
                                            '  the procedure, please re-run the program and choose 3 and than 4.' + ef.rs + fg.rs)
    print(ef.italic + ef.bold + fg.yellow + '- Choose 5, if the program crases, and there is a "list index out of range"\n'
                                            '  line in the error message.\n' + ef.rs + fg.rs)

    choose_from_menu()

    while True:
        while True:
            y = str(input(fg.li_green + 'Do you want to do something else? (y/n) ' + fg.rs))
            if y in yn:
                break

        if y == 'y':
            choose_from_menu()

        elif y == 'n':

            update_sheet_list(sheetname)

            print(fg.cyan + 'Good bye!' + fg.rs)
            break
