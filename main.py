from assets import *


def choose_from_menu():
    while True:
        x = input('Choose a number from above (1, 2 or 3): ')
        if x in menu:
            break

    main_menu(x)

def is_new_month():
    truth = None

    while True:
        value = input('Is today is in the new month/year you want to add? (y/n) ')
        if value in yn:
            break

    if value == 'y':
        truth = True
    elif value == 'n':
        truth = False

    return truth

def main_menu(x):
    workb, creds = open_spread_sheet('Pénz másolata')

    if x == '1':
        truth = is_new_month()
        add_new_month(wbook=workb, creds=creds, is_new_month=truth)

    elif x == '3':
        add_new_category()


if __name__ == '__main__':
    menu = ['1', '2']
    yn = ['y', 'n']

    print('~'*25, ' Welcome to Walet editor JohnSlow! ', '~'*25)
    print('What would you like to do?')
    print('  1 - Add new month')
    print('  2 - Add new category')

    choose_from_menu()

    while True:
        while True:
            y = str(input('Do you want to do something else? (y/n) '))
            if y in yn:
                break

        if y == 'y':
            choose_from_menu()

        elif y == 'n':
            print('Good bye!')
            break