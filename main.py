from assets import *

print('~'*25, ' Welcome to Walet editor! ', '~'*25)
print('What would you like to do?')
print('  1 - Add new month')
print('  2 - Add new year')
print('  3 - Add new category')

workb, creds = open_spread_sheet('Pénz másolata')

x = input('Choose a number: ')
while x != '1' or x != '2' or x != '3':
    x = input('Please choose form the numbers above: ')

    if x == '1':
        value = str(input('Is today is in the new month you want to add? (y/n) '))
        truth = None

        while value != 'y' or value != 'n':
            value = input('Please write only "y" or "n"! ')
            if value == 'y':
                truth = True
                break
            elif value == 'n':
                truth = False
                break

        add_new_month(wbook=workb, creds=creds, is_new_month=truth)
        break

    elif x == '2':
        print('not implemented yet')

    elif x == '3':
        print('not implemented yet')