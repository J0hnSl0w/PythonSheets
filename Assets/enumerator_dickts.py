mounthly_columns = {'01.': ['B', 2],
                    '02.': ['C', 3],
                    '03.': ['D', 4],
                    '04.': ['E', 5],
                    '05.': ['F', 6],
                    '06.': ['G', 7],
                    '07.': ['H', 8],
                    '08.': ['I', 9],
                    '09.': ['J', 10],
                    '10.': ['K', 11],
                    '11.': ['L', 12],
                    '12.': ['L', 13]}

values = {'mc_value': "=SUM(SUMIF('{date}'!$C$4:$C$100;'{sheet_name}'!$A2;'{date}'!$A$4:$A$100))",
          'sv_value': "=SUMIF('{date}'!$D$4:$D$100;'{sheet_name}'!$A2;'{date}'!$A$4:$A$100)*-1",
          'starting': ['A3', "='{date}'!A2"],
          'bank_account': ['H2', "=$A$2+'Megtakarítás részletező {date}'!N38"],
          'date': [13, 'E', 3]}

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

frame_spec = {'bg': 'black'}

logger_spec = {'state': 'disabled',
               'width': 80,
               'height': 10}

prints = {'calc_date_exception': 'Kérlek, rendezd a napi kiadásokat tartalmazó lapokat balról jobbra\n'
                                 'növekvő sorrendbe, és válaszd a "Táblázat adatok frissítése" gombot!\n',
          'update_sheet_list': 'A lapok adatait tartalmazó fájl frissítése, kérlek várj  ...\n'
                               '    Ha ezen folyamat közben történik valami, indítsd el újra a programot,\n'
                               '    és válaszd az "Táblázat adatok frissítése" gombot.\n'
                               '    Lehetséges, hogy csak 1 perc várakozás után fog újra működni a program.\n',
          'edit_mounthly_costs_sheet': '    Ha ezen folyamat közben történik valami, indítsd el újra a programot,\n'
                                       '    és válaszd a "Kiadás összesítő frissítése", majd a \n'
                                       '                 "Megtakarítás összesítő frissítése" gombokat.\n'
                                       '    Lehetséges, hogy csak 1 perc várakozás után fog \n'
                                       '    újra működni a program.\n',
          'edit_savings_sheet': '    Ha ezen folyamat közben történik valami, válaszd a \n'
                                '               "Megtakarítás összesítő frissítése" gombot.\n'
                                '    Lehetséges, hogy csak 1 perc várakozás után fog újra működni a program.\n',
          'edit_new_month': 'Új hónap hozzádása folyamatban  ...\n'
                            '  - Ha a program sikeresen hozzáadta az új hónapot, de utána hiba üzenetet küld,\n'
                            '      várj egy percet és válaszd a "Kiadás összesítő frissítése", majd a \n'
                            '                                   "Megtakarítás összesítő frissítése" gombokat.\n'
                            '   - Ha nem tudta végigcsinálni ezt a folyamatot, nézd meg az interneten, \n'
                            '     hogy létre jött-e új hónap. Ha igen, töröld ki és add hozzá újra a honapot.\n'
                            '   Lehet hogy várni kell 1 percet.\n'}

