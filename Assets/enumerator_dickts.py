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

values = {
    'mc_value': "=SUM(SUMIF('{date}'!$C$4:$C$100;'{sheet_name}'!$A2;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'{sheet_name}'!$A2;'{date}'!$E$4:$E$100))",
    'starting_kp': ['A3', "='{prev_date}'!A2"],
    'starting_b': ['E3', "='{prev_date}'!E2"],
    'savings_1': ['K2', "=-(SUM(SUMIF('{date}'!$C$4:$C$100;'Havi kiadások {year}'!$A5;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'Havi kiadások {year}'!$A5;'{date}'!$E$4:$E$100)))"],
    'savings_2': ['K3', "=-(SUM(SUMIF('{date}'!$C$4:$C$100;'Havi kiadások {year}'!$A6;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'Havi kiadások {year}'!$A6;'{date}'!$E$4:$E$100)))"],
    'savings_3': ['K4', "=SUM(K2:K3)+'{prev_date}'!K4"],
    'uribol_1': ['K7', "=-(SUM(SUMIF('{date}'!$C$4:$C$100;'Havi kiadások {year}'!$A7;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'Havi kiadások {year}'!$A7;'{date}'!$E$4:$E$100)))"],
    'uribol_2': ['K8', "=-(SUM(SUMIF('{date}'!$C$4:$C$100;'Havi kiadások {year}'!$A8;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'Havi kiadások {year}'!$A8;'{date}'!$E$4:$E$100)))"],
    'uribol_3': ['K9', "=SUM(K7:K8)+'{prev_date}'!K9"],
}

text_spec = {'bg': '#35455D',
             'fg': '#BFD1DF',
             'width': 50,
             'height': 2,
             'font_c1': 'Sans 15 bold',
             'font_c2': 'Sans 12'}

buttons_spec = {'bg': '#92B1B6',
                'fg': 'black',
                'widht': 30,
                'height': 2,
                'font': 'Sans 10'}

frame_spec = {'bg': '#35455D'}

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