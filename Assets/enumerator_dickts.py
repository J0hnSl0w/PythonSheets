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

values_mesi = {'mc_value': "=SUM(SUMIF('{date}'!$C$4:$C$100;'{sheet_name}'!$A2;'{date}'!$A$4:$A$100))",
               'sv_value': "=SUMIF('{date}'!$D$4:$D$100;'{sheet_name}'!$A2;'{date}'!$A$4:$A$100)*-1",
               'starting': ['A3', "='{date}'!A2"],
               'bank_account': ['H2', "=$A$2+'Megtakarítás részletező {date}'!N38"],
               'date': [13, 'E', 3]}

values_klarka = {'mc_value': "=SUM(SUMIF('{date}'!$C$4:$C$50;'{sheet_name}'!$A2;'{date}'!$A$4:$A$50);SUMIF('{date}'!$G$4:$G$50;'{sheet_name}'!$A2;'{date}'!$E$4:$E$41))",
                 'starting_kp': ['A3', "='{date}'!A2"], 'starting_b': ['E3', "='{date}'!E2"],
                 'savings_1': "=-(SUM(SUMIF('{date}'!$C$4:$C$100;'Havi kiadások {date}'!$A4;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'Havi kiadások {date}'!$A4;'{date}'!$E$4:$E$100)))",
                 'savings_2': "=-(SUM(SUMIF('{date}'!$C$4:$C$100;'Havi kiadások {date}'!$A5;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'Havi kiadások {date}'!$A5;'{date}'!$E$4:$E$100)))",
                 'savings_3': "=SUM(K2:K3)+'{date}'!K4",
                 'uribol_1': "=-(SUM(SUMIF('{date}'!$C$6:$C$100;'Havi kiadások {date}'!$A6;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'Havi kiadások {date}'!$A6;'{date}'!$E$4:$E$100)))",
                 'uribol_2': "=-(SUM(SUMIF('{date}'!$C$7:$C$100;'Havi kiadások {date}'!$A7;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'Havi kiadások {date}'!$A7;'{date}'!$E$4:$E$100)))",
                 'uribol_3': "=SUM(K2:K3)+'{date}'!K9",
                 'bank_account': ['H2', "=$A$2+'Megtakarítás részletező {date}'!N38"]}
