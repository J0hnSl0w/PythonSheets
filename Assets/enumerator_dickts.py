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