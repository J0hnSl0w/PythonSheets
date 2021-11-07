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

values = {'mc_value': "=SUM(SUMIF('{date}'!$C$4:$C$100;'{sheet_name}'!$A2;'{date}'!$A$4:$A$100);SUMIF('{date}'!$G$4:$G$100;'{sheet_name}'!$A2;'{date}'!$E$4:$E$100))",
          'starting_otp': ['A3', "='{date}'!A2"],
          'starting_erste': ['E3', "='{date}'!E2"],
          'bank_account': ['H2', "=$A$2+'Megtakarítás részletező {date}'!N38"],
          'date': [13, 'E', 3]
}
