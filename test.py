import pandas as pd

def get_mpc(fraction, element, value):
    mpc = {
        'песок/суспесь': {
            'pH': None, 
            'benz': 0.02,
            'Cu': 33,
            'Zn': 55,
            'Pb': 32,
            'Cd': 0.5,
            'Ni': 20,
            'As': 2,
            'Hg': 2.1,
            'oil': 1,     
        },
        'суглинок, pH<=5,5': {
            'pH': None, 
            'benz': 0.02,
            'Cu': 66,
            'Zn': 110,
            'Pb': 65,
            'Cd': 1,
            'Ni': 40,
            'As': 5,
            'Hg': 2.1,
            'oil': 1,     
        },
        'суглинок, pH>5,5': {
            'pH': None, 
            'benz': 0.02,
            'Cu': 132,
            'Zn': 220,
            'Pb': 130,
            'Cd': 2.0,
            'Ni': 80,
            'As': 10,
            'Hg': 2.1,
            'oil': 1,     
        }
    }

def get_pv(element, value):    
    bv = {
        'pH': None, 
        'benz': None,
        'Cu': 18,
        'Zn': 43.10,
        'Pb': 19.11,
        'Cd': 0.17,
        'Ni': 15.3,
        'As': 2.62,
        'Hg': 0.03,
        'oil': None,     
        }







def get_source():

    df = pd.read_excel('исходник.xlsx', dtype=str, skiprows=12, header=None)

    df = df.rename(
        columns = {
            0: 'test_number',
            1: 'poligon_number',
            2: 'sample_number',
            3: 'pH',
            4: 'benz',
            5: 'Cu',
            6: 'Zn',
            7: 'Pb',
            8: 'Cd',
            9: 'Ni',
            10: 'As',
            11: 'Hg',
            12: 'oil',
            13: 'fraction'
            }
        )


    df['depth'] = df.apply(lambda row: row['test_number'] if row['poligon_number']!=row['poligon_number'] else None, axis=1)
    df['depth']= df['depth'].ffill()
    df['depth']= df.apply(lambda row: f"{row['depth'].replace('Глубина отбора образцов, м: ', '')}\n({row['fraction']})" , axis=1)
    df = df[df['poligon_number']== df['poligon_number']]

    return df




def get_result(source_df):

    res_list = []
    for sr in source_df.itertuples():
        print(sr)
          
        res_row_1 = {
            0: f"Пробная площадка № {sr.poligon_number} (ПП{sr.poligon_number})",
            2: None,
            3: None,
            4: None,
            5: None,
            6: None,
            7: None,
            8: None,
            9: None,
            10: None,
            11: None,
            12: None,
            13: None,
            14: None
        }

        res_row_2 = {
            0: sr.test_number,
            2: f"{sr.depth}\n{sr.fraction}",
            3: 'Сi,мг/кг',
            4: sr.pH,
            5: sr.benz,
            6: sr.Cu,
            7: sr.Zn,
            8: sr.Pb,
            9: sr.Cd,
            10: sr.Ni,
            11: sr.As,
            12: sr.Hg,
            13: sr.oil,
            14: None
        }

        res_row_3 = {
            0: sr.test_number,
            2: f"{sr.depth}\n{sr.fraction}",
            3: 'Сi,мг/кг',
            4: sr.pH,
            5: sr.benz,
            6: sr.Cu,
            7: sr.Zn,
            8: sr.Pb,
            9: sr.Cd,
            10: sr.Ni,
            11: sr.As,
            12: sr.Hg,
            13: sr.oil,
            14: None
        }

        df = pd.DataFrame([res_row_1, res_row_2])

        print(df)

        res_list.append(df)
    
    res_df = pd.concat(res_list)

    return res_df    

if __name__ == '__main__':
    source_df = get_source()
    result_df = get_result(source_df) 
    print(result_df)