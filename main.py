import pandas as pd
import numpy
import statistics


def Reverse(lst):
    return [ele for ele in reversed(lst)]


asas_df = pd.read_excel(r'C:\Users\Amgh\Desktop\portal.xlsx', sheet_name='آساس', usecols=[8])
price_asas = asas_df.values.tolist()
flat_asas = []
for sublist in price_asas:
    for item in sublist:
        flat_asas.append(item)

asas_final = Reverse(flat_asas)

asas_month = asas_final[0:29]
