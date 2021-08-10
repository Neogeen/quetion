import os
import pandas as pd
import shutil

spisok = os.listdir(r'C:\Users\new2r\Desktop\endss\end')

for i in spisok:
    print(i)

# table = pd.read_excel('Bs.xlsx', sheet_name='Отделить')
#
# for i in spisok:
#     j = i.split(' ')
#     for numb in range(406):
#         ab = j[0]+" "+j[1]
#         if ab == table.loc[numb]['bs']:
#             shutil.copy(f'C:\\Users\\new2r\\Desktop\\endss\\1\\{i}', r'C:\Users\new2r\Desktop\endss\end')
#             break