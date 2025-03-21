import os
import pandas as pd

data_file_folder = r'D:\hsg11'

df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='Sheet', header=None))

df_master = pd.concat(df, axis=0, ignore_index=True)

df_master.to_excel('final.xlsx', index=False, header=False)

print('Complete!')

