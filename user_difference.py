import os
import pandas as pd

BASE_DIR = os.getcwd()

#a = pd.read_csv(os.path.join(BASE_DIR,'current_csv_files\\users_final.csv'))
#b = pd.read_csv(os.path.join(BASE_DIR,'previous_csv_files\\users_final.csv'))
a = pd.read_csv(os.path.join(BASE_DIR,'current_csv_files\\users_final.csv'))
b = pd.read_csv(os.path.join(BASE_DIR,'previous_csv_files\\users_final.csv'))
c = pd.concat([a,b], axis=0)

c.drop_duplicates(keep=False, inplace=True) # Set keep to False if you don't want any
                                              # of the duplicates at all
c.reset_index(drop=True, inplace=True)
# print(a)
# print(b)
print(c)

c.to_csv('user_changes_pd.csv', index_label=False, index=False)