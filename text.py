import pandas as pd
import numpy as np
data_names = pd.read_excel(r'C:\Users\xcxhy\Desktop\company.xlsx')
data_names = data_names.iloc[:,1]
data_names = data_names.values
print(data_names,type(data_names),data_names.shape,type(data_names[3000]))