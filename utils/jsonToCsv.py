# -*- coding: utf-8 -*-
"""
Created on Wed Mar 25 08:50:28 2020

@author: RROJASQ
"""

import json
import csv
import pandas as pd
import numpy as np
#https://dev.azure.com/grupoepm/_apis/projects?api-version=5.1
df = pd.read_json (r'D:\rrojasq\OneDrive - Grupo EPM\Descargas\AzDevOpsProjects.json')

#columns_csv = ['',]
columns = ['name', 'description', 'url']
df_csv = pd.DataFrame(columns=columns)
for serie_data in df['value']:
    print(serie_data['name'], ' - ', serie_data['description'])
    df_new = pd.DataFrame([[serie_data['name'],serie_data['description'],serie_data['url']]], columns=columns)
    df_csv = df_csv.append(df_new)
    #df_data = pd.DataFrame.from_dict(serie_data)


df_csv.to_csv (r'D:\rrojasq\OneDrive - Grupo EPM\Descargas\AzDevOpsProjects.csv', index = None)