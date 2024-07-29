import pandas as pd

DIRECTORY = 'VFUSE'

sheet = pd.read_csv('temp.csv', sep='|')

#Key-Value -> Value-Key

dictionary = {}

for _, i in sheet.iterrows():
    if i['Value'] in dictionary:
        dictionary[i['Value']].append(i['Address'])
    else:
        dictionary[i['Value']] = [i['Address']]

print(dictionary)