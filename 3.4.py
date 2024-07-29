import pandas as pd
from pandas.tseries.api import guess_datetime_format
import re

DIRECTORY = 'VFUSE'
CATEGORIES = ['Integer', 'Float', 'Percentage', 'Scientific Notation', 'Date',
              'Time', 'Currency', 'Email', 'Other']

#sheet = pd.read_excel(DIRECTORY + '/1/7b5a0a10-e241-4c0d-a896-11c7c9bf2040.xls', engine='xlrd')
sheet = pd.read_csv('temp.csv', sep='|')

#Regex courtesy of regexlib.com and stackoverflow
def category(string):
    if pd.isna(string):
        return 'Other'
    if re.match('^(\+|-)?\d+$', string) or re.match('^\d{1,3}(,\d{1,3})*$', string): #Steven Smith
        return 'Integer'
    if re.match('^[-+]?\d*\.?\d*$', string) or re.match('^\d{1,3}(,\d{3})*(\.\d+)?$', string): #Steven Smith/Stack Overflow (https://stackoverflow.com/questions/5917082/regular-expression-to-match-numbers-with-or-without-commas-and-decimals-in-text)
        return 'Float'
    if re.match('^[-+]?\d*\.?\d*%$', string) or re.match('^\d{1,3}(,\d{3})*(\.\d+)?%$', string):
        return 'Percentage'
    if re.match('^[-+]?[$]\d*\.?\d{2}$', string) or re.match('^\d{1,3}(,\d{3})*(\.\d{2})?$', string): #Michael Ash
        return 'Currency'
    if re.match('\b-?[1-9](?:\.\d+)?[Ee][-+]?\d+\b', string): #Michael Ash
        return 'Scientific Notation'
    if re.match("^((([!#$%&'*+\-/=?^_`{|}~\w])|([!#$%&'*+\-/=?^_`{|}~\w][!#$%&'*+\-/=?^_`{|}~\.\w]{0,}[!#$%&'*+\-/=?^_`{|}~\w]))[@]\w+([-.]\w+)*\.\w+([-.]\w+)*)$", string): #Dave Black RFC 2821
        return 'Email'
    if datetime := guess_datetime_format(string):
        return datetime
    return 'Other'

sheet['Category'] = sheet['Value'].apply(lambda x: category(x))

#Algorithm 1

m = len(sheet)
n = len(sheet.columns)

visited = [[False] * n] * m

areas = []

FormatDict = {}

def dfs(r, c, val_type):
    if visited[r][c] or val_type != FormatDict[sheet[r, c]]:
        return [r, c, r - 1, c - 1]
    visited[r][c] = True
    bounds = [r, c, r, c]
    for i in [[r - 1, c], [r, c - 1], [r + 1, c], [c, r + 1]]:
        if not visited[i] and val_type == FormatDict[sheet[i]]:
            new_bounds = dfs(i[0], i[1], val_type)
            bounds = [min(new_bounds[0], bounds[0]), min(new_bounds[1], bounds[1]), max(new_bounds[2], bounds[2]), max(new_bounds[3], bounds[3])]
    return bounds

for r in range(m):
    for c in range(n):
        if not visited[r][c]:
            val_type = FormatDict[sheet[r, c]]
            bounds = dfs(r, c, val_type)
            areas += [(bounds[0], bounds[1]), (bounds[2], bounds[3]), val_type]

print(areas)