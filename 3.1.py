import os
import pandas as pd
import xlrd
import string

DIRECTORY = 'VFUSE'

file_dict = {}

#colindex conversion from here: https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26

def divmod_excel(n):
    a, b = divmod(n, 26)
    if b == 0:
        return a - 1, b + 26
    return a, b

def parse_colindex(num):
    chars = []
    while num > 0:
        num, d = divmod_excel(num)
        chars.append(string.ascii_uppercase[d - 1])
    return ''.join(reversed(chars))

def get_format(xf):
    format_array = []

    #Border
    if xf.border.bottom_line_style:
        format_array.append('Bottom Border')    

    #Fill
    if xf.background.background_colour_index != 65:
        format_array.append('Fill Color')

    #Bold
    if wb.font_list[xf.font_index].bold:
        format_array.append('Font Bold')
    
    return format_array

for root, dirs, files in os.walk(DIRECTORY):
    for file in files:
        if file == 'readme.txt':
            continue
        markdown = pd.DataFrame(columns = ['Address', 'Value', 'Format'])
        try:
            wb = xlrd.open_workbook(os.path.join(root, file), logfile=open(os.devnull,'w'), formatting_info=True)
        except xlrd.biffh.XLRDError:
            continue
        sheet = pd.read_excel(wb, engine='xlrd')
        sheet = sheet.apply(lambda x: x.str.replace('\n', '<br>') if x.dtype == 'object' else x)
        for rowindex, i in sheet.iterrows():
            for colindex, j in enumerate(sheet.columns.tolist()):
                new_row = pd.DataFrame([parse_colindex(colindex + 1) + str(rowindex + 2), i[j],
                                        get_format(wb.xf_list[wb.sheet_by_index(0).cell(rowindex, colindex).xf_index])]).T
                new_row.columns = markdown.columns
                markdown = pd.concat([markdown, new_row])                
        markdown.to_csv('temp.csv', sep='|', index=False)
        exit(1)