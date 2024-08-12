import argparse
import os
import pandas as pd
import string
import xlrd

from SheetCompressor import SheetCompressor
from SpreadsheetLLM import SpreadsheetLLM

original_size = 0
new_size = 0

#Takes a file, compresses it, and writes to output folder
def compress_spreadsheet(file):
    sheet_compressor = SheetCompressor()
    if file == 'readme.txt':
        return
    try:
        wb = xlrd.open_workbook(os.path.join(root, file), logfile=open(os.devnull,'w'), formatting_info=True)
    except xlrd.biffh.XLRDError:
        return 
    sheet = pd.read_excel(wb, engine='xlrd')
    sheet = sheet.apply(lambda x: x.str.replace('\n', '<br>') if x.dtype == 'object' else x)

    #Move columns to row 1
    sheet.loc[-1] = sheet.columns
    sheet.index += 1
    sheet.sort_index(inplace=True)
    sheet.columns = list(range(len(sheet.columns)))

    #Structural-anchor-based Extraction
    sheet = sheet_compressor.anchor(sheet)

    #Encoding 
    markdown = sheet_compressor.encode(wb, sheet) #Paper encodes first then anchors; I chose to do this in reverse

    #Data-Format Aggregation
    markdown['Category'] = markdown['Value'].apply(lambda x: sheet_compressor.get_category(x))
    category_dict = sheet_compressor.inverted_category(markdown) 
    try:
        areas = sheet_compressor.identical_cell_aggregation(sheet, category_dict)
    except RecursionError:
        return

    #Inverted-index Translation
    compress_dict = sheet_compressor.inverted_index(markdown)

    return areas, compress_dict

def llm(model, filename):
    spreadsheet_llm = SpreadsheetLLM(model)
    
    with open('output/' + filename + '_areas.txt') as f:
        area = f.readlines()
    with open('output/' + filename + '_dict.txt') as f:
        table = f.readlines()
    if args.table:
        print(spreadsheet_llm.identify_table(area))
    if args.question:
        print(spreadsheet_llm.question_answer(table, args.question))

#Converts index to column letter; courtesy of https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
def parse_colindex(num):
    
    #Modified divmod function for Excel; courtesy of https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
    def divmod_excel(n):
        a, b = divmod(n, 26)
        if b == 0:
            return a - 1, b + 26
        return a, b
    
    chars = []
    while num > 0:
        num, d = divmod_excel(num)
        chars.append(string.ascii_uppercase[d - 1])
    return ''.join(reversed(chars))
    
def write_areas(file, areas):
    string = ''
    for i in areas:
        string += ('(' + i[2] + '|' + parse_colindex(i[0][1] + 1) + str(i[0][0] + 1) + ':' 
                    + parse_colindex(i[1][1] + 1) + str(i[1][0] + 1) + '), ')
    with open(file, 'w+', encoding="utf-8") as f:
        f.writelines(string)

def write_dict(file, dict):
    string = ''
    for key, value in dict.items():
        string += (str(value) + ',' + str(key) + '|')
    with open(file, 'w+', encoding="utf-8") as f:
        f.writelines(string)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--compress', action=argparse.BooleanOptionalAction, default=True, help="compress dataset into txt files; must run for LLM to work")
    parser.add_argument('--directory', type=str, default='VFUSE', help='directory of excel files')
    parser.add_argument('--file', type=str, default='7b5a0a10-e241-4c0d-a896-11c7c9bf2040', help='file to work with')
    parser.add_argument('--model', type=str, choices={'gpt-3.5', 'gpt-4', 'mistral', 'llama-2', 'llama-3', 'phi-3'}, default='gpt-4', help='llm to use')
    parser.add_argument('--table', action=argparse.BooleanOptionalAction, default=True, help='Whether or not to identify number of tables')
    parser.add_argument('--question', type=str, help='question to ask llm')
    args = parser.parse_args()

    if args.compress:
        for root, dirs, files in os.walk(args.directory):
            for file in files:
                try:
                    areas, compress_dict = compress_spreadsheet(file)
                except TypeError:
                    continue
                write_areas('output/' + file.split('.')[0] + '_areas.txt', areas)
                write_dict('output/' + file.split('.')[0] + '_dict.txt', compress_dict)
                original_size += os.path.getsize(os.path.join(root, file))
                new_size += os.path.getsize('output/' + file.split('.')[0] + '_areas.txt')
                new_size += os.path.getsize('output/' + file.split('.')[0] + '_dict.txt')
        print('Compression Ratio: {}'.format(str(original_size / new_size)))
    llm(args.model, args.file)