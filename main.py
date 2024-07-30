import os
import pandas as pd
import xlrd

from SheetCompressor import SheetCompressor

DIRECTORY = 'VFUSE'   

def main():
    for root, dirs, files in os.walk(DIRECTORY):
        for file in files:
            sheet_compressor = SheetCompressor()
            if file == 'readme.txt':
                continue
            try:
                wb = xlrd.open_workbook(os.path.join(root, file), logfile=open(os.devnull,'w'), formatting_info=True)
            except xlrd.biffh.XLRDError:
                return 
            sheet = pd.read_excel(wb, engine='xlrd')
            sheet = sheet.apply(lambda x: x.str.replace('\n', '<br>') if x.dtype == 'object' else x)
            sheet = sheet.T.reset_index().T.reset_index().drop(columns = 'index')

            #Structural-anchor-based Extraction
            sheet = sheet_compressor.anchor(sheet)

            #Encoding 
            markdown = sheet_compressor.encode(wb, sheet) #Paper encodes first then anchors; I chose to do this in reverse

            #Data-Format Aggregation
            markdown['Category'] = markdown['Value'].apply(lambda x: sheet_compressor.category(x))
            category_dict = sheet_compressor.inverted_category(markdown) 
            try:
                areas = sheet_compressor.identical_cell_aggregation(sheet, category_dict)
            except RecursionError:
                continue

            #Inverted-index Translation
            compress_dict = sheet_compressor.inverted_index(markdown)

            #Write to files
            sheet_compressor.write_areas('output/' + file.split('.')[0] + '_areas.txt', areas)
            sheet_compressor.write_dict('output/' + file.split('.')[0] + '_dict.txt', compress_dict)

if __name__ == "__main__":
    main()