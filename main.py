import os
import pandas as pd
import xlrd

from SheetCompressor import SheetCompressor

DIRECTORY = 'VFUSE'   

sheet_compressor = SheetCompressor()

def main():
    for root, dirs, files in os.walk(DIRECTORY):
        for file in files:
            if file == 'readme.txt':
                continue
            try:
                wb = xlrd.open_workbook(os.path.join(root, file), logfile=open(os.devnull,'w'), formatting_info=True)
            except xlrd.biffh.XLRDError:
                return 
            sheet = pd.read_excel(wb, engine='xlrd')
            sheet = sheet.apply(lambda x: x.str.replace('\n', '<br>') if x.dtype == 'object' else x)

            sheet = sheet_compressor.anchor(sheet) #3.2
            markdown = sheet_compressor.encode(wb, sheet) #3.1

            markdown['Category'] = markdown['Value'].apply(lambda x: sheet_compressor.category(x)) #Get categories (part of 3.4)

            dictionary = sheet_compressor.inverted_index(markdown) #3.3

            print(dictionary)

            #Write the dictionary
            

if __name__ == "__main__":
    main()