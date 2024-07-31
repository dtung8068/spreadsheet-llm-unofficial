import datetime
import numpy as np
import pandas as pd
import re
import string
from pandas.tseries.api import guess_datetime_format

CATEGORIES = ['Integer', 'Float', 'Percentage', 'Scientific Notation', 'Date',
              'Time', 'Currency', 'Email', 'Other']
K = 4

class SheetCompressor:
    def __init__(self):
        self.row_candidates = []
        self.column_candidates = []
        self.row_lengths = {}
        self.column_lengths = {}

    #Modified divmod function for Excel; courtesy of https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
    def divmod_excel(self, n):
        a, b = divmod(n, 26)
        if b == 0:
            return a - 1, b + 26
        return a, b

    #Converts index to column letter; courtesy of https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
    def parse_colindex(self, num):
        chars = []
        while num > 0:
            num, d = self.divmod_excel(num)
            chars.append(string.ascii_uppercase[d - 1])
        return ''.join(reversed(chars))

    #Obtain border, fill, bold info about cell; incomplete
    def get_format(self, xf, wb):
        format_array = []

        #Border
        if xf.border.top_line_style:
            format_array.append('Top Border')
        
        if xf.border.bottom_line_style:
            format_array.append('Bottom Border') 

        if xf.border.left_line_style:
            format_array.append('Left Border')
    
        if xf.border.right_line_style:
            format_array.append('Right Border')

        #Fill
        if xf.background.background_colour_index != 65:
            format_array.append('Fill Color')

        #Bold
        if wb.font_list[xf.font_index].bold:
            format_array.append('Font Bold')
        
        return format_array

    #Encode spreadsheet into markdown format
    def encode(self, wb, sheet):
        markdown = pd.DataFrame(columns = ['Address', 'Value', 'Format'])
        for rowindex, i in sheet.iterrows():
            for colindex, j in enumerate(sheet.columns.tolist()):
                new_row = pd.DataFrame([self.parse_colindex(colindex + 1) + str(rowindex + 1), i[j],
                                        self.get_format(wb.xf_list[wb.sheet_by_index(0).cell(rowindex, colindex).xf_index], wb)]).T
                new_row.columns = markdown.columns
                markdown = pd.concat([markdown, new_row])
        return markdown
    
    #Checks for identical dtypes across row/column
    def get_dtype_row(self, sheet):
        current_type = []
        for i, j in sheet.iterrows():
            if current_type != (temp := j.apply(type).to_list()):
                current_type = temp
                self.row_candidates.append(i)
    
    def get_dtype_column(self, sheet):
        current_type = []
        for i, j in enumerate(sheet.columns):
            if current_type != (temp := sheet[j].apply(type).to_list()):
                current_type = temp
                self.column_candidates.append(i)
    
    #Checks for length of text across row/column, looks for outliers, marks as candidates
    def get_length_row(self, sheet):
        for i, j in sheet.iterrows():
            self.row_lengths[i] = sum(j.apply(lambda x: 0 if isinstance(x, float) or isinstance(x, int)
                                              or isinstance(x, datetime.datetime) else len(x)))
        mean = np.mean(list(self.row_lengths.values()))
        std = np.std(list(self.row_lengths.values()))
        min = np.max(mean - 2 * std, 0)
        max = mean + 2 * std
        self.row_lengths = dict((k, v) for k, v in self.row_lengths.items() if v < min or v > max)

    def get_length_column(self, sheet):
        for i, j in enumerate(sheet.columns):
            self.column_lengths[i] = sum(sheet[j].apply(lambda x: 0 if isinstance(x, float) or isinstance(x, int)
                                                        or isinstance(x, datetime.datetime) else len(x)))
        mean = np.mean(list(self.column_lengths.values()))
        std = np.std(list(self.column_lengths.values()))
        min = np.max(mean - 2 * std, 0)
        max = mean + 2 * std
        self.column_lengths = dict((k, v) for k, v in self.column_lengths.items() if v < min or v > max)
    
    #Given num, obtain all integers from num - k to num + k inclusive
    def surrounding_k(self, num, k):
        return list(range(num - k, num + k + 1))

    def anchor(self, sheet):
        self.get_dtype_row(sheet)
        self.get_dtype_column(sheet)
        self.get_length_row(sheet)
        self.get_length_column(sheet)

        #Keep candidates found in both dtype/length method
        self.row_candidates = np.intersect1d(list(self.row_lengths.keys()), self.row_candidates)
        self.column_candidates = np.intersect1d(list(self.column_lengths.keys()), self.column_candidates)

        #Beginning/End are candidates
        self.row_candidates = np.append(self.row_candidates, [0, len(sheet) - 1]).astype('int32')
        self.column_candidates = np.append(self.column_candidates, [0, len(sheet.columns) - 1]).astype('int32')

        #Get K closest rows/columns to each candidate
        self.row_candidates = np.unique(list(np.concatenate([self.surrounding_k(i, K) for i in self.row_candidates]).flat))
        self.column_candidates = np.unique(list(np.concatenate([self.surrounding_k(i, K) for i in self.column_candidates]).flat))

        #Truncate negative/out of bounds
        self.row_candidates = self.row_candidates[(self.row_candidates >= 0) & (self.row_candidates < len(sheet))]
        self.column_candidates = self.column_candidates[(self.column_candidates >= 0) & (self.column_candidates < len(sheet.columns))]

        sheet = sheet.iloc[self.row_candidates, self.column_candidates]

        #Remap coordinates
        sheet = sheet.reset_index().drop(columns = 'index')
        sheet.columns = list(range(len(sheet.columns)))

        return sheet
    
    #Converts markdown to value-key pair
    def inverted_index(self, markdown):
        dictionary = {}
        for _, i in markdown.iterrows():
            if i['Value'] in dictionary:
                dictionary[i['Value']].append(i['Address'])
            else:
                dictionary[i['Value']] = [i['Address']]
        dictionary = {k: v for k, v in dictionary.items() if not pd.isna(k)}
        return dictionary
    
    def inverted_category(self, markdown):
        dictionary = {}
        for _, i in markdown.iterrows():
                dictionary[i['Value']] = i['Category']
        return dictionary
    
    #Regex to NFS
    def category(self, string):
        if pd.isna(string):
            return 'Other'
        if isinstance(string, float):
            return 'Float'
        if isinstance(string, int):
            return 'Integer'
        if isinstance(string, datetime.datetime):
            return 'yyyy/mm/dd'
        if re.match('^(\+|-)?\d+$', string) or re.match('^\d{1,3}(,\d{1,3})*$', string): #Steven Smith
            return 'Integer'
        if re.match('^[-+]?\d*\.?\d*$', string) or re.match('^\d{1,3}(,\d{3})*(\.\d+)?$', string): #Steven Smith/Stack Overflow (https://stackoverflow.com/questions/5917082/regular-expression-to-match-numbers-with-or-without-commas-and-decimals-in-text)
            return 'Float'
        if re.match('^[-+]?\d*\.?\d*%$', string) or re.match('^\d{1,3}(,\d{3})*(\.\d+)?%$', string):
            return 'Percentage'
        if re.match('^[-+]?[$]\d*\.?\d{2}$', string) or re.match('^[-+]?[$]\d{1,3}(,\d{3})*(\.\d{2})?$', string): #Michael Ash
            return 'Currency'
        if re.match('\b-?[1-9](?:\.\d+)?[Ee][-+]?\d+\b', string): #Michael Ash
            return 'Scientific Notation'
        if re.match("^((([!#$%&'*+\-/=?^_`{|}~\w])|([!#$%&'*+\-/=?^_`{|}~\w][!#$%&'*+\-/=?^_`{|}~\.\w]{0,}[!#$%&'*+\-/=?^_`{|}~\w]))[@]\w+([-.]\w+)*\.\w+([-.]\w+)*)$", string): #Dave Black RFC 2821
            return 'Email'
        if datetime_format := guess_datetime_format(string):
            return datetime_format
        return 'Other'
    
    def identical_cell_aggregation(self, sheet, dictionary):

        def replace_nan(sheet):
            if pd.isna(sheet):
                return 'Other'
            else:
                return dictionary[sheet]
            
        m = len(sheet)
        n = len(sheet.columns)

        visited = [[False] * n for _ in range(m)]
        areas = []
        def dfs(r, c, val_type):
            match = replace_nan(sheet.iloc[r, c])
            if visited[r][c] or val_type != match:
                return [r, c, r - 1, c - 1]
            visited[r][c] = True
            bounds = [r, c, r, c]
            for i in [[r - 1, c], [r, c - 1], [r + 1, c], [r, c + 1]]:
                if (i[0] < 0) or (i[1] < 0) or (i[0] >= len(sheet)) or (i[1] >= len(sheet.columns)):
                    continue
                match = replace_nan(sheet.iloc[i[0], i[1]])
                if not visited[i[0]][i[1]] and val_type == match: 
                    new_bounds = dfs(i[0], i[1], val_type)
                    bounds = [min(new_bounds[0], bounds[0]), min(new_bounds[1], bounds[1]), max(new_bounds[2], bounds[2]), max(new_bounds[3], bounds[3])]
            return bounds

        for r in range(m):
            for c in range(n):
                if not visited[r][c]:
                    val_type = replace_nan(sheet.iloc[r, c])
                    bounds = dfs(r, c, val_type)
                    areas.append([(bounds[0], bounds[1]), (bounds[2], bounds[3]), val_type])
        return areas
    
    def write_areas(self, file, areas):
        string = ''
        for i in areas:
            string += ('(' + i[2] + '|' + self.parse_colindex(i[0][1] + 1) + str(i[0][0]) + ':' 
                       + self.parse_colindex(i[1][1] + 1) + str(i[1][0]) + '), ')
        with open(file, 'w+', encoding="utf-8") as f:
            f.writelines(string)

    def write_dict(self, file, dict):
        string = ''
        for key, value in dict.items():
            for i in value:
                string += (i + ',' + str(key) + '|')
        with open(file, 'w+', encoding="utf-8") as f:
            f.writelines(string)
