import datetime
import numpy as np
import pandas as pd
import re
import string
from pandas.tseries.api import guess_datetime_format

from IndexColumnConverter import IndexColumnConverter

CATEGORIES = ['Integer', 'Float', 'Percentage', 'Scientific Notation', 'Date',
              'Time', 'Currency', 'Email', 'Other']
K = 4

class SheetCompressor:
    def __init__(self):
        self.row_candidates = []
        self.column_candidates = []
        self.row_lengths = {}
        self.column_lengths = {}

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
        converter = IndexColumnConverter()
        markdown = pd.DataFrame(columns = ['Address', 'Value', 'Format'])
        for rowindex, i in sheet.iterrows():
            for colindex, j in enumerate(sheet.columns.tolist()):
                new_row = pd.DataFrame([converter.parse_colindex(colindex + 1) + str(rowindex + 1), i[j],
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

    def anchor(self, sheet):
        
        #Given num, obtain all integers from num - k to num + k inclusive
        def surrounding_k(num, k):
            return list(range(num - k, num + k + 1))
        
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
        self.row_candidates = np.unique(list(np.concatenate([surrounding_k(i, K) for i in self.row_candidates]).flat))
        self.column_candidates = np.unique(list(np.concatenate([surrounding_k(i, K) for i in self.column_candidates]).flat))

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

        #Takes array of Excel cells and combines adjacent cells
        def combine_cells(array):
            
            # Correct version
            # 2d version of summary ranges from leetcode
            # For each row, run summary ranges to get a 1d array, then run summary ranges for each column 

            # Greedy version
            if len(array) == 1:
                return array[0]
            return array[0] + ':' + array[-1]
        
        dictionary = {}
        for _, i in markdown.iterrows():
            if i['Value'] in dictionary:
                dictionary[i['Value']].append(i['Address'])
            else:
                dictionary[i['Value']] = [i['Address']]
        dictionary = {k: v for k, v in dictionary.items() if not pd.isna(k)}
        dictionary = {k: combine_cells(v) for k, v in dictionary.items()}
        return dictionary
    
    #Key-Value to Value-Key for categories
    def inverted_category(self, markdown):
        dictionary = {}
        for _, i in markdown.iterrows():
                dictionary[i['Value']] = i['Category']
        return dictionary
    
    #Regex to NFS
    def get_category(self, string):
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

        #Handles nan edge cases
        def replace_nan(sheet):
            if pd.isna(sheet):
                return 'Other'
            else:
                return dictionary[sheet]

        #DFS for checking bounds
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

        m = len(sheet)
        n = len(sheet.columns)

        visited = [[False] * n for _ in range(m)]
        areas = []

        for r in range(m):
            for c in range(n):
                if not visited[r][c]:
                    val_type = replace_nan(sheet.iloc[r, c])
                    bounds = dfs(r, c, val_type)
                    areas.append([(bounds[0], bounds[1]), (bounds[2], bounds[3]), val_type])
        return areas