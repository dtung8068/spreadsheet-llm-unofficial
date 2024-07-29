#May require "Supplemental Materials"

#"Novel Heuristic Method (Appendix C)"

#Step 1: Create boundary lines
    #Iterate through data, look for differences in data type? (Maybe stick with NaN for now for simplicity)
    #Initially start with first row/column and last row/column
#Step 2: Use any 2 rows and columns as top/bottom/left/right boundaries
    # Exhaustive Search? Sounds like a lot...
#Step 3: Filter unreasonable boundary candidates
    # No thresholds published
#Step 4: Pairwise enumeration of overlapping candidates
    # Check for headers/candidate boundaries

import pandas as pd
import numpy as np

K = 4

#sheet = pd.read_csv('temp.csv', sep='|')
sheet = pd.read_excel('VFUSE/1/7b5a0a10-e241-4c0d-a896-11c7c9bf2040.xls')

#Step 1

row_candidates = []
column_candidates = []

current_type = []
#Iterate through rows
for i, j in sheet.iterrows():
    if current_type != (temp := j.apply(type).to_list()):
        current_type = temp
        row_candidates.append(i)

#Iterate through columns

current_type = []
for i, j in enumerate(sheet.columns):
    if current_type != (temp := sheet[j].apply(type).to_list()):
        current_type = temp
        column_candidates.append(i)

#Step 2
#candidates = [x1, x2, y1, y2]
candidates = [(a, b, c, d) for idx, a in enumerate(row_candidates) for b in row_candidates[idx + 1:] for idy, c in enumerate(column_candidates) for d in column_candidates[idy + 1:]]

#Step 3

row_lengths = {}
column_lengths = {}

for i, j in sheet.iterrows():
    row_lengths[i] = sum(j.apply(lambda x: 0 if isinstance(x, float) or isinstance(x, int) else len(x)))
mean = np.mean(list(row_lengths.values()))
std = np.std(list(row_lengths.values()))
min = np.max(mean - 2 * std, 0)
max = mean + 2 * std
row_lengths = dict((k, v) for k, v in row_lengths.items() if v < min or v > max)
row_lengths[0] = 0
row_lengths[len(sheet) - 1] = 0

for i, j in enumerate(sheet.columns):
    column_lengths[i] = sum(sheet[j].apply(lambda x: 0 if isinstance(x, float) or isinstance(x, int) else len(x)))
mean = np.mean(list(column_lengths.values()))
std = np.std(list(column_lengths.values()))
min = np.max(mean - 2 * std, 0)
max = mean + 2 * std
column_lengths = dict((k, v) for k, v in column_lengths.items() if v < min or v > max)
column_lengths[0] = 0
column_lengths[len(sheet.columns) - 1] = 0

def surrounding_k(num, k):
    return list(range(num - k, num + k + 1))

#Step 4
row_lengths = np.unique(list(np.concatenate([surrounding_k(i, K) for i in row_lengths]).flat)) #
column_lengths = np.unique(list(np.concatenate([surrounding_k(i, K) for i in column_lengths]).flat))

row_lengths = row_lengths[(row_lengths >= 0) & (row_lengths < len(sheet))]
column_lengths = column_lengths[(column_lengths >= 0) & (column_lengths < len(sheet.columns))]

sheet = sheet.iloc[row_lengths, column_lengths]
sheet.to_csv('temp.csv')