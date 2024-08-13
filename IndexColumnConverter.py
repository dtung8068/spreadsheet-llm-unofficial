import string

class IndexColumnConverter():

    def __init__(self):
        return
    
    #Converts index to column letter; courtesy of https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
    def parse_colindex(self, num):
        
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