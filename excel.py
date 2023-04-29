import xlwings, re, numpy


class Excel_regex():

    def __init__(self) -> None:

        '''Инициализация базовых переменных'''

        self.keys_range =             []
        self.values_range =           []

        self.keys_sheet =             ''
        self.values_sheet =           ''

        self.row_list =               []
        self.column_list =            []

        self.key_list =               ''
        self.value_list =             ''

        self.keys_filepath =          ''
        self.values_filepath =        ''

        self.sequence =               ''

        self.row_count =              0
        self.column_count =           0

        self.row_ranges =            []
        self.row_seps =              []

        self.column_ranges =         []
        self.column_seps =           []

    def create_filters(self, row_ranges, row_seps, column_ranges,  column_seps):
        '''Конечная, тут по адресам считываются значения, создаются регулярные выражения,
        производится проверка, а потом выводится в новую книгу'''

        self.row_ranges = row_ranges
        self.row_seps = row_seps
        self.column_ranges = column_ranges
        self.column_seps = column_seps

        self.row_list =               []
        self.column_list =            []

        self.key_list =               ''
        self.value_list =             ''

        self.key_list = self.get_data(self.keys_filepath, self.keys_sheet, self.keys_range)
        self.value_list = self.get_data(self.values_filepath, self.values_sheet, self.values_range)

        for i in range(len(self.value_list)):
            for j in range(len(self.value_list[i])):
                self.value_list[i][j] = round(self.value_list[i][j], 2)

        for i in self.row_ranges:
            self.row_list.append(self.get_data(self.values_filepath, self.values_sheet, i))
        for i in self.column_ranges:
            self.column_list.append(self.get_data(self.values_filepath, self.values_sheet, i))

        #& Преобразование float в int и избавление от объединенных ячеек (можно назвать костылем)
        for i in range(len(self.row_list)):
            for j in range(len(self.row_list[i])):
                if self.row_list[i][j] == None:
                    if self.row_list[i][j-1] != None:
                        tmp = self.row_list[i][j-1]
                    else:
                        tmp = self.row_list[i][j+1]
                    
                    self.row_list[i][j] = tmp
                try:
                    self.row_list[i][j] = int(self.row_list[i][j])
                    # row_list[i][j] = round(float(row_list[i][j]), 2)
                except:
                    pass

        for i in range(len(self.column_list)):
            for j in range(len(self.column_list[i])):
                if self.column_list[i][j] == None:
                    if self.column_list[i][j-1] != None:
                        tmp = self.column_list[i][j-1]
                    else:
                        tmp = self.column_list[i][j+1]
                    
                    self.column_list[i][j] = tmp
                try:
                    self.column_list[i][j] = int(self.column_list[i][j])
                    # column_list[i][j] = round(float(column_list[i][j]), 2)
                except:
                    pass

        #& Разделение строк, сепарация
        for i in range(len(row_seps)):
            if row_seps[i] != '':
                tmp = []
                for j in self.row_list[i]:
                    tmp.append(j.split(row_seps[i]))
                tmp = numpy.array(tmp).transpose()
                self.row_list.pop(i)
                tmp = tmp.tolist()
                tmp.reverse()
                for n in tmp:
                    self.row_list.insert(i, n)
        
        for i in range(len(column_seps)):
            if column_seps[i] != '':
                tmp = []
                for j in self.column_list[i]:
                    tmp.append(j.split(column_seps[i]))
                tmp = numpy.array(tmp).transpose()
                self.column_list.pop(i)
                tmp = tmp.tolist()
                tmp.reverse()
                for n in tmp:
                    self.column_list.insert(i, n)

        return len(self.row_list) + len(self.column_list)
        
        # print('строки:',row_list, '\n\n', 'столбцы:',column_list)
        # print('\n\n\n', value_list, '\n\n\n', key_list)
        # return 'work!'

    def source_file_and_cells(self, keys_or_values: bool = True) -> list[str]:
        '''True for keys and False for values
            \nreturn:\n
            [0] - address range\n
            [1] - sheet name\n
            [2] - filepath\n'''

        tmp = ['', '', '']
        # tmp_address =   ''
        # tmp_sheet =     ''
        # tmp_filepath =  ''

        tmp[0] = self.get_address()

        tmp[1] = self.get_active_sheet()

        tmp[2] = xlwings.books.active.fullname

        if keys_or_values:
            self.values_range = tmp[0]
            self.values_sheet = tmp[1]
            self.values_filepath = tmp[2]
        else:
            self.keys_range = tmp[0]
            self.keys_sheet = tmp[1]
            self.keys_filepath = tmp[2]

        return tmp

    def set_sequence(self, sequence: str='') -> list[str]:
        tmp = sequence.split(' ')
        tmp_2 = []

        for i in range(len(tmp)):
            if tmp[i] != '':
                tmp_2.append(tmp[i])

        self.sequence = tmp_2

        return self.sequence

    def get_address(self):
        '''Get active cells range'''
        tmp = xlwings.load_address()
        
        if len(tmp.split(':')) == 1:
            tmp += ':' + tmp

        return tmp

    def get_active_sheet(self) -> str:
        tmp = str(xlwings.books.active.sheets.active).split(']')
        tmp = tmp[-1].split('>')[0]

        return tmp

    def get_data(self, file_path, sheet_name, range):
        book = xlwings.Book(file_path)
        sheet = book.sheets(sheet_name)
        return sheet[range].value
    
    def execute(self) -> None:
        result = {}
        any_chr = '(.+|)'

        for i in range(len(self.value_list)):
            for j in range(len(self.value_list[i])):
                tmp_regex = any_chr
                
                for n in self.sequence:
                    n = int(n)

                    if n <= len(self.row_list):
                        tmp_regex += self.row_list[n-1][j]
                    else:
                        n -= len(self.row_list)
                        tmp_regex += self.column_list[n-1][i]
                    tmp_regex += any_chr
                
                filtered_value = list(filter(lambda v: re.match(tmp_regex, v), self.key_list))
                
                for m in filtered_value:
                    if m != []:
                        result[m] = self.value_list[i][j]

        new_book = xlwings.Book()
        new_sheet = new_book.sheets[0]
        new_sheet.range('A1').expand().value = result