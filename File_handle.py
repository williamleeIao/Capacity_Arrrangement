import openpyxl
import pandas as pd
import itertools
import re


# open the template file and show every column
# implement property get at here

class Excel_Operation:

    def __init__(self):  # method_initliazing; will need to change if more than Z column
        value = [i for i in range(0, 26, 1)]
        alphabet = [chr(i) for i in range(65, 91, 1)]
        self.dictionary_alphabet = dict(zip(value, alphabet))
        print(self.dictionary_alphabet)

    def file_load(self, file_path, sheet_name, start_row, end_row):  # File initializing
        self.book = openpyxl.load_workbook(file_path)
        self.work_sheet = self.book[sheet_name]
        self.dataframe1 = pd.read_excel(file_path)
        print(self.dataframe1)
        value_to_search = self.__construct_list__(start_row, end_row)
        return value_to_search

    # rule 0
    def __value_to_skip__(self, value_to_skip):  # value_from the list box
        # convert title to list
        self.__total_algo = {}
        for i in value_to_skip:
            self.__total_algo[i] = list(self.dataframe1[i])
        print(self.__total_algo)

    # maybe no need
    def get_first_col(self, row_read="1", pick_col=[]):
        # check amount column
        list_value = []
        list_col_pick_up = pick_col
        for col_pick_up in list_col_pick_up:
            cell_to_read = col_pick_up + row_read
            list_value.append(self.work_sheet[cell_to_read].value)
        return list_value

    def __construct_list__(self, start_col, end_col):
        first_row = self.work_sheet[start_col:end_col]
        value_to_search = []  # mark as private variable
        # Construct list source from excel to check same or not
        for c1 in first_row:
            if c1[0].value is not None:
                # if c1[0].value not in Value_to_skip:
                value_to_search.append(c1[0].value)
            else:
                break
        return value_to_search

    # get column and show in the listbox
    def get_column_name(self, text=""):
        # need to have different name for the list, same name cannot capture
        row = "1"
        for i in range(65, 91, 1):
            text = chr(i)
            cell_to_read = text + row
            if self.work_sheet[cell_to_read].value == text:
                return i

    def __check_word__(self, word):
        spec_symbols = '+-'
        match = [l in spec_symbols for l in word]
        print(match)
        symbol = [word[i] for i in range(0, len(match)) if match[i]]

        group = [k for k, g in itertools.groupby(match)]
        print(group)
        return group, symbol

    def __arrangement__(self, outcome_checkword, whole_lst):
        p_0 = outcome_checkword[0]
        p_1 = outcome_checkword[1]
        main = whole_lst[0]
        non_main = whole_lst[1:len(whole_lst)]
        true_ctr = 0
        false_ctr = 0
        lstA = ['+']  # always first one is +
        lst2D = []
        for individual in p_0:
            if not individual:
                lstA.append(non_main[false_ctr])
                lst2D.append(lstA)
                lstA = []  # Reset the list
                false_ctr += 1
            else:
                lstA.append(p_1[true_ctr])
                true_ctr += 1
        return lst2D

    # First page to handle # rule 1
    def __run__(self, algorithm, value_to_search):
        one_d_list = []
        two_d_list = []
        multiple_algo = algorithm.split(';')
        # multiple algorithm out now
        # get the row from the text
        for single_algo in multiple_algo:
            # single algo out, one algo have multiple operation
            one_d_list.append(single_algo)
            multiple_value = re.split("\\+|\\-", single_algo)
            for value in multiple_value:
                value = value.strip()
                if value in value_to_search:
                    one_d_list.append(value_to_search.index(value))
            # one_d_list = [value_to_search.index(value) for value in multiple_value if value in value_to_search]
            two_d_list.append(one_d_list)
            one_d_list = []
        return two_d_list

    # rule 2
    def __alphabet_replacement__(self, index_interested):
        filter_list = []
        two_d_filter_list = []
        for index in index_interested:
            for i in range(1, 4, 1):
                index[i] = self.dictionary_alphabet.get(index[i])

        print(index_interested)
        return index_interested

    # rule 3
    def __algorithm_computation__(self, index_interested):
        total_value = []
        end_line = []
        # total_algo = {}
        start_cal_row = 2
        total = 0
        # depend how many algorithm need to compute by bracket
        for values in index_interested:
            operation_list = self.__check_word__(values[0])
            lst_arrangment = self.__arrangement__(operation_list, values)
            while True:
                for individual_element in lst_arrangment:
                    # put + /-
                    # individual_element is an 1D array
                    if self.work_sheet[individual_element[1] + str(start_cal_row)].value is not None:
                        # operation happen at here
                        if individual_element[0] == '+':
                            total = total + self.work_sheet[individual_element[1] + str(start_cal_row)].value
                        elif individual_element[0] == '-':
                            total = total - self.work_sheet[individual_element[1] + str(start_cal_row)].value
                        else:
                            pass
                        end_line.append(False)
                    else:
                        # the sheet is empty
                        # if value all is empty , assume end of the line
                        end_line.append(True)
                        total = 0
                if all(end_line): break
                total_value.append(total)
                start_cal_row += 1
                end_line = []  # reset
                total = 0
            start_cal_row = 2
            self.__total_algo[values[0]] = total_value
            total_value = []
        print(self.__total_algo)

    def __convert_to_df__(self):
        # create columns
        columns = list(self.__total_algo.keys())
        # create a new dictionary
        df = pd.DataFrame(data=self.__total_algo, columns=columns)
        print(df)
        return df

    @property
    def get___value_to_search(self):
        return self.__value_to_search

    @property
    def get_property_algo(self):
        return self.__total_algo

    def run_all_rule(self, algorithm, value_to_skip):
        value_to_search =[]
        first_row = self.work_sheet['A':"Z"]  # get a excel address
        self.__value_to_skip__(value_to_skip)  # remove any column that need to skip. This is not involve any algorithm calculation
        # fetch first row value out , this is need to use for searching
        for c1 in first_row:
            if c1[0].value is not None:
                # if c1[0].value not in Value_to_skip:
                value_to_search.append(c1[0].value)
            else:
                break
        index_interested = self.__run__(algorithm, value_to_search)
        index_interested = self.__alphabet_replacement__(index_interested)
        self.__algorithm_computation__(index_interested)
        df = self.__convert_to_df__()
        return df

    def save_new_file(self, dataframe_to_save, path_to_save):
        pass


Value_to_skip = ['Part Number', 'Description']  # <---This is variable
algorithm = 'October 2020+ November 2020 - December 2020;January 2021 + February 2021 + March 2021 ;April 2021 + May 2021 + June 2021'
file_path = r'C:\Users\willlee\Desktop\CONT Forecast Sept-20.xlsx'
excel = Excel_Operation()
value_to_search = excel.file_load(file_path, 'CONT Forecast', 'A', 'Z')
excel.run_all_rule(algorithm, Value_to_skip)

