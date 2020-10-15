from openpyxl import load_workbook
import re
import pandas as pd
import itertools

algorithm = 'October 2020+ November 2020 - December 2020;January 2021 + February 2021 + March 2021 ;April 2021 + May 2021 + June 2021'
total_algo = {}
# this is get from list box UI
Value_to_skip = ['Part Number', 'Description']  # <---This is variable


def arrangement(outcome_checkword, whole_lst):
    p_0 = outcome_checkword[0]
    p_1 = outcome_checkword[1]
    main = whole_lst[0]
    non_main = whole_lst[1:len(whole_lst)]
    true_ctr = 0
    false_ctr = 0
    lstA = ['+'] # always first one is +
    lst2D =[]
    for individual in p_0:
        if not individual:
            lstA.append(non_main[false_ctr])
            lst2D.append(lstA)
            lstA =[] # Reset the list
            false_ctr += 1
        else:
            lstA.append(p_1[true_ctr])
            true_ctr += 1
    return lst2D


def check_word(word):
    spec_symbols = '+-'
    match = [l in spec_symbols for l in word]
    print(match)
    symbol = [word[i] for i in range(0, len(match)) if match[i]]

    group = [k for k, g in itertools.groupby(match)]
    print(group)
    return group, symbol


def run(algorithm, value_to_search):
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


book = load_workbook(r'C:\Users\willlee\Desktop\CONT Forecast Sept-20.xlsx')

sheet = book.active

# first row is the name
# read the first_row
# first row only
# mapping
value = [i for i in range(0, 26, 1)]
alphabet = [chr(i) for i in range(65, 91, 1)]

dataframe1 = pd.read_excel(r'C:\Users\willlee\Desktop\CONT Forecast Sept-20.xlsx')
print(dataframe1)
# convert title to list
for i in Value_to_skip:
    total_algo[i] = list(dataframe1[i])

print(total_algo)

# Connstruct list
dictionary_alphaber = dict(zip(value, alphabet))
print(dictionary_alphaber)
start = 'A'
end = 'Z'

value_to_search = []
first_row = sheet[start:end]
first_row_values = ['October 2020', 'December 2020', 'March 2021', 'May 2021']
index_interested = []
for c1 in first_row:
    print(c1[0])
    print(c1[0].value)

# Construct list source from excel to check same or not
for c1 in first_row:
    if c1[0].value is not None:
        # if c1[0].value not in Value_to_skip:
        value_to_search.append(c1[0].value)
    else:
        break

print(value_to_search)
# try this out
index_interested = run(algorithm, value_to_search)
# index_interested = [value_to_search.index(value) for value in first_row_values if value in value_to_search]

# print(index_interested)
# filter_list = [dictionary_alphaber.get(index) for index in index_interested]
# Check for the list
# two D list
# replace number with alphabet
filter_list = []
two_d_filter_list = []
for index in index_interested:
    for i in range(1, 4, 1):
        index[i] = dictionary_alphaber.get(index[i])
# replace
print(index_interested)
# filter_list = [dictionary_alphaber.get(index) for index in index_interested for i in range(1,4,1)]
print()
# generate list to calculate
total_value = []
end_line = []
# total_algo = {}
start_cal_row = 2
total = 0
# depend how many algorithm need to compute by bracket
for values in index_interested:
    operation_list = check_word(values[0])
    lst_arrangment = arrangement(operation_list, values)
    while True:
        for individual_element  in lst_arrangment:
            # put + /-
            # individual_element is an 1D array
            if sheet[individual_element[1] + str(start_cal_row)].value is not None:
                # operation happen at here
                if individual_element[0] == '+':
                    total = total + sheet[individual_element[1] + str(start_cal_row)].value
                elif individual_element[0] == '-':
                    total = total - sheet[individual_element[1] + str(start_cal_row)].value
                else: pass
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
    total_algo[values[0]] = total_value
    total_value = []
print(total_algo)  # <---dictionary to put in data frame

# Create a new excel_list or data frame
# Create Columns

# Value to skip get from original file as data frame
# create columns
columns = list(total_algo.keys())
# create a new dictionary
df = pd.DataFrame(data=total_algo, columns=columns)
print(df)
# create data
# data =
# = pd.DataFrame(Value_to_skip[0], Value_to_skip[1],
#                Value_to_skip[0]: [1, 2, 3],
# Value_to_skip[1]: []
# )
print()
