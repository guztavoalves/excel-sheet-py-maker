"""
Filename: excel_sheet_py_maker.py
Description: Make Excel Sheets Easy simplifies creating Excel sheets. 
Customize columns, set data types, and input entries effortlessly to generate organized, ready-to-use spreadsheets 
quickly and efficiently.
Author: Gustavo Alves
Date: 2024-26-10
Site: https://github.com/guztavoalves
Version: 1.0
"""
from datetime import datetime
import os, time
import pandas as pd

# Initial configurations
#
input_types = ['text','number','date','currency']
columns_config = {}

# SHOW INFO
#
def show_columns_info(cols):
    clean()
    print(f'\nYour spreadsheet has {len(cols)} columns:\n')

    has_columns_config = len(columns_config) > 0

    for i, col in enumerate(cols,1):
        colum_type = columns_config[col] if has_columns_config else '-'
        print(f'Column #{i}: [{col}] | Type: [{colum_type}]')

def show_input_types(colum):
    clean()
    print(f'\nWhat type of data will column [{colum}] receive?\n')
    for i, itp in enumerate(input_types,1):
        print(i, itp, sep=' - ')


# TOOL
#
def clean():
    os.system('cls')
    msg_title()


# INPUTS
#

def input_get_col_name():
    clean()
    return input('\nEnter the name of each column separated by commas: \n')

def input_get_col_value(col, col_type='text'):
    clean()
    return input(f'\nEnter data for the field of type {col_type} [{com}]: \n')

def input_press_to_continue():
    input('Press ENTER to try again...')


# MESSAGES / ASK
#
def msg_title():
    print("""Ｅｘｃｅｌ  Ｓｈｅｅｔ  Ｐｙ  Ｍａｋｅｒ""")

def msg_exit():
    clean()
    print('\nSee you later! =)')

def msg_sheet_create():
    clean()
    print('\nCreating table...\n')

def msg_sucess():
    clean()
    print('\nExcel file saved successfully!\n')

def msg_invalid_data():
    clean()
    print('\nEnter at least one valid value!\n')
    input_press_to_continue()

def msg_column_empty():
    clean()
    print('\nEnter at least one column name!\n')
    input_press_to_continue()

def msg_invalid_type_select():
    print('\nSelect one of the available numbers!\n')

def ask_continue_make_sheets():
    clean()
    if input('\nDo you want to create another spreadsheet? (Y/n) \n').lower() in ['y','yes']:
        return True
    
    return False

def ask_reset_columns_register():
    if input('\nDo you want to redo the spreadsheet columns? (Y/n) \n').lower() in ['y','yes']:
        return True
    
    return False

def ask_reset_columns_config():
    if input('\nDo you want to redo the spreadsheet column settings? (Y/n) \n').lower() in ['y','yes']:
        return True
    
    return False

def ask_configure_columns():
    clean()
    if input('\nDo you want to configure the columns in your spreadsheet? (Y/n) \n') in ['y','yes']:
        return True
    
    return False

def ask_continue_insert_colum():
    clean()
    if input('\nDo you want to insert another column? (Y/n) \n') in ['y','yes']:
        return True
    
    return False

def ask_continue_insert_entries():
    clean()
    if input('\nDo you want to insert another record? (Y/n) \n') in ['y','yes']:
        return True
    
    return False


# VALIDATORS
#
def validator(data, data_type):

    match data_type:

        case 'text':
            return valid_text(data)

        case 'number':
            return valid_number(data)

        case 'date':
            return valid_date(data)

        case 'currency':
            return valid_currency(data)

    return False

def valid_number(data):

    if data.isnumeric():
        return data
    
    return False

def valid_text(data):

    if valid_number(data) or type(data) != str:
        return False
    
    return data

def valid_date(data):

    try:
        datetime.strptime(data, '%d/%m/%Y')
    except:
        return False

    return data

def valid_currency(data):

    try:
        data = '{:.2f}'.format(float(data))
    except:
        return False

    return data


# MAKER
#
def get_save_dir():
    return os.path.dirname(__file__)

def make_data_frame(data):
    return pd.DataFrame(data)

def save_to_excel(data_obj, filename):
    save_path = os.path.join(get_save_dir(),'sheets',str(time.time())+'_'+filename+'.xlsx')
    data_obj.to_excel(save_path, index=False)

def get_data():
    return make_data_entries()

def make_data_entries():
    data = []
    cols = []

    while True:

        col_entries = make_column_entries()
        
        if col_entries == None and len(cols) <= 0:
            msg_column_empty()
            continue

        if col_entries:

            for col_entry in col_entries:
                cols.append(col_entry.strip())

        if not ask_continue_insert_colum():
            show_columns_info(cols)

            if ask_reset_columns_register():
                cols.clear()
                continue

            break    

    
    while True:

        if ask_configure_columns():
            columns_config.update(configure_columns(cols))
        else:
            columns_config.update(configure_columns(cols,'text'))

        show_columns_info(cols)

        if not ask_reset_columns_config():
            break

        columns_config.clear()


    while True:

        entry = {}
        for col in cols:
            sheet_entry = make_sheet_entry(col)
            entry.update(sheet_entry)

        data.append(entry)

        if not ask_continue_insert_entries():
            break
    
    return data

def configure_columns(cols, col_type=None):

    columns = {}
    for col in cols:

        if col_type == None:
            show_input_types(col)
            columns.update({col : input_types[select_input_type()-1]})
        else:
            columns.update({col : col_type})
        
    return columns

def make_column_entries():

    while True:
        columns = input_get_col_name()
        if columns:
            break
        else:
            return None

    columns_list = columns.split(',')
    return columns_list
        
def make_sheet_entry(col):

    while True:
        value = validator(input_get_col_value(col, columns_config[col]), columns_config[col])

        if value:
            return {col : value}
        else:
            msg_invalid_data()

def select_input_type():

    while True:
        data_input = input('Enter the number: ')
        data_valid = int(data_input) if valid_number(data_input) else None

        if data_valid and data_valid <= len(input_types):
            break

        msg_invalid_type_select()

    return data_valid


# INIT
#
def main():
        
    while True:
        clean()
        filename = input('\nEnter the name of the excel file: \n')

        data = get_data()
        msg_sheet_create()

        save_to_excel(make_data_frame(data), filename)
        msg_sucess()

        if not ask_continue_make_sheets():
            msg_exit()
            break