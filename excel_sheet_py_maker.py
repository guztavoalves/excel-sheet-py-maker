"""
Filename: excel_sheet_py_maker.py
Description: Make Excel Sheets Easy simplifies creating Excel sheets. 
Customize columns, set data types, and input entries effortlessly to generate organized, ready-to-use spreadsheets 
quickly and efficiently.
Author: Gustavo Alves
Date: 2024-26-10
Site: https://github.com/guztavoalves
"""
from datetime import datetime
import os, time
import pandas as pd

# Initial configurations
#
input_types = ['texto','número','data','moeda']
columns_config = {}

# SHOW INFO
#
def show_columns_info(cols):
    clean()
    print(f'\nSua planilha possui {len(cols)} colunas:\n')

    has_columns_config = len(columns_config) > 0

    for i, col in enumerate(cols,1):
        colum_type = columns_config[col] if has_columns_config else '-'
        print(f'Coluna #{i}: [{col}] | Tipo: [{colum_type}]')

def show_input_types(colum):
    clean()
    print(f'\nQual o tipo de dados a coluna [{colum}] irá receber?\n')
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
    return input('\nDigite o nome de cada coluna separado por virgula: \n')

def input_get_col_value(col, col_type='texto'):
    clean()
    return input(f'\nDigite os dados para o campo do tipo {col_type} [{col}]: \n')

def input_press_to_continue():
    input('Pressione ENTER para tentar novamente...')


# MESSAGES / ASK
#
def msg_title():
    print("""Ｅｘｃｅｌ  Ｓｈｅｅｔ  Ｐｙ  Ｍａｋｅｒ""")

def msg_exit():
    clean()
    print('\nAté logo! =)')

def msg_sheet_create():
    clean()
    print('\nCriando tabela...\n')

def msg_sucess():
    clean()
    print('\nArquivo excel salvo com sucesso!\n')

def msg_invalid_data():
    clean()
    print('\nDigite ao menos um valor válido!\n')
    input_press_to_continue()

def msg_column_empty():
    clean()
    print('\nDigite ao menos o nome de uma coluna!\n')
    input_press_to_continue()

def msg_invalid_type_select():
    print('\nSelecione um dos números disponíveis!\n')

def ask_continue_make_sheets():
    clean()
    if input('\nDeseja criar outra planilha? (Y/n) \n') in ['Y','y','yes','s','S','Sim','sim']:
        return True
    
    return False

def ask_reset_columns_register():
    if input('\nDeseja refazer as colunas da planilha? (Y/n) \n') in ['Y','y','yes','s','S','Sim','sim']:
        return True
    
    return False

def ask_reset_columns_config():
    if input('\nDeseja refazer as configurações das colunas da planilha? (Y/n) \n') in ['Y','y','yes','s','S','Sim','sim']:
        return True
    
    return False

def ask_configure_columns():
    clean()
    if input('\nDeseja configurar as colunas da sua planilha? (Y/n) \n') in ['Y','y','yes','s','S','Sim','sim']:
        return True
    
    return False

def ask_continue_insert_colum():
    clean()
    if input('\nDeseja inserir outra coluna? (Y/n) \n') in ['Y','y','yes','s','S','Sim','sim']:
        return True
    
    return False

def ask_continue_insert_entries():
    clean()
    if input('\nDeseja inserir outro registro? (Y/n) \n') in ['Y','y','yes','s','S','Sim','sim']:
        return True
    
    return False


# VALIDATORS
#
def validator(data, data_type):

    match data_type:

        case 'texto':
            return valid_text(data)

        case 'número':
            return valid_number(data)

        case 'data':
            return valid_date(data)

        case 'moeda':
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
            columns_config.update(configure_columns(cols,'texto'))

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
        data_input = input('Digite o número: ')
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
        filename = input('\nDigite o nome do arquivo excel: \n')

        data = get_data()
        msg_sheet_create()

        save_to_excel(make_data_frame(data), filename)
        msg_sucess()

        if not ask_continue_make_sheets():
            msg_exit()
            break