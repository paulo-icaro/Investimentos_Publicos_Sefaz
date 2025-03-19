# ---------------------------------------------------- #
# --- Data Processing - Investments on Ceara State --- #
# ---------------------------------------------------- #

# -*- coding: utf-8 -*-


# ----------------- #
# --- Libraries --- #
# ----------------- #
import pandas as pd
import os
from openpyxl import load_workbook


# --------------------- #
# --- Dataset Files --- #
# --------------------- #
folder_files = os.listdir('Dataset/')
dataset_full = pd.DataFrame()


# ----------------------- #
# --- Data_Processing --- #
# ----------------------- #
for x in range(len(folder_files)):
    period = folder_files[x][0:8]
    kind = folder_files[x][9:14]
    
    dataset = pd.read_excel(io = 'Dataset/' + folder_files[x],
                          header = 10,
                          usecols= 'C, F, I:K, N, Q, R',
                          names = ['codigo', 'descricao', 'lei', 'lei+cred', 'empenhado', 'pago', '%emp', '%pago'],
                          dtype = {'codigo':str})
    dataset = dataset.dropna()                                                                              # Remove na's
    dataset = dataset.assign(projeto = '', periodo = folder_files[x][0:8], tipo = folder_files[x][9:14])    # Add empty column


    # Identifying lines where column codigo has 3 characters 
    for i in range(len(dataset)):
        n_char = len(str(dataset['codigo'].iloc[i]))
        
        if n_char == 3:
            last_projeto = dataset['codigo'].iloc[i]
            dataset['projeto'].iloc[i] = ''
        else:
            dataset['projeto'].iloc[i] = last_projeto
        
    # Removing cases where column codigo has 3 characters
    remove_rows = dataset['projeto'] == ''
    dataset = dataset[~remove_rows]

    # Reordering and renaming
    dataset = dataset.reindex(columns = ['periodo', 'tipo', 'projeto', 'codigo', 'lei', 'lei+cred', 'empenhado', 'pago', '%emp', '%pago'])
    dataset.rename(columns = {'projeto':'programa', 'codigo':'regiao'}, inplace = True)
    
    # Pile datasets
    if x == 0:    
        dataset_full = dataset
    else:
        dataset_full = pd.concat([dataset_full, dataset])  
    
# Adjustments for percentage formatting
dataset_full['%emp'] = dataset_full['%emp']/100
dataset_full['%pago'] = dataset_full['%pago']/100


# ----------------------- #
# --- Storing Results --- #
# ----------------------- #
# Obs: when using with statement there is no need to save the sheet after opening it for formating
with pd.ExcelWriter(path = 'investimentos_siof_ceara.xlsx', engine='xlsxwriter') as writer:
    dataset_full.to_excel(excel_writer = writer, sheet_name = 'investimentos', index = False)

    # Just Formatting the Excel Sheet
    workbook = writer.book
    worksheet = writer.sheets['investimentos']
    money_formatting = workbook.add_format({'num_format':'R$#,##0'})
    perc_formatting = workbook.add_format({'num_format':'0.0%'})
    worksheet.set_column('E:H', 15, money_formatting)
    worksheet.set_column('I:J', 15, perc_formatting)
    worksheet.set_column('A:D', 15)
    
    
for x in range(len(folder_files)):
    print(x)