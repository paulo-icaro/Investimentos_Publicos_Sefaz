# ========================================================================= #
# === Data Processing - Investments on Ceara State - Program and Region === #
# ========================================================================= #

# --- Script by Paulo Icaro --- #

# ================= #
# === Libraries === #
# ================= #
import pandas as pd
import os
#from openpyxl import load_workbook


# ================================== #
# === Defining the data you want === #
# ================================== #
info_desired = ''

while info_desired not in {'p', 'pr', 'rp'}:
    info_desired = input('Como você deseja a base de dados ? Use (P) para Programa e (PR) para Programa e Região: ').strip().lower()
    
    if info_desired not in {'p', 'pr', 'rp'}: 
        print('Opção inválida !')
    elif info_desired == 'p':
        print('Tratando as informações apenas por Programa ...')
        break
    else:
        print('Tratando as informações por Programa e Região...')
        break


# ===================== #
# === Dataset Files === #
# ===================== #
folder_files = os.listdir('Dataset/Investimentos_Programa_Regiao/')
dataset_full = pd.DataFrame()


# ======================= #
# === Data Processing === #
# ======================= #
for x in range(len(folder_files)):  
    
    dataset = pd.read_excel(io = 'Dataset/Investimentos_Programa_Regiao/' + folder_files[x],
                          header = 10,
                          usecols= 'C, F, I:K, N, Q, R',
                          names = ['codigo', 'descricao', 'lei', 'lei+cred', 'empenhado', 'pago', '%emp', '%pago'],
                          dtype = {'codigo':str})#, '%emp':float, '%pago':float})
    dataset = dataset.dropna()                                                                              # Remove na's
    dataset['%emp'] = dataset['%emp'].astype(str).str.replace(',', '.', regex = False)                      # Setting standard decimal separator
    dataset['%pago'] = dataset['%pago'].astype(str).str.replace(',', '.', regex = False)                    # Setting standard decimal separator
    
    # --------------- #
    # --- Program --- #
    # --------------- #
    if info_desired == 'p':
        dataset = dataset.assign(periodo = folder_files[x][0:8],
                                 tipo = folder_files[x][9:14],
                                 ano = folder_files[x][4:8],
                                 mes = folder_files[x][0:3])    
        
        # Identifying Program rows
        for i in range(len(dataset)):
            n_char = len(str(dataset['codigo'].iloc[i]))
        
            if n_char == 2:
                dataset['descricao'].iloc[i] = ''
        
        # Removing cases where column descricao has 2 characters
        remove_rows = dataset['descricao'] == ''
        dataset = dataset[~remove_rows]
        
        # Reordering and renaming
        dataset = dataset.reindex(columns = ['periodo', 'ano', 'mes', 'tipo', 'codigo', 'descricao', 'lei', 'lei+cred', 'empenhado', 'pago', '%emp', '%pago'])
        dataset.rename(columns = {'descricao':'program', 'codigo':'cod_program'}, inplace = True)
    
    
    
    # -------------------------- #
    # --- Program and Region --- #
    # -------------------------- #        
    else:        
        dataset = dataset.assign(cod_program = '', 
                                 program = '',                                                                 # Add empty column program
                                 periodo = folder_files[x][0:8],
                                 tipo = folder_files[x][9:14],
                                 ano = folder_files[x][4:8],
                                 mes = folder_files[x][0:3])    
    
        # Identifying Program rows
        for i in range(len(dataset)):
            n_char = len(str(dataset['codigo'].iloc[i]))
    
            if n_char == 3:
                dataset['program'].iloc[i] = ''
                cod_last_program = dataset['codigo'].iloc[i]
                last_program = dataset['descricao'].iloc[i]
            else:
                dataset['cod_program'].iloc[i] = cod_last_program
                dataset['program'].iloc[i] = last_program
            
        # Removing cases where column codigo has 3 characters
        remove_rows = dataset['program'] == ''
        dataset = dataset[~remove_rows]
    
        # Reordering and renaming
        dataset = dataset.reindex(columns = ['periodo', 'ano', 'mes', 'tipo', 'codigo', 'descricao', 'cod_program', 'program', 'lei', 'lei+cred', 'empenhado', 'pago', '%emp', '%pago'])
        dataset.rename(columns = {'descricao':'regiao', 'codigo':'cod_regiao'}, inplace = True)
        

    # Pile datasets
    if x == 0:    
        dataset_full = dataset
    else:
        dataset_full = pd.concat([dataset_full, dataset])



# Adjustments for percentage formatting
dataset_full['%emp'] = pd.to_numeric(dataset_full['%emp'])/100
dataset_full['%pago'] = pd.to_numeric(dataset_full['%pago'])/100





# ======================= #
# === Storing Results === #
# ======================= #
# Obs: when using with statement there is no need to save the sheet after opening it for formating
if info_desired == 'p':       
    
    # Vertical dataset adjustment
    dataset_full = dataset_full.melt(
        id_vars = ['periodo', 'ano', 'mes', 'tipo', 'cod_program', 'program'],
        value_vars = ['lei', 'lei+cred', 'empenhado', 'pago', '%emp', '%pago'],
        var_name = 'categoria',
        value_name = 'valor'
        )
    
    with pd.ExcelWriter(path = 'investimentos_siof_ceara_programa.xlsx', engine='xlsxwriter') as writer:
        dataset_full.to_excel(excel_writer = writer, sheet_name = 'investimentos_programa', index = False)

        # Just Formatting the Excel Sheet (not needed in case of vertical adjustment)
        #workbook = writer.book
        #worksheet = writer.sheets['investimentos_programa']
        #money_formatting = workbook.add_format({'num_format':'R$#,##0'})
        #perc_formatting = workbook.add_format({'num_format':'0.0%'})
        #worksheet.set_column('G:J', 15, money_formatting)
        #worksheet.set_column('K:L', 15, perc_formatting)
        #worksheet.set_column('A:F', 15)
    
    # Full Cleasing
    del(dataset, folder_files, i, info_desired, n_char, remove_rows, writer, x)#, money_formatting, perc_formatting, workbook, worksheet)

else:
    
    # Vertical dataset adjustment
    dataset_full = dataset_full.melt(
        id_vars = ['periodo', 'ano', 'mes', 'tipo', 'cod_regiao', 'regiao', 'cod_program', 'program'],
        value_vars = ['lei', 'lei+cred', 'empenhado', 'pago', '%emp', '%pago'],
        var_name = 'categoria',
        value_name = 'valor'
        )
    
    
    
    with pd.ExcelWriter(path = 'investimentos_siof_ceara_programa_regiao.xlsx', engine='xlsxwriter') as writer:
        dataset_full.to_excel(excel_writer = writer, sheet_name = 'investimentos_programa_regiao', index = False)

        # Just Formatting the Excel Sheet (not needed in case of vertical adjustment)
        #workbook = writer.book
        #worksheet = writer.sheets['investimentos_programa_regiao']
        #money_formatting = workbook.add_format({'num_format':'R$#,##0'})
        #perc_formatting = workbook.add_format({'num_format':'0.0%'})
        #worksheet.set_column('I:L', 15, money_formatting)
        #worksheet.set_column('M:N', 15, perc_formatting)
        #worksheet.set_column('A:F', 15)
    
    # Full Cleasing
    del(dataset, folder_files, i, info_desired, n_char, remove_rows, writer, x, last_program, cod_last_program)#, money_formatting, perc_formatting, workbook, worksheet)

        

        
'''for x in range(len(folder_files)):
    print(x)'''