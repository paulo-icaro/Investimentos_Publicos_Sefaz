# ========================================================================= #
# === DATA PROCESSING - INVESTMENTS ON CEARA STATE - PROGRAM AND REGION === #
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

while info_desired not in {'p', 'r', 'pr', 'rp'}:
    info_desired = input('Como você deseja a base de dados ? Use (P) para Programa, (R) para Região e (PR) para Programa e Região: ').strip().lower()
    
    if info_desired not in {'p', 'r', 'pr', 'rp'}: 
        print('Opção inválida !')
    elif info_desired == 'p':
        print('Tratando as informações por Programa ...')
        break
    elif info_desired == 'r':
        print('Tratando as informações por Região ...')
        break
    else:
        print('Tratando as informações por Programa e Região ...')
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
                          usecols= 'C, F, K, N',
                          names = ['codigo', 'descricao', 'empenhado', 'pago'],
                          dtype = {'codigo':str})
    dataset = dataset.dropna()                                                                             # Remove na's
    
    
    # --------------- #
    # --- Program --- #
    # --------------- #
    if info_desired == 'p':
        dataset = dataset.assign(periodo = folder_files[x][0:8],
                                 tipo = folder_files[x][9:14],
                                 ano = folder_files[x][4:8],
                                 mes = folder_files[x][0:3])    
        
        # --- Replacements --- #
        replacements = {'JAN':'01', 'FEV':'02', 'MAR':'03', 'ABR':'04', 'MAI':'05', 'JUN':'06', 'JUL':'07', 'AGO':'08', 'SET':'09', 'OUT':'10', 'NOV':'11', 'DEZ':'12'}
        for old, new in replacements.items():
            dataset['mes'] = dataset['mes'].replace(old,new)
        
        # --- Identifying Region rows --- #
        program_flag = dataset['codigo'].str.len() == 2            
        
        # --- Removing cases where column descricao has 2 characters --- #
        dataset = dataset[~program_flag]
        
        # --- Reordering and renaming --- #
        dataset = dataset.reindex(columns = ['periodo', 'ano', 'mes', 'tipo', 'codigo', 'descricao', 'empenhado', 'pago'])
        dataset.rename(columns = {'descricao':'programa', 'codigo':'cod_programa'}, inplace = True)
        
        
        
    # -------------- #
    # --- Region --- #
    # -------------- #
    if info_desired == 'r':
        dataset = dataset.assign(periodo = folder_files[x][0:8],
                                 tipo = folder_files[x][9:14],
                                 ano = folder_files[x][4:8],
                                 mes = folder_files[x][0:3])    
        
        # --- Replacements --- #
        replacements = {'JAN':'01', 'FEV':'02', 'MAR':'03', 'ABR':'04', 'MAI':'05', 'JUN':'06', 'JUL':'07', 'AGO':'08', 'SET':'09', 'OUT':'10', 'NOV':'11', 'DEZ':'12'}
        for old, new in replacements.items():
            dataset['mes'] = dataset['mes'].replace(old,new)        
        
        # --- Identifying Program rows --- #
        region_flag = dataset['codigo'].str.len() == 3        
        
        # --- Removing cases where column descricao has 3 characters --- #
        dataset = dataset[~region_flag]
        
        # --- Reordering and renaming --- #
        dataset = dataset.reindex(columns = ['periodo', 'ano', 'mes', 'tipo', 'codigo', 'descricao', 'empenhado', 'pago'])
        dataset.rename(columns = {'descricao':'regiao', 'codigo':'cod_regiao'}, inplace = True)

        
        
    # -------------------------- #
    # --- Program and Region --- #
    # -------------------------- #        
    if info_desired == 'rp' or info_desired == 'pr':        
        dataset = dataset.assign(cod_programa = None, 
                                 programa = None,                                                                 # Add empty column programa
                                 periodo = folder_files[x][0:8],
                                 tipo = folder_files[x][9:14],
                                 ano = folder_files[x][4:8],
                                 mes = folder_files[x][0:3])
        
        # --- Replacements --- #
        replacements = {'JAN':'01', 'FEV':'02', 'MAR':'03', 'ABR':'04', 'MAI':'05', 'JUN':'06', 'JUL':'07', 'AGO':'08', 'SET':'09', 'OUT':'10', 'NOV':'11', 'DEZ':'12'}
        for old, new in replacements.items():
            dataset['mes'] = dataset['mes'].replace(old,new)            
    
        # --- Identifying Program rows --- #
        program_flag = dataset['codigo'].str.len() == 3        
        
        # --- Filling Columns --- #
        dataset.loc[program_flag, 'cod_programa'] = dataset['codigo']
        dataset.loc[program_flag, 'programa'] = dataset['descricao']
        dataset[['cod_programa','programa']] = dataset[['cod_programa','programa']].ffill()        
            
        # --- Removing cases where column codigo has 3 characters --- #
        dataset = dataset[~program_flag]
    
        # --- Reordering and renaming --- #
        dataset = dataset.reindex(columns = ['periodo', 'ano', 'mes', 'tipo', 'codigo', 'descricao', 'cod_programa', 'programa', 'empenhado', 'pago'])
        dataset.rename(columns = {'descricao':'regiao', 'codigo':'cod_regiao'}, inplace = True)
        
        
    

    # ------------------------- #    
    # --- Stacking datasets --- #r
    # ------------------------- #
    if x == 0:    
        dataset_full = dataset
    else:
        dataset_full = pd.concat([dataset_full, dataset])   
 
    
 
# ======================================= #
# === Adjustments for cumulative data === #
# ======================================= #

# --- Sorting --- #
if info_desired == 'p':
    dataset_full = dataset_full.sort_values(by = ['cod_programa', 'tipo', 'ano', 'mes']).reset_index(drop = True)
elif info_desired == 'r':
    dataset_full = dataset_full.sort_values(by = ['cod_regiao', 'tipo', 'ano', 'mes']).reset_index(drop = True)
else:
    dataset_full = dataset_full.sort_values(by = ['tipo', 'regiao', 'cod_programa','ano', 'mes']).reset_index(drop = True)


# --- Cumulative data adjustment --- #
dataset_full['empenhado_mensal'] = dataset_full['empenhado'] - dataset_full['empenhado'].shift(1)     # Inserting adjusted values
dataset_full['pago_mensal'] = dataset_full['pago'] - dataset_full['pago'].shift(1)                    # Inserting adjusted values


# --- Loop for adjusting date truncated values --- #
for i in range(len(dataset_full)):  
    if i == 0:
        dataset_full.loc[0, 'empenhado_mensal'] = dataset_full.loc[0, 'empenhado']
        dataset_full.loc[0, 'pago_mensal'] = dataset_full.loc[0, 'pago']
    if i != 0 and int(dataset_full.loc[i, 'mes']) - int(dataset_full.loc[i-1, 'mes']) != 1:
        dataset_full.loc[i, 'empenhado_mensal'] = dataset_full.loc[i, 'empenhado']
        dataset_full.loc[i, 'pago_mensal'] = dataset_full.loc[i, 'pago']
    
    
# --- Renaming columns --- #
dataset_full.rename(columns = {'empenhado':'empenhado_acumulado', 'pago':'pago_acumulado'}, inplace = True)



# ======================= #
# === Storing Results === #
# ======================= #
# Obs: when using with statement there is no need to save the sheet after opening it for formating
if info_desired == 'p':       
    
    # --- Vertical dataset adjustment --- #
    dataset_full = dataset_full.melt(
        id_vars = ['periodo', 'ano', 'mes', 'tipo', 'cod_programa', 'programa'],
        value_vars = ['empenhado_acumulado', 'pago_acumulado', 'empenhado_mensal', 'pago_mensal'],
        var_name = 'categoria',
        value_name = 'valor'
        )
    
    # --- Storing --- #
    with pd.ExcelWriter(path = 'investimentos_siof_ceara_programa.xlsx', engine='xlsxwriter') as writer:
        dataset_full.to_excel(excel_writer = writer, sheet_name = 'investimentos_programa', index = False)

        # Just Formatting the Excel Sheet (not needed in case of vertical adjustment)
        '''workbook = writer.book
        worksheet = writer.sheets['investimentos_programa']
        money_formatting = workbook.add_format({'num_format':'R$#,##0'})
        perc_formatting = workbook.add_format({'num_format':'0.0%'})
        worksheet.set_column('G:J', 15, money_formatting)
        worksheet.set_column('K:L', 15, perc_formatting)
        worksheet.set_column('A:F', 15)'''
    
    # Full Cleasing
    del(dataset, folder_files, i, info_desired, writer, x, new, old, replacements, program_flag)#, money_formatting, perc_formatting, workbook, worksheet)
    
elif info_desired == 'r':       
    
    # --- Vertical dataset adjustment --- #
    dataset_full = dataset_full.melt(
        id_vars = ['periodo', 'ano', 'mes', 'tipo', 'cod_regiao', 'regiao'],
        value_vars = ['empenhado_acumulado', 'pago_acumulado', 'empenhado_mensal', 'pago_mensal'],
        var_name = 'categoria',
        value_name = 'valor'
        )
    
    # --- Storing --- #
    with pd.ExcelWriter(path = 'investimentos_siof_ceara_regiao.xlsx', engine='xlsxwriter') as writer:
        dataset_full.to_excel(excel_writer = writer, sheet_name = 'investimentos_regiao', index = False)

        # Just Formatting the Excel Sheet (not needed in case of vertical adjustment)
        '''workbook = writer.book
        worksheet = writer.sheets['investimentos_regiao']
        money_formatting = workbook.add_format({'num_format':'R$#,##0'})
        perc_formatting = workbook.add_format({'num_format':'0.0%'})
        worksheet.set_column('G:J', 15, money_formatting)
        worksheet.set_column('K:L', 15, perc_formatting)
        worksheet.set_column('A:F', 15)'''
    
    # --- Full Cleasing --- #
    del(dataset, folder_files, i, info_desired, writer, x, new, old, replacements, region_flag)#, money_formatting, perc_formatting, workbook, worksheet)


else:
    
    # --- Vertical dataset adjustment --- #
    dataset_full = dataset_full.melt(
        id_vars = ['periodo', 'ano', 'mes', 'tipo', 'cod_regiao', 'regiao', 'cod_programa', 'programa'],
        value_vars = ['empenhado_acumulado', 'pago_acumulado', 'empenhado_mensal', 'pago_mensal'],
        var_name = 'categoria',
        value_name = 'valor'
        )
    
    
    
    with pd.ExcelWriter(path = 'investimentos_siof_ceara_programa_regiao.xlsx', engine='xlsxwriter') as writer:
        dataset_full.to_excel(excel_writer = writer, sheet_name = 'investimentos_programa_regiao', index = False)

        # Just Formatting the Excel Sheet (not needed in case of vertical adjustment)
        '''workbook = writer.book
        worksheet = writer.sheets['investimentos_programa_regiao']
        money_formatting = workbook.add_format({'num_format':'R$#,##0'})
        perc_formatting = workbook.add_format({'num_format':'0.0%'})
        worksheet.set_column('I:L', 15, money_formatting)
        worksheet.set_column('M:N', 15, perc_formatting)
        worksheet.set_column('A:F', 15)'''
    
    # --- Full Cleasing --- #
    del(dataset, folder_files, i, info_desired, writer, x, new, old, replacements, program_flag)#, money_formatting, perc_formatting, workbook, worksheet)