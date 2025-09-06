# ========================================================================= #
# === Data Processing - Investments on Ceara State - Program and Region === #
# ========================================================================= #

# --- Script by Paulo Icaro --- #

# ================= #
# === Libraries === #
# ================= #
import pandas as pd
import os


# ===================== #
# === Dataset Files === #
# ===================== #
folder_files = os.listdir('Dataset/Investimentos_Funcao/')
dataset_full = pd.DataFrame()


# ======================= #
# === Data Processing === #
# ======================= #
for x in range(len(folder_files)):
    dataset = pd.read_excel(io = 'Dataset/Investimentos_Funcao/' + folder_files[x],
                            header = 10,
                            usecols= 'C, F, I:K, N, Q, R',
                            dtype = {'Código':str})
    dataset = dataset.dropna()                                                                              # Remove na's
    dataset['% Emp.'] = dataset['% Emp.'].astype(str).str.replace(',', '.', regex = False)                      # Setting standard decimal separator
    dataset['% Pago'] = dataset['% Pago'].astype(str).str.replace(',', '.', regex = False)                    # Setting standard decimal separator
    
    # ---------------- #
    # --- Function --- #
    # ---------------- #
    dataset = dataset.assign(periodo = folder_files[x][6:10] + '/' + folder_files[x][2:5],
                             tipo = folder_files[x][11:16],
                             ano = folder_files[x][6:10],
                             mes = folder_files[x][2:5])    
    
    
    for i in range(len(dataset)):
        n_char = len(str(dataset['Código'].iloc[i]))
    
        if n_char != 2:
            dataset['Descrição'].iloc[i] = ''
            
    # Removing cases where column descricao has more than 2 characters
    remove_rows = dataset['Descrição'] == ''
    dataset = dataset[~remove_rows]

    # Reordering and renaming
    dataset = dataset.reindex(columns = ['periodo', 'ano', 'mes', 'tipo', 'Código', 'Descrição', 'Lei', 'Lei+Cred.', 'Empenhado', 'Pago', '% Emp.', '% Pago'])
    dataset.rename(columns = {'Descrição':'Função'}, inplace = True)
    
    # Pile datasets
    if x == 0:    
        dataset_full = dataset
    else:
        dataset_full = pd.concat([dataset_full, dataset])



# Adjustments for percentage formatting
dataset_full['% Emp.'] = pd.to_numeric(dataset_full['% Emp.'])/100
dataset_full['% Pago'] = pd.to_numeric(dataset_full['% Pago'])/100



# ======================= #
# === Storing Results === #
# ======================= #
with pd.ExcelWriter(path = 'investimentos_siof_ceara_funcao.xlsx', engine = 'xlsxwriter') as writer:
    dataset_full.to_excel(excel_writer = writer, sheet_name = 'investimentos_funcao', index = False)
    
    # Quick Formatting
    workbook = writer.book
    worksheet = writer.sheets['investimentos_funcao']
    money_formatting = workbook.add_format({'num_format':'R$#,##0'})
    perc_formatting = workbook.add_format({'num_format':'0.0%'})
    worksheet.set_column('G:J', 15, money_formatting)    
    worksheet.set_column('K:L', 15, perc_formatting)


# Full Cleasing
del(dataset, folder_files, i, money_formatting, n_char, perc_formatting, remove_rows, workbook, worksheet, writer, x)