# ============================================== #
# === Data Processing - Debts on Ceara State === #
# ============================================== #

# --- Script by Paulo Icaro --- #

# ================= #
# === Libraries === #
# ================= #
import pandas as pd
import os


# ===================== #
# === Dataset Files === #
# ===================== #
folder_files = os.listdir('Dataset/Despesas_Orcamentarias/')
dataset_full = pd.DataFrame()


# ======================= #
# === Data Processing === #
# ======================= #
for x in range(len(folder_files)):
    dataset = pd.read_csv(filepath_or_buffer = 'Dataset/Despesas_Orcamentarias/' + folder_files[x],
                          sep = ';',  
                          header = 5,
                          encoding = 'latin1')
    dataset = dataset.dropna()
    dataset = dataset[dataset.Conta == 'INVESTIMENTOS']
    dataset = dataset[dataset.Coluna == 'DESPESAS PAGAS ATÉ O BIMESTRE (j)']
    dataset = dataset.drop(columns = 'População', axis = 1)
    dataset = dataset.assign(periodo = folder_files[x][0:4] + '/' + folder_files[x][5:7],
                             ano = folder_files[x][0:4],
                             bimestre = folder_files[x][5:7],
                             municipio = dataset['Instituição'].str[24:].str[:-5])
    
    
    if x != 0:
        dataset_full = pd.concat([dataset_full, dataset])
    else:
        dataset_full = dataset


# Adjustments for money formatting
dataset_full['Valor'] = dataset_full['Valor'].str.replace(',' , '.')
dataset_full['Valor'] = pd.to_numeric(dataset_full['Valor'])


# Dataset matrix
matrix_dataset = dataset_full.pivot(index = 'municipio', columns = 'periodo', values = 'Valor')


# ======================= #
# === Storing Results === #
# ======================= #
with pd.ExcelWriter(path = 'despesas_siconfi_ceara.xlsx', engine = 'xlsxwriter') as writer:
    dataset_full.to_excel(excel_writer = writer, sheet_name = 'despesas', index = False)
    
    # Quick Formatting
    workbook = writer.book
    worksheet = writer.sheets['despesas']
    money_formatting = workbook.add_format({'num_format':'R$#,##0'})
    worksheet.set_column('G:G', 15, money_formatting)


# Full Cleasing
del(dataset, x, folder_files, workbook, worksheet, writer, money_formatting)