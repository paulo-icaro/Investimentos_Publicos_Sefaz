# =============================================================== #
# === Data Processing - Investments on Ceara State - Function === #
# =============================================================== #

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
                            usecols= 'C, F, K, N',
                            dtype = {'Código':str})
    dataset = dataset.dropna()                                                                                # Remove na's
        
    
    # ---------------- #
    # --- Function --- #
    # ---------------- #
    
    # --- Adding columns --- #
    dataset = dataset.assign(periodo = folder_files[x][6:10] + '/' + folder_files[x][2:5],
                             tipo = folder_files[x][11:16],
                             ano = folder_files[x][6:10],
                             mes = folder_files[x][2:5])    
    
    # --- Making a few replacements --- #
    replacements = {'JAN':'01', 'FEV':'02', 'MAR':'03', 'ABR':'04', 'MAI':'05', 'JUN':'06', 'JUL':'07', 'AGO':'08', 'SET':'09', 'OUT':'10', 'NOV':'11', 'DEZ':'12'}
    for old, new in replacements.items():
        dataset['mes'] = dataset['mes'].replace(old,new)
    
    # --- Identifying non Function rows --- #
    function_flag = dataset['Código'].str.len() == 2
            
    # --- Removing cases where column descricao has more than 2 characters --- #
    dataset = dataset[function_flag]

    # --- Reordering and renaming --- #
    dataset = dataset.reindex(columns = ['periodo', 'ano', 'mes', 'tipo', 'Código', 'Descrição', 'Empenhado', 'Pago'])
    dataset.rename(columns = {'Descrição':'funcao', 'Código':'codigo', 'Empenhado':'empenhado', 'Pago':'pago'}, inplace = True)
    
    # --- Stacking datasets --- #
    if x == 0:    
        dataset_full = dataset
    else:
        dataset_full = pd.concat([dataset_full, dataset])



# ======================================== #
# === Adjustments for Cumulative Data  === #
# ======================================== #

# --- Sorting --- #
dataset_full = dataset_full.sort_values(by = ['funcao', 'tipo', 'ano', 'mes']).reset_index(drop = True)

# --- Cumulative data adjustment --- #
dataset_full['empenhado_mensal'] = dataset_full['empenhado'] - dataset_full['empenhado'].shift(1)       # Inserting adjusted values
dataset_full['pago_mensal'] = dataset_full['pago'] - dataset_full['pago'].shift(1)                      # Inserting adjusted values

# --- Loop for adjusting some truncated values --- #
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

# --- Vertical dataset adjustment --- #
dataset_full = dataset_full.melt(
    id_vars = ['periodo', 'ano', 'mes', 'tipo', 'codigo', 'funcao'],
    value_vars = ['empenhado_acumulado', 'pago_acumulado', 'empenhado_mensal', 'pago_mensal'],
    var_name = 'categoria',
    value_name = 'valor'
    )

# --- Storing --- #
with pd.ExcelWriter(path = 'investimentos_siof_ceara_funcao.xlsx', engine = 'xlsxwriter') as writer:
    dataset_full.to_excel(excel_writer = writer, sheet_name = 'investimentos_funcao', index = False)
    
    # Quick Formatting
    workbook = writer.book
    worksheet = writer.sheets['investimentos_funcao']
    money_formatting = workbook.add_format({'num_format':'R$#,##0'})
    perc_formatting = workbook.add_format({'num_format':'0.0%'})
    worksheet.set_column('G:J', 15, money_formatting)    
    worksheet.set_column('K:L', 15, perc_formatting)


# --- Full Cleasing --- #
del(dataset, folder_files, money_formatting, perc_formatting, workbook, worksheet, writer, x)