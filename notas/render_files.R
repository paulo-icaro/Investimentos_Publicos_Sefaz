# ==================== #
# === RENDER FILES === #
# ==================== #

# --- Script by: Paulo Icaro --- #


# ------------------- #
# --- Bibliotecas --- #
# ------------------- #
library(quarto)


# -------------------- #
# --- Renderização --- #
# -------------------- #
files = c('/home/pauloicaro/Documentos/Repositorios/Investimentos_Publicos_Sefaz/notas/data_processing_investimentos_funcao.qmd', '/home/pauloicaro/Documentos/Repositorios/Investimentos_Publicos_Sefaz/notas/data_processing_investimentos_programa_regiao.qmd')
formats = c('pdf', 'gfm')

for (i in seq_along(files)){
  for (j in seq_along(formats)){
    quarto_render(input = files[i], output_format = formats[j])
  }
}


# --------------- #
# --- Limpeza --- #
# --------------- #
rm(files, formats, i, j)