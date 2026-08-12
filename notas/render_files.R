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
files = c('data_processing_investimentos_funcao.qmd', 'data_processing_investimentos_programa_regiao.qmd')
formats = c('pdf', 'gfm', 'html')

for (i in seq_along(files)){
  for (j in seq_along(formats)){
    quarto_render(input = files[i], output_format = formats[j])
  }
}


# --------------- #
# --- Limpeza --- #
# --------------- #
rm(files, formats, i, j)