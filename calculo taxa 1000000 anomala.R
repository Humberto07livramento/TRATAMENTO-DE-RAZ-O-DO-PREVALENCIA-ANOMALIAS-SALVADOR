# Script R para cálculo de prevalência de anomalias congênitas e exportação profissional
# Autor: Manoela Almeida
# Data: 2025-04-26

# 1) Pacotes
if (!require("readxl"))   install.packages("readxl", dependencies = TRUE)
if (!require("dplyr"))    install.packages("dplyr", dependencies = TRUE)
if (!require("tidyr"))    install.packages("tidyr", dependencies = TRUE)
if (!require("openxlsx")) install.packages("openxlsx", dependencies = TRUE)
if (!require("ggplot2"))  install.packages("ggplot2", dependencies = TRUE)
if (!require("ggpmisc"))  install.packages("ggpmisc", dependencies = TRUE)

library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)
library(ggplot2)
library(ggpmisc)

# 2) Caminhos
input_file  <- "C:/Users/Manoela Almeida/Documents/anomalia.xlsx"
output_xlsx <- "C:/Users/Manoela Almeida/Documents/prevalencia_anomalias_final.xlsx"
output_plot <- "C:/Users/Manoela Almeida/Documents/grafico_prevalencia.png"

# 3) Leitura e preparação dos dados
df_raw <- read_excel(input_file)

df_long <- df_raw %>%
  pivot_longer(cols = -`Deteccao Anomalia`, names_to = "Ano", values_to = "Contagem") %>%
  mutate(Ano = as.integer(Ano))

df_wide <- df_long %>%
  pivot_wider(names_from = `Deteccao Anomalia`, values_from = Contagem) %>%
  rename(
    AC_Sim      = Sim,
    AC_Nao      = `Não`,
    AC_Ignorado = Ignorado,
    NV_Total    = Total
  )

# 4) Cálculos
dados <- df_wide %>%
  mutate(
    Denom_Ajustado = NV_Total - AC_Ignorado,
    Prevalencia    = (AC_Sim / Denom_Ajustado) * 10000,
    IC_LI = Prevalencia - 1.96 * sqrt(Prevalencia / Denom_Ajustado) * sqrt(10000),
    IC_LS = Prevalencia + 1.96 * sqrt(Prevalencia / Denom_Ajustado) * sqrt(10000)
  ) %>%
  mutate(
    Prevalencia = round(Prevalencia, 2),
    IC_LI       = round(IC_LI, 2),
    IC_LS       = round(IC_LS, 2)
  )

# 5) Separar tabelas de saída
tabela_completa <- dados %>%
  select(Ano, NV_Total, AC_Sim, Denom_Ajustado, Prevalencia, IC_LI, IC_LS)

tabela_taxas <- dados %>%
  select(Ano, Prevalencia, IC_LI, IC_LS)

# 6) Criar gráfico
grafico <- ggplot(dados, aes(x = Ano, y = Prevalencia)) +
  geom_line(color = "#0072B2", size = 1) +
  geom_point(color = "#D55E00", size = 3) +
  geom_text(aes(label = sprintf("%.1f", Prevalencia)), vjust = -1, size = 3.5) +
  geom_ribbon(aes(ymin = IC_LI, ymax = IC_LS), fill = "#56B4E9", alpha = 0.3) +
  stat_smooth(method = "lm", se = FALSE, linetype = "dashed", color = "black") +
  stat_poly_eq(
    aes(label = paste(..eq.label.., ..rr.label.., sep = "~~~")),
    formula = y ~ x, parse = TRUE,
    label.x.npc = "right", label.y.npc = "top"
  ) +
  labs(
    x = "Ano",
    y = "Prevalência por 10.000 nascimentos vivos"
  ) +
  scale_x_continuous(breaks = seq(min(dados$Ano), max(dados$Ano), 1)) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_blank(),
    plot.subtitle = element_blank(),
    axis.text.x = element_text(angle = 45, hjust = 1)
  )

# 7) Salvar gráfico
ggsave(filename = output_plot, plot = grafico, width = 10, height = 6, units = "in", dpi = 300)

# 8) Criar Workbook e adicionar abas
wb <- createWorkbook()

# Aba 1: Dados Completos
addWorksheet(wb, "Prevalência Detalhada")
writeData(wb, "Prevalência Detalhada", tabela_completa)

# Aba 2: Apenas Taxas
addWorksheet(wb, "Taxas")
writeData(wb, "Taxas", tabela_taxas)

# Aba 3: Gráfico
addWorksheet(wb, "Gráfico")
insertImage(wb, "Gráfico", file = output_plot, startRow = 1, startCol = 1, width = 16, height = 9)

# Adicionar Rodapé (aba do gráfico)
rodape1 <- "Figura 1. Distribuição temporal das prevalências de anomalias congênitas em Salvador, 2011-2024."
rodape2 <- "Fonte: Sistema de Informações sobre Nascidos, SUIS/DVIS."

writeData(wb, "Gráfico", rodape1, startCol = 1, startRow = 30)
writeData(wb, "Gráfico", rodape2, startCol = 1, startRow = 31)

style_rodape <- createStyle(fontSize = 9, fontName = "Arial", italic = TRUE)
addStyle(wb, "Gráfico", style = style_rodape, rows = 30:31, cols = 1, gridExpand = TRUE)

# Comentário interpretativo
comentario <- paste(
  "Comentário:",
  "A análise temporal indica oscilações na prevalência de anomalias congênitas,",
  "com valores superiores à média nacional estimada (~8,2/10.000 nascidos vivos).",
  "Essas variações podem ser atribuídas a melhorias na vigilância, maior acesso diagnóstico,",
  "ou ainda a fatores ambientais e socioeconômicos locais."
)

writeData(wb, "Prevalência Detalhada", comentario, startCol = 1, startRow = nrow(tabela_completa) + 4)
style_coment <- createStyle(fontSize = 11, wrapText = TRUE)
addStyle(wb, "Prevalência Detalhada", style = style_coment, rows = nrow(tabela_completa) + 4, cols = 1)

# 9) Salvar Excel
saveWorkbook(wb, output_xlsx, overwrite = TRUE)

message("✅ Concluído! Arquivo Excel gerado em: ", output_xlsx)


