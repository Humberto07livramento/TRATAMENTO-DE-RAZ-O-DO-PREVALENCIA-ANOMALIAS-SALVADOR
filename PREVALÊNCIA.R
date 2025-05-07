# Instalar e carregar os pacotes necessários (caso não estejam instalados)
if (!require("writexl")) install.packages("writexl", dependencies = TRUE)
if (!require("ggplot2")) install.packages("ggplot2", dependencies = TRUE)
library(writexl)
library(ggplot2)

# Criação dos dados simulados
dados <- data.frame(
  Municipio      = c("São Luís", "São Luís", "Imperatriz", "Imperatriz", "Timon", "Timon"),
  Ano            = c(2001, 2002, 2001, 2002, 2001, 2002),
  NascidosVivos  = c(2500, 2600, 1500, 1600, 1000, 1050),
  AC_Sim         = c(50, 55, 30, 35, 20, 22),
  AC_Nao         = c(2430, 2535, 1455, 1550, 970, 1018),
  AC_Ignorado    = c(20, 10, 15, 15, 10, 10)
)

# Cálculo do Denominador Ajustado: NascidosVivos - AC_Ignorado
dados$Denom_Ajustado <- dados$NascidosVivos - dados$AC_Ignorado

# Cálculo da Prevalência (por 10.000 NV): (AC_Sim / Denom_Ajustado) * 10000
dados$Prevalencia <- (dados$AC_Sim / dados$Denom_Ajustado) * 10000

# Cálculo dos Intervalos de Confiança (IC 95%) utilizando a aproximação normal
dados$IC_LI <- dados$Prevalencia - 1.96 * sqrt(dados$Prevalencia / dados$Denom_Ajustado) * sqrt(10000)
dados$IC_LS <- dados$Prevalencia + 1.96 * sqrt(dados$Prevalencia / dados$Denom_Ajustado) * sqrt(10000)

# Arredondar os valores para duas casas decimais
dados$Prevalencia <- round(dados$Prevalencia, 2)
dados$IC_LI       <- round(dados$IC_LI, 2)
dados$IC_LS       <- round(dados$IC_LS, 2)

# Exibe os dados no console
print(dados)

# Definindo o diretório de exportação na pasta "Documentos" do Windows
dest_dir <- file.path(path.expand("~"), "Documents")

# Exportação dos dados para um arquivo Excel na pasta "Documentos"
write_xlsx(dados, file.path(dest_dir, "resultado.xlsx"))

# Criação de um gráfico moderno com ggplot2, separando os municípios em painéis
grafico <- ggplot(dados, aes(x = as.factor(Ano), y = Prevalencia)) +
  geom_line(aes(group = 1), color = "steelblue", size = 1.2) +
  geom_point(aes(color = Municipio), size = 3) +
  geom_errorbar(aes(ymin = IC_LI, ymax = IC_LS, color = Municipio), width = 0.2, size = 0.8) +
  facet_wrap(~ Municipio) +
  theme_minimal() +
  labs(title = "Prevalência de AC por Município e Ano",
       subtitle = "Dados simulados com intervalos de confiança (IC 95%)",
       x = "Ano",
       y = "Prevalência (por 10.000 NV)",
       color = "Município") +
  theme(
    plot.title = element_text(face = "bold", size = 16, hjust = 0.5),
    plot.subtitle = element_text(size = 12, hjust = 0.5),
    strip.text = element_text(face = "bold", size = 12),
    legend.position = "top"
  )

# Exibe o gráfico na janela de plotagem
print(grafico)

# Exporta o gráfico para um arquivo de imagem (PNG) na pasta "Documentos"
ggsave(filename = file.path(dest_dir, "grafico_moderno.png"), plot = grafico, width = 8, height = 6, dpi = 300)

