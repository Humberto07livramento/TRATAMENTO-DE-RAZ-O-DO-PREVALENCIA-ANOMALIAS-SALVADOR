
# Instala e carrega os pacotes necessários
if (!require("knitr")) install.packages("knitr")
if (!require("writexl")) install.packages("writexl")
library(knitr)
library(writexl)

# Função para calcular a prevalência (por 10.000) e seu IC95%
calc_prev <- function(cases, births) {
  prev <- (cases / births) * 10000
  se <- (sqrt(cases) / births) * 10000
  lower <- prev - 1.96 * se
  upper <- prev + 1.96 * se
  return(c(prev = prev, lower = lower, upper = upper))
}

# Função para calcular a razão de prevalência, seu IC95% e valor p
calc_ratio <- function(ref_cases, ref_births, grp_cases, grp_births) {
  prev_ref <- (ref_cases / ref_births) * 10000
  prev_grp <- (grp_cases / grp_births) * 10000
  RR <- prev_grp / prev_ref
  se_log <- sqrt(1 / ref_cases + 1 / grp_cases)
  lower <- exp(log(RR) - 1.96 * se_log)
  upper <- exp(log(RR) + 1.96 * se_log)
  z <- log(RR) / se_log
  p <- 2 * (1 - pnorm(abs(z)))
  return(c(RR = RR, lower = lower, upper = upper, p = p))
}

# Função auxiliar para montar os resultados para uma variável
# 'var_name': nome da variável
# 'grp_names': vetor com nomes dos grupos (primeiro é sempre o de referência)
# 'n': vetor com número total de nascidos (na mesma ordem)
# 'cases': vetor com número de casos (na mesma ordem)
# 'ref_index': índice do grupo de referência (default 1)
monta_resultados <- function(var_name, grp_names, n, cases, ref_index = 1) {
  # Calcula prevalência para cada grupo
  prev_calc <- sapply(1:length(n), function(i) calc_prev(cases[i], n[i]))
  prev_calc <- t(prev_calc)  # matriz com colunas: prev, lower, upper
  
  # Separa os resultados para o grupo de referência e o comparado (quando há dois grupos)
  if(length(n) == 2) {
    ratio <- calc_ratio(cases[ref_index], n[ref_index], cases[-ref_index], n[-ref_index])
    RR_text <- c("1.00", sprintf("%.2f (%.2f–%.2f)", ratio["RR"], ratio["lower"], ratio["upper"]))
    p_text <- c(NA, sprintf("%.3f", ratio["p"]))
  } else {
    RR_text <- rep(NA, length(n))
    p_text <- rep(NA, length(n))
  }
  
  # Formata os valores de prevalência com 1 casa decimal
  prev_text <- sprintf("%.1f (%.1f–%.1f)", prev_calc[, "prev"], prev_calc[, "lower"], prev_calc[, "upper"])
  
  # Cria a tabela de resultados para a variável
  df <- data.frame(
    Variavel = var_name,
    Grupo = grp_names,
    n = n,
    Casos = cases,
    "Prevalência (IC95%)" = prev_text,
    "Razão de prevalência (IC95%)" = RR_text,
    "Valor p" = p_text,
    stringsAsFactors = FALSE
  )
  return(df)
}

# ----- Cálculos para cada variável (dados fictícios) -----

# 1. Idade da mãe
idade <- monta_resultados(
  var_name = "Idade da mãe",
  grp_names = c("< 35 anos", "≥ 35 anos"),
  n = c(5536, 537),
  cases = c(18, 27),
  ref_index = 1
)

# 2. Consultas no pré-natal (referência: ≥ 7 consultas)
consultas <- monta_resultados(
  var_name = "Consultas pré-natais",
  grp_names = c("≥ 7 consultas", "0 a 6 consultas"),
  n = c(1834, 4213),
  cases = c(7, 14),
  ref_index = 1
)

# 3. Estado civil (referência: Com companheiro)
estado_civil <- monta_resultados(
  var_name = "Estado civil",
  grp_names = c("Com companheiro", "Sem companheiro"),
  n = c(2608, 3390),
  cases = c(10, 10),
  ref_index = 1
)

# 4. Escolaridade (referência: ≥ 12 anos)
escolaridade <- monta_resultados(
  var_name = "Escolaridade materna",
  grp_names = c("≥ 12 anos", "0 a 11 anos"),
  n = c(537, 5441),
  cases = c(2, 18),
  ref_index = 1
)

# 5. Tipo de parto (referência: Vaginal)
parto <- monta_resultados(
  var_name = "Tipo de parto",
  grp_names = c("Vaginal", "Cesáreo"),
  n = c(3074, 3010),
  cases = c(8, 15),
  ref_index = 1
)

# 6. Sexo (referência: Feminino)
sexo <- monta_resultados(
  var_name = "Sexo",
  grp_names = c("Feminino", "Masculino"),
  n = c(2648, 3366),
  cases = c(8, 12),
  ref_index = 1
)

# 7. Peso ao nascer (referência: ≥ 2500 g)
peso <- monta_resultados(
  var_name = "Peso ao nascer",
  grp_names = c("≥ 2500 g", "< 2500 g"),
  n = c(4634, 1447),
  cases = c(13, 17),
  ref_index = 1
)

# 8. Apgar 5º minuto (referência: 8 a 10)
apgar <- monta_resultados(
  var_name = "Apgar 5º minuto",
  grp_names = c("8 a 10", "0 a 7"),
  n = c(4268, 1175),
  cases = c(14, 18),
  ref_index = 1
)

# Combina todas as tabelas em uma só
tabela_final <- rbind(idade, consultas, estado_civil, escolaridade, parto, sexo, peso, apgar)

# Exibe a tabela formatada no console
kable(tabela_final, 
      caption = "Tabela 2: Cálculos de prevalência, razão de prevalência e valor p (dados fictícios)",
      align = c("l", "l", "r", "r", "c", "c", "c"))

# Define o diretório para salvar o arquivo (ajuste se necessário)
dest_dir <- file.path(path.expand("~"), "Documents")

# Verifica se o diretório existe; se não, cria-o (incluindo subpastas, se necessário)
if (!dir.exists(dest_dir)) {
  dir.create(dest_dir, recursive = TRUE)
}

# Exporta a tabela para um arquivo Excel no diretório especificado
write_xlsx(tabela_final, file.path(dest_dir, "Tabela2_Resultados.xlsx"))
