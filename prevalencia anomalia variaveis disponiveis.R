# ===== 1. Instalação e carregamento de pacotes =====
if (!require("knitr"))   install.packages("knitr")
if (!require("writexl")) install.packages("writexl")
library(knitr)
library(writexl)

# ===== 2. Funções de cálculo =====
calc_prev <- function(cases, births) {
  prev  <- (cases / births) * 10000
  se    <- (sqrt(cases) / births) * 10000
  lower <- prev - 1.96 * se
  upper <- prev + 1.96 * se
  c(prev = prev, lower = lower, upper = upper)
}

calc_ratio <- function(ref_cases, ref_births, grp_cases, grp_births) {
  prev_ref <- (ref_cases / ref_births) * 10000
  prev_grp <- (grp_cases / grp_births) * 10000
  RR       <- prev_grp / prev_ref
  se_log   <- sqrt(1 / ref_cases + 1 / grp_cases)
  lower    <- exp(log(RR) - 1.96 * se_log)
  upper    <- exp(log(RR) + 1.96 * se_log)
  z        <- log(RR) / se_log
  p        <- 2 * (1 - pnorm(abs(z)))
  c(RR = RR, lower = lower, upper = upper, p = p)
}

monta_resultados <- function(var_name, grp_names, n, cases, ref_index = 1) {
  prev_mat <- t(sapply(seq_along(n), function(i) calc_prev(cases[i], n[i])))
  if (length(n) == 2) {
    r       <- calc_ratio(cases[ref_index], n[ref_index],
                          cases[-ref_index], n[-ref_index])
    RR_text <- c("1.00", sprintf("%.2f (%.2f–%.2f)", r["RR"], r["lower"], r["upper"]))
    p_text  <- c(NA, sprintf("%.3f", r["p"]))
  } else {
    RR_text <- rep(NA, length(n))
    p_text  <- rep(NA, length(n))
  }
  prev_text <- sprintf("%.1f (%.1f–%.1f)",
                       prev_mat[, "prev"], prev_mat[, "lower"], prev_mat[, "upper"])
  data.frame(
    Variavel                       = var_name,
    Grupo                          = grp_names,
    n                              = n,
    Casos                          = cases,
    "Prevalência (IC95%)"          = prev_text,
    "Razão de prevalência (IC95%)" = RR_text,
    "Valor p"                      = p_text,
    stringsAsFactors = FALSE
  )
}

# ===== 3. Cálculos para cada variável =====
idade <- monta_resultados("Idade da mãe",
                          c("< 35 anos","≥ 35 anos"), c(341451,86889), c(2446,782))
consultas <- monta_resultados("Consultas pré-natais",
                              c("≥ 7","0–6"), c(258420,181715), c(1793,1394))
estado_civil <- monta_resultados("Estado civil",
                                 c("Com companheiro","Sem companheiro"), c(325144,116723), c(2394,807))
escolaridade <- monta_resultados("Escolaridade materna",
                                 c("≥ 12 anos","0–11"), c(103089,338109), c(727,2469))
parto <- monta_resultados("Tipo de parto",
                          c("Vaginal","Cesáreo"), c(221687,222551), c(1407,1804))
sexo <- monta_resultados("Sexo",
                         c("Feminino","Masculino"), c(217589,227198), c(1271,1879))
peso <- monta_resultados("Peso ao nascer",
                         c("≥ 2500 g","< 2500 g"), c(398705,46247), c(2489,739))
apgar <- monta_resultados("Apgar 5º minuto",
                          c("8–10","0–7"), c(428854,12633), c(2908,296))

# ===== 4. Raça/Cor com dados reais de nascidos vivos =====
race_births <- c(Branca=45510, Preta=124127, Amarela=1787, Parda=260537, Indígena=531)
race_cases  <- c(Branca=286,   Preta=915,     Amarela=9,    Parda=1942,    Indígena=2)

# 4.1 Prevalência e IC
prev_list  <- lapply(seq_along(race_births),
                     function(i) calc_prev(race_cases[i], race_births[i]))
prev_mat   <- do.call(rbind, prev_list)
colnames(prev_mat) <- c("prev","lower","upper")

# 4.2 RR e IC vs Branca
ratio_list <- lapply(seq_along(race_births), function(i) {
  if (i==1) return(c(RR=1, lower=1, upper=1, p=NA))
  calc_ratio(race_cases[1], race_births[1],
             race_cases[i], race_births[i])
})
ratio_mat <- do.call(rbind, ratio_list)
colnames(ratio_mat) <- c("RR","lower","upper","p")

# 4.3 Formatação dos textos
prev_txt  <- sprintf("%.1f (%.1f–%.1f)",
                     prev_mat[,"prev"], prev_mat[,"lower"], prev_mat[,"upper"])
RR_txt    <- ifelse(ratio_mat[,"RR"]==1,
                    "1.00",
                    sprintf("%.2f (%.2f–%.2f)",
                            ratio_mat[,"RR"],
                            ratio_mat[,"lower"],
                            ratio_mat[,"upper"]))
p_txt     <- ifelse(is.na(ratio_mat[,"p"]), NA, sprintf("%.3f", ratio_mat[,"p"]))

race_df <- data.frame(
  Variavel                       = "Raça/Cor",
  Grupo                          = names(race_births),
  n                              = as.integer(race_births),
  Casos                          = as.integer(race_cases),
  "Prevalência (IC95%)"          = prev_txt,
  "Razão de prevalência (IC95%)" = RR_txt,
  "Valor p"                      = p_txt,
  stringsAsFactors = FALSE
)

# ===== 5. Montagem das duas tabelas =====
tabela_prevalencia <- do.call(rbind, 
                              list(idade, consultas, estado_civil, escolaridade,
                                   parto, sexo, peso, apgar, race_df)
)

build_abs_rel <- function(df) {
  with(df, data.frame(
    Variavel   = Variavel,
    Grupo      = Grupo,
    n          = n,
    Casos      = Casos,
    Percentual = sprintf("%.1f%%", (Casos/n)*100),
    stringsAsFactors = FALSE
  ))
}
tabela_absoluto_relativo <- do.call(rbind,
                                    lapply(list(idade, consultas, estado_civil, escolaridade,
                                                parto, sexo, peso, apgar, race_df),
                                           build_abs_rel)
)

# ===== 6. Exportação para Excel (duas abas) =====
docs_dir <- if (.Platform$OS.type == "windows") {
  file.path(Sys.getenv("USERPROFILE"), "Documents")
} else {
  file.path(path.expand("~"), "Documents")
}
if (!dir.exists(docs_dir)) dir.create(docs_dir, recursive = TRUE)

output_file <- file.path(docs_dir, "Resultados_Duas_Abas_RacaCorrigida.xlsx")
if (file.exists(output_file)) file.remove(output_file)

write_xlsx(
  list(
    Prevalencia       = tabela_prevalencia,
    Absoluto_Relativo = tabela_absoluto_relativo
  ),
  output_file
)
message("✅ Excel gerado em: ", normalizePath(output_file, winslash = "/", mustWork = FALSE))

