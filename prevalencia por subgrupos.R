# Carregar pacotes
library(readxl)
library(dplyr)
library(stringr)
library(tidyr)
library(writexl)

# Caminhos
in_file    <- file.path(Sys.getenv("USERPROFILE"), "Documents", "extrato de seleção.xlsx")
out_file   <- file.path(Sys.getenv("USERPROFILE"), "Documents", "resumo_duas_abas.xlsx")

# Total de nascidos vivos
total_nascidos_vivos <- 444962

# 1) Ler e preparar 'dados'
dados <- read_excel(in_file, sheet = "extrato de seleção", skip = 2, col_names = FALSE) %>%
  setNames(c("Anomalia", "Total")) %>%
  separate(Anomalia, into = c("Codigo_CID10", "Descricao"), sep = " ", extra = "merge", fill = "right") %>%
  mutate(
    Codigo_CID10   = str_trim(Codigo_CID10),
    Descricao      = str_trim(Descricao),
    Total          = as.numeric(Total),
    Subgrupo_CID10 = str_extract(Codigo_CID10, "^[A-Z]\\d{2}")
  ) %>%
  mutate(Classificacao = case_when(
    Subgrupo_CID10 == "Q05"                            ~ "Espinha bífida (Q05)",
    Subgrupo_CID10 == "D18"                            ~ "Hemangioma e linfangioma de qualquer localização (D18)",
    Subgrupo_CID10 == "Q65"                            ~ "Deformidades congênitas do quadril (Q65)",
    Subgrupo_CID10 == "Q53"                            ~ "Testículo não descido (Q53)",
    Subgrupo_CID10 == "Q66"                            ~ "Deformidades congênitas dos pés (Q66)",
    Subgrupo_CID10 == "Q41"                            ~ "Ausência, atresia ou estenose do intestino delgado (Q41)",
    Subgrupo_CID10 >= "Q00" & Subgrupo_CID10 <= "Q07"  ~ "Sistema nervoso (Q00-Q07)",
    Subgrupo_CID10 >= "Q20" & Subgrupo_CID10 <= "Q28"  ~ "Aparelho circulatório (Q20-Q28)",
    Subgrupo_CID10 >= "Q35" & Subgrupo_CID10 <= "Q37"  ~ "Fenda labial e palatina (Q35-Q37)",
    Subgrupo_CID10 >= "Q38" & Subgrupo_CID10 <= "Q45"  ~ "Aparelho digestivo (Q38-Q45)",
    Subgrupo_CID10 >= "Q60" & Subgrupo_CID10 <= "Q64"  ~ "Aparelho urinário (Q60-Q64)",
    Subgrupo_CID10 >= "Q65" & Subgrupo_CID10 <= "Q79"  ~ "Aparelho osteomuscular (Q65-Q79)",
    Subgrupo_CID10 >= "Q80" & Subgrupo_CID10 <= "Q89"  ~ "Outras malformações congênitas (Q80-Q89)",
    Subgrupo_CID10 >= "Q90" & Subgrupo_CID10 <= "Q99"  ~ "Anomalias cromossômicas (Q90-Q99)",
    TRUE                                               ~ NA_character_
  ))

# 2) Construir o resumo 'prevalencia'
prevalencia <- dados %>%
  filter(!is.na(Classificacao)) %>%
  group_by(Classificacao) %>%
  summarise(Frequencia = sum(Total, na.rm = TRUE), .groups = "drop") %>%
  mutate(
    p       = Frequencia / total_nascidos_vivos,
    SE      = sqrt(p * (1 - p) / total_nascidos_vivos),
    Prev10k = p * 10000,
    IC_inf  = pmax((p - 1.96 * SE) * 10000, 0),
    IC_sup  = (p + 1.96 * SE) * 10000
  ) %>%
  mutate(
    Frequencia       = formatC(Frequencia, format = "d", big.mark = ".", decimal.mark = ","),
    Prev10k          = formatC(Prev10k,  format = "f", digits = 1, decimal.mark = ","),
    IC_inf           = formatC(IC_inf,   format = "f", digits = 1, decimal.mark = ","),
    IC_sup           = formatC(IC_sup,   format = "f", digits = 1, decimal.mark = ","),
    Prevalencia_IC95 = paste0(Prev10k, " (", IC_inf, " – ", IC_sup, ")")
  ) %>%
  select(Classificacao, Frequencia, Prevalencia_IC95)

# 3) Salvar em um único Excel com duas abas
sheets <- list(
  dados       = dados,
  prevalencia = prevalencia
)
write_xlsx(sheets, out_file)

cat("✔️ Arquivo com duas abas salvo em:\n  ", out_file, "\n")

