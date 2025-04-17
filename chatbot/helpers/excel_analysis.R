# helpers/excel_analysis.R

# 📊 Analyse de structure de feuille Excel enrichie

# Détection de la structure globale (dimensions, headers, colonnes vides)
detect_structure <- function(df) {
  n_cols <- ncol(df)
  n_rows <- nrow(df)
  non_empty_ratio <- mean(!is.na(df))
  empty_cols <- which(colSums(is.na(df)) > 0.95 * nrow(df))
  has_column_headers <- all(sapply(df[1, ], function(x) is.character(x) || is.factor(x)))
  
  structure_summary <- list(
    n_rows = n_rows,
    n_cols = n_cols,
    empty_cols = empty_cols,
    non_empty_ratio = round(non_empty_ratio, 2),
    headers_detected = has_column_headers
  )
  return(structure_summary)
}

# Détection des blocs sémantiques : années, montants, textes

get_column_examples <- function(df, cols, n = 3) {
  sapply(cols, function(i) {
    colname <- names(df)[i]
    vals <- na.omit(unique(df[[i]]))
    vals <- head(vals, n)
    paste0("- `", colname, "` (ex: ", paste(vals, collapse = ", "), ")")
  })
}

summarize_montant_distribution <- function(df, montant_cols) {
  sapply(montant_cols, function(i) {
    col <- df[[i]]
    col <- suppressWarnings(as.numeric(col))
    col <- col[!is.na(col)]
    if (length(col) == 0) return(NA)
    paste0("- `", names(df)[i], "` : Min=", format(min(col)), ", Max=", format(max(col)), ", Médiane=", format(median(col)))
  })
}

detect_semantic_blocks <- function(df) {
  colnames_lower <- tolower(colnames(df))
  years <- grep("^20[2-3][0-9]$", colnames(df), value = TRUE)
  montant_cols <- which(sapply(df, function(x) is.numeric(x) && suppressWarnings(max(x, na.rm = TRUE) > 1000)))
  texte_cols <- which(sapply(df, function(x) is.character(x) || is.factor(x)))
  
  objectifs_cols <- grep("objectif|indicateur|résultat|taux", colnames_lower, value = TRUE)
  code_cols <- grep("code|categ|prog", colnames_lower, value = TRUE)
  
  list(
    colonnes_annees = years,
    colonnes_montants = montant_cols,
    colonnes_textes = texte_cols,
    colonnes_objectifs = objectifs_cols,
    colonnes_codes = code_cols
  )
}

# Synthèse interprétable pour un LLM
analyser_feuille_excel <- function(df, sheet_name = "Feuille") {
  structure <- detect_structure(df)
  blocks <- detect_semantic_blocks(df)
  
  montant_stats <- summarize_montant_distribution(df, blocks$colonnes_montants)
  exemples_texte <- get_column_examples(df, blocks$colonnes_textes)
  
  summary <- paste0(
    "🧾 **Analyse de la feuille `", sheet_name, "` :**\n",
    "- Dimensions : ", structure$n_rows, " lignes × ", structure$n_cols, " colonnes\n",
    "- En-têtes détectées : ", ifelse(structure$headers_detected, "✅ Oui", "❌ Non"), "\n",
    "- Colonnes vides : ", length(structure$empty_cols), "\n",
    "- Ratio de cellules remplies : ", structure$non_empty_ratio * 100, "%\n\n",
    
    "🔎 **Types sémantiques :**\n",
    if (length(blocks$colonnes_annees)) paste0("- 📅 Années : ", paste(blocks$colonnes_annees, collapse = ", "), "\n") else "",
    if (length(blocks$colonnes_objectifs)) paste0("- 🎯 Objectifs : ", paste(blocks$colonnes_objectifs, collapse = ", "), "\n") else "",
    if (length(blocks$colonnes_codes)) paste0("- 🧾 Codes : ", paste(blocks$colonnes_codes, collapse = ", "), "\n") else "",
    
    if (length(montant_stats)) paste0("- 💶 Montants :\n", paste(montant_stats, collapse = "\n"), "\n") else "",
    if (length(exemples_texte)) paste0("- ✏️ Textes :\n", paste(exemples_texte, collapse = "\n"), "\n") else "",
    
    "\n🔍 **Extrait du tableau :**\n",
    paste(capture.output(print(utils::head(df, 5))), collapse = "\n")
  )
  
  return(summary)
}
