# helper_budget_mapping.R

library(jsonlite)
library(glue)
library(purrr)

# Utilise votre fonction interne llm_chat() pour interroger Mistral
# et extract_json_array() pour repérer le JSON dans le texte renvoyé.
map_budget_entries <- function(entries_df, tags_json, model = NULL) {
  if (nrow(entries_df) == 0 || length(tags_json) == 0) return(NULL)
  
  # 1 Texte des entrées budgétaires
  entries_txt <- apply(entries_df, 1, function(row) {
    glue::glue("- Axe: '{row[['Axe']]}', Description: '{row[['Description']]}', Montant: {row[['Montant']]}")
  }) %>% paste(collapse = "\n")
  
  # 2 Texte des tags disponibles
  tags_txt <- seq_along(tags_json) %>%
    purrr::map_chr(function(i) {
      tag <- tags_json[[i]]
      labs <- paste(tag$labels, collapse = ", ")
      glue::glue("{i}) Cellule {tag$cell_address} – labels : {labs}")
    }) %>% paste(collapse = "\n")
  
  # 3️ Prompt système et utilisateur
  system_prompt <- paste(
    "Tu es un assistant budgétaire expert.",
    "Pour chaque entrée budgétaire, choisis la meilleure cellule Excel parmi la liste.",
    "Réponds UNIQUEMENT par un JSON array au format :",
    "[",
    "  {\"Axe\":\"...\",\"Description\":\"...\",\"cellule\":\"B10\"},",
    "  …",
    "]"
  )
  user_prompt <- paste(
    "Éléments budgétaires :\n", entries_txt,
    "\n\nTags disponibles :\n", tags_txt
  )
  
  # 4️ Appel LLM
  msgs <- list(
    list(role = "system", content = system_prompt),
    list(role = "user",   content = user_prompt)
  )
  if (!is.null(model)) msgs[[1]]$model <- model  # optionnel si vous gérez plusieurs modèles
  
  raw <- llm_chat(msgs)
  if (is.null(raw)) return(NULL)
  
  # 5️ Extraction et parsing du JSON
  blob <- extract_json_array(raw)
  if (is.null(blob)) {
    warning("Aucun JSON détecté dans la réponse LLM.")
    return(NULL)
  }
  
  tryCatch({
    mapping <- fromJSON(blob, simplifyDataFrame = TRUE)
    # Validation minimale
    if (!all(c("Axe","Description","cellule") %in% names(mapping))) {
      warning("Le JSON renvoyé ne contient pas les clés attendues.")
      return(NULL)
    }
    mapping
  }, error = function(e) {
    warning("Impossible de parser le JSON de mapping : ", e$message)
    NULL
  })
}
