map_budget_entries <- function(entries_df, tags_json, model = NULL) {
  if (nrow(entries_df) == 0 || length(tags_json) == 0) return(NULL)
  
  # 1. Formatage des entrées budgétaires
  entries_txt <- apply(entries_df, 1, function(row) {
    glue::glue("- Axe: '{row[['Axe']]}', Description: '{row[['Description']]}', Montant: {row[['Montant']]}")
  }) %>% paste(collapse = "\n")
  
  # 2. Formatage des tags disponibles
  tags_txt <- seq_along(tags_json) %>%
    purrr::map_chr(function(i) {
      tag <- tags_json[[i]]
      labs <- paste(tag$labels, collapse = ", ")
      glue::glue("{i}) Cellule {tag$cell_address} – labels : {labs} – ID: {tag$id}")
    }) %>% paste(collapse = "\n")
  
  # 3. Prompt enrichi
  system_prompt <- paste(
    "Tu es un assistant budgétaire expert.",
    "Pour chaque entrée budgétaire, choisis la cellule Excel la plus pertinente parmi la liste de tags.",
    "Réponds exclusivement avec un tableau JSON au format suivant :",
    "[",
    "  {",
    "    \"Axe\": \"...\",",
    "    \"Description\": \"...\",",
    "    \"cellule\": \"B10\",",
    "    \"tags_utilisés\": [\"Label1\", \"Label2\"],",
    "    \"tag_id\": 4",
    "  },",
    "  ...",
    "]",
    "Assure-toi que les `tag_id` et `tags_utilisés` correspondent exactement à ceux présents dans les tags disponibles."
  )
  
  user_prompt <- paste(
    "Éléments budgétaires :\n", entries_txt,
    "\n\nTags disponibles :\n", tags_txt
  )
  
  msgs <- list(
    list(role = "system", content = system_prompt),
    list(role = "user",   content = user_prompt)
  )
  
  if (!is.null(model)) msgs[[1]]$model <- model
  
  # 4. Appel du LLM
  raw <- llm_chat(msgs)
  if (is.null(raw)) return(NULL)
  
  # 5. Extraction JSON
  blob <- extract_json_array(raw)
  if (is.null(blob)) {
    warning("Aucun JSON détecté dans la réponse LLM.")
    return(NULL)
  }
  
  tryCatch({
    mapping <- fromJSON(blob, simplifyDataFrame = TRUE)
    
    # Validation minimale
    required_keys <- c("Axe", "Description", "cellule", "tag_id", "tags_utilisés")
    if (!all(required_keys %in% names(mapping))) {
      warning("Clés attendues manquantes dans le JSON : ", paste(setdiff(required_keys, names(mapping)), collapse = ", "))
      return(NULL)
    }
    
    mapping
  }, error = function(e) {
    warning("Impossible de parser le JSON de mapping : ", e$message)
    NULL
  })
}
