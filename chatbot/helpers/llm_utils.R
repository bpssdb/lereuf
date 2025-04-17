verify_llm_passage <- function(full_text, passage) {
  # Normalisation basique (trim, caractères invisibles)
  normalized_text <- gsub("[\r\n\t]+", " ", full_text)
  normalized_text <- trimws(normalized_text)
  
  normalized_passage <- gsub("[\r\n\t]+", " ", passage)
  normalized_passage <- trimws(normalized_passage)
  
  grepl(fixed = TRUE, normalized_passage, x = normalized_text)
}

# 📌 Fonction pour extraire le bloc JSON proprement
extract_json_array <- function(text) {
  open_brackets <- gregexpr("\\[", text)[[1]]
  close_brackets <- gregexpr("\\]", text)[[1]]
  
  if (open_brackets[1] == -1 || close_brackets[1] == -1) {
    message("Aucune structure JSON détectée dans le texte.")
    return(NULL)
  }
  
  for (start in open_brackets) {
    for (end in rev(close_brackets)) {
      if (end > start) {
        json_candidate <- substr(text, start, end)
        if (startsWith(json_candidate, "[") && endsWith(json_candidate, "]")) {
          return(json_candidate)
        }
      }
    }
  }
  
  message("Impossible de trouver un tableau JSON correct.")
  return(NULL)
}
