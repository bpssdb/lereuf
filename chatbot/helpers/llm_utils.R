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

library(stringr)
library(tokenizers)  # pour tokenize_sentences()

sanitize_for_sentences <- function(text) {
  # Ajouter un point après chaque ligne qui ne se termine pas par un signe de ponctuation
  lines <- unlist(strsplit(text, "\n"))
  lines <- ifelse(grepl("[\\.!?]$", lines), lines, paste0(lines, "."))
  paste(lines, collapse = " ")
}

attach_source_phrases <- function(budget_df, full_text) {
  txt <- sanitize_for_sentences(full_text)
  sents <- tokenizers::tokenize_sentences(txt)[[1]]
  
  budget_df$SourcePhrase <- vapply(seq_len(nrow(budget_df)), function(i) {
    row     <- budget_df[i, ]
    montant <- as.character(row$Montant)
    mot     <- tolower(str_extract(row$Description, "^[^ ]+"))
    
    # --- construire regex montant (5000 ou 5 000)
    num_pat <- str_replace_all(montant,
                               "(?<=\\d)(?=(?:\\d{3})+$)", "[ \\u00A0]?")  
    num_pat <- paste0("(?<!\\d)", num_pat, "(?!\\d)")
    
    # --- variantes d’unité (millions, milliards…)
    unit_variants <- c("M€","M €","million(?:s)?(?: d['’]?euros?)?",
                       "Md€","Md €","milliard(?:s)?(?: d['’]?euros?)?")
    unit_pat <- paste0("\\s*(?:", str_c(unit_variants, collapse="|"), ")")
    
    # --- priorités de recherche
    patterns <- list(
      paste0(num_pat, unit_pat, ".*?\\b", mot, "\\b"),  # montant+unité+mot
      paste0(num_pat, unit_pat),                        # montant+unité
      paste0(num_pat, ".*?\\b", mot, "\\b"),            # montant+mot
      num_pat                                           # montant seul
    )
    
    # 3) pour chaque pattern, chercher dans les phrases
    for (pat in patterns) {
      re <- regex(pat, ignore_case = TRUE)
      hit <- sents[str_detect(sents, re)]
      if (length(hit) > 0) return(hit[1])
    }
    
    # 4) rien trouvé → NA
    NA_character_
  }, character(1), USE.NAMES = FALSE)
  
  budget_df
}

# en haut de votre script, juste après vos library()
# utilitaire pour convertir "AB12" → list(row=12, col=28)
col2num <- function(col) {
  letters <- strsplit(col, "")[[1]]
  sum(sapply(seq_along(letters), function(i) {
    match(letters[i], LETTERS) * 26^(length(letters) - i)
  }))
}
parse_address <- function(addr) {
  m <- regexec("^([A-Z]+)([0-9]+)$", addr)
  parts <- regmatches(addr, m)[[1]]
  list(
    row = as.integer(parts[3]),
    col = col2num(parts[2])
  )
}
