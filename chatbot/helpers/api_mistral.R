## helpers.R: modules pour appels LLM et traitement budgétaire

# Librairies
library(httr)
library(jsonlite)
library(glue)
library(logger)

# -- CONFIGURATION & UTILITAIRES --
# URL de l'API Mistral
mistral_api_url <- "https://api.mistral.ai/v1/chat/completions"

# Clé API récupérée depuis une variable globale ou l'environnement
get_api_key <- function() {
  if (exists("mistral_api_key", envir = .GlobalEnv)) {
    return(get("mistral_api_key", envir = .GlobalEnv))
  }
  key <- Sys.getenv("MISTRAL_API_KEY")
  if (nzchar(key)) return(key)
  log_error("Clé API Mistral introuvable : ni 'mistral_api_key' ni la variable d'env MISTRAL_API_KEY.")
  stop("Clé API Mistral manquante")
}

# Fonction interne pour envoyer un appel chat à l'API LLM
llm_chat <- function(messages, model = "mistral-small-latest", verbose = FALSE) {
  api_key <- tryCatch(get_api_key(), error = function(e) {
    log_error(e$message)
    NULL
  })
  if (is.null(api_key)) return(NULL)
  
  if (verbose) message("[LLM] Envoi requête avec modèle '", model, "' et ", length(messages), " messages...")
  
  # Exécute la requête POST, avec ou sans httr::verbose()
  resp <- tryCatch({
    if (verbose) {
      POST(
        url = mistral_api_url,
        add_headers(
          Authorization = paste("Bearer", api_key),
          `Content-Type` = "application/json"
        ),
        body = list(model = model, messages = messages),
        encode = "json",
        timeout(60),
        verbose()
      )
    } else {
      POST(
        url = mistral_api_url,
        add_headers(
          Authorization = paste("Bearer", api_key),
          `Content-Type` = "application/json"
        ),
        body = list(model = model, messages = messages),
        encode = "json",
        timeout(60)
      )
    }
  }, error = function(e) {
    log_error(sprintf("Erreur HTTP LLM: %s", e$message))
    NULL
  })
  
  if (is.null(resp)) return(NULL)
  stat <- status_code(resp)
  raw_text <- content(resp, as = "text", encoding = "UTF-8")
  if (stat != 200) {
    log_error(sprintf("API LLM status %s: %s", stat, raw_text))
    return(NULL)
  }
  
  parsed <- tryCatch({
    fromJSON(raw_text, simplifyVector = FALSE)
  }, error = function(e) {
    log_error(sprintf("Impossible de parser JSON LLM: %s", e$message))
    NULL
  })
  
  if (!is.list(parsed) || is.null(parsed$choices) || length(parsed$choices) < 1) {
    log_error("Réponse inattendue LLM : pas de champ 'choices'.")
    return(NULL)
  }
  parsed$choices[[1]]$message$content
}

# Nettoyage de JSON encodé en markdown
clean_json_blob <- function(text) {
  if (is.null(text) || length(text) == 0) return("")
  cleaned <- gsub("```json", "", text, fixed = TRUE)
  cleaned <- gsub("```", "", cleaned, fixed = TRUE)
  trimws(cleaned)
}

# Extraction du premier tableau JSON dans un texte
extract_json_array <- function(text) {
  if (is.null(text) || length(text) == 0) return(NULL)
  m <- regexpr("\\[.*?\\]", text, perl = TRUE)
  if (m == -1) return(NULL)
  substring(text, m, m + attr(m, "match.length") - 1)
}

# -- PREPROMPT Global --
preprompt <- list(
  role = "system",
  content = paste(
    "Vous êtes BudgiBot, un assistant intelligent dédié à la Direction du Budget française.",
    "Vos réponses doivent être concises, professionnelles et adaptées à un public expert.",
    "Si l'utilisateur envoie un fichier, proposez une synthèse en deux lignes et demandez ce qu'il attend.",
    "Suggérez l'outil BPSS si l'utilisateur parle de remplir un fichier ou produire un fichier final."
  )
)

# ----------------------------------------------------------
# get_mistral_response: chat générique
# ----------------------------------------------------------
get_mistral_response <- function(chat_history, model = NULL) {
  msgs <- append(list(preprompt), chat_history)
  resp <- llm_chat(msgs, if (is.null(model)) "mistral-small-latest" else model, verbose = TRUE)
  if (is.null(resp)) {
    return(
      "Bien pris, je t'en remercie vivement ! \n Bien à toi."
    )
  }
  resp
}

# ----------------------------------------------------------
# analyze_labels: identifie axes et contexte via LLM
# ----------------------------------------------------------
analyze_labels <- function(labels) {
  if (length(labels) == 0) {
    message("❌ Aucun label fourni pour analyse.")
    return(NULL)
  }
  labels_txt <- paste(labels, collapse = ", ")
  prompt_sys <- paste(
    "Tu es un assistant budgétaire.",
    "Analyse la liste de labels suivants et identifie les axes pertinents et leur contexte pour une extraction budgétaire.",
    "Labels :", labels_txt,
    "Retourne STRICTEMENT un JSON avec deux champs :",
    "{\"axes\": [...], \"contexte_general\": \"...\"}",
    "NE FOURNIS PAS D'EXPLICATION HORS JSON."
  )
  raw <- llm_chat(list(list(role = "system", content = prompt_sys)))
  if (is.null(raw)) return(NULL)
  cleaned <- clean_json_blob(raw)
  tryCatch({
    fromJSON(cleaned, simplifyVector = TRUE)
  }, error = function(e) {
    message("❌ Erreur JSON analyse_labels: ", e$message)
    NULL
  })
}

# ----------------------------------------------------------
# get_budget_data: extraction de données budgétaires via LLM
# ----------------------------------------------------------
get_budget_data <- function(content_text, axes = NULL) {
  if (!is.null(axes)) {
    axes_txt <- if (is.data.frame(axes)) {
      paste(glue("{axe}: {description}"), collapse = "; ")
    } else if (is.list(axes)) {
      paste(sapply(axes, function(x) glue("{x$axe}: {x$description}")), collapse = "; ")
    } else {
      as.character(axes)
    }
    example <- toJSON(lapply(seq_along(axes), function(i) list(
      Axe = axes[[i]]$axe,
      Description = axes[[i]]$description,
      Montant = 0,
      Variations_emplois = +250, 
      Unité = "€",
      Probabilite = 0.0,
      Nature = ""
    )), auto_unbox = TRUE)
    prompt_sys <- paste(
      "Tu es un assistant budgétaire.",
      "En te basant sur ces axes :", axes_txt,
      "Retourne UNIQUEMENT un tableau JSON des données budgétaires.",
      "Exemple format :", example,
      "NE FOURNIS RIEN D'AUTRE."
    )
  } else {
    prompt_sys <- paste(
      "Tu es un assistant budgétaire.",
      "Analyse le texte fourni et retourne UNIQUEMENT un tableau JSON avec les données budgétaires détectées au format :",
      "[ {\"Axe\":..., \"Description\":..., \"Montant\":..., \"Variations_emplois\":..., \"Unité\":..., \"Probabilite\":..., \"Nature\":...} ]"
    )
  }
  msgs <- list(
    list(role = "system", content = prompt_sys),
    list(role = "user", content = content_text)
  )
  raw <- llm_chat(msgs)
  if (is.null(raw)) return(NULL)
  cleaned <- clean_json_blob(raw)
  blob <- extract_json_array(cleaned)
  if (is.null(blob)) {
    message("Aucune donnée budgétaire détectée.")
    return(NULL)
  }
  fromJSON(blob, simplifyVector = TRUE)
}

# ----------------------------------------------------------
# select_tag_for_entry: choisit le tag Excel le plus adapté Normalement n'est pas utilisé. 
# ----------------------------------------------------------
select_tag_for_entry <- function(entry, tags_json) {
  if (length(tags_json) == 0) return(NA_character_)
  txt <- paste0(seq_along(tags_json), ") ", sapply(tags_json, function(t) paste0(
    "[cell=", t$cell_address, "] labels=", paste(t$labels, collapse = "; ")
  )), collapse = "\n")
  user_p <- paste0(
    "Parmi ces tags, choisis le numéro qui correspond le mieux à Axe='", entry$Axe, "'",
    if (!is.null(entry$Description)) paste0(" et Description='", entry$Description, "'"),
    ".\nTags :\n", txt,
    "\nRéponds UNIQUEMENT par ce numéro."
  )
  resp <- llm_chat(c(list(preprompt), list(list(role="user", content=user_p))))
  idx <- suppressWarnings(as.integer(gsub("[^0-9]", "", resp)))
  if (!is.na(idx) && idx >= 1 && idx <= length(tags_json)) {
    tags_json[[idx]]$cell_address
  } else {
    NA_character_
  }
}