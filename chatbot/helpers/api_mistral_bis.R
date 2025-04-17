# Définition d'un objet de type liste nommé "preprompt"
# Ce prompt sert de message système pour l'API LLM. Il définit le rôle de BudgiBot ainsi que ses consignes de réponse.
preprompt <- list(
  role = "system",
  content = "Vous êtes BudgiBot, un assistant intelligent dédié à la Direction du Budget française. Vos réponses doivent être concises, professionnelles et adaptées à un public expert. Si l'utilisateur envoie un fichier, proposez une synthèse en deux lignes et demandez ce qu'il attend de cet envoi. Tu peux également suggérer d'utiliser l'outil Excel Budgétaire (outil BPSS) si l'utilisateur parle de remplir un fichier, utiliser un PPES, BUD45 ou produire un fichier final. Si besoin, mentionne qu'un bouton est disponible. Enfin, tu as la possibilité de faire des extractions de données budgétaires via un bouton disponible dans le chat."
)

# ----------------------------------------------------------
# Fonction get_mistral_response
# ----------------------------------------------------------
# Cette fonction interroge l'API Mistral pour obtenir une réponse du LLM.
# Elle prend en paramètre "chat_history", qui est la liste des messages de la conversation.
get_mistral_response <- function(chat_history) {
  
  # Définition de l'URL de l'API Mistral.
  api_url <- "https://api.mistral.ai/v1/chat/completions"
  
  # Récupération de la clé API depuis la variable globale "mistral_api_key" qui a été initialisée dans global_chatbot.R.
  api_key <- mistral_api_key
  print(api_key)
  # Ajout du prompt système (preprompt) au début de l'historique de la conversation
  messages <- append(list(preprompt), chat_history)
  
  # Envoi d'une requête POST à l'API via httr
  response <- POST(
    url = api_url,
    httr::add_headers(
      # L'en-tête "Authorization" contient la clé API sous la forme "Bearer [clé]".
      `Authorization` = paste("Bearer", api_key),
      `Content-Type` = "application/json"
    ),
    body = list(
      # Spécifie le modèle utilisé par l'API
      model = "mistral-small-latest",
      # Envoie la liste de messages construite précédemment
      messages = messages
    ),
    encode = "json", 
    httr::verbose()  # Active le mode verbose pour afficher les détails de la requête HTTP (utile pour le débogage)
  )
  
  # Si la réponse HTTP a un code 200 (succès) alors on récupère et retourne la réponse
  if (status_code(response) == 200) {
    content <- content(response, "parsed")
    return(content$choices[[1]]$message$content)
  } else {
    # En cas d'erreur, afficher le message d'erreur et renvoyer une réponse par défaut
    error_content <- content(response, "text")
    print(paste("Error response:", error_content))
    return("Bien pris, je t'en remercie vivement !\nMes équipes te reviennent au plus vite,\nBien à toi.")
  }
}

# ----------------------------------------------------------
# Fonction analyze_labels
# ----------------------------------------------------------
# Cette fonction interroge l'API pour analyser une liste de labels et identifier les axes pertinents
# ainsi que leur contexte, en renvoyant un objet JSON structuré.
analyze_labels <- function(labels) {
  
  # Définition de l'URL de l'API
  api_url <- "https://api.mistral.ai/v1/chat/completions"
  
  # Récupération de la clé API via la variable globale mistral_api_key
  api_key <- mistral_api_key
  
  # Si aucun label n'est fourni, afficher un message et retourner NULL
  if (length(labels) == 0) {
    message("❌ Aucun label fourni pour analyse.")
    return(NULL)
  }
  
  # Concaténation des labels en une chaîne séparée par une virgule
  labels_text <- paste(labels, collapse = ", ")
  
  # Construction du prompt système pour l'analyse des labels :
  # Le LLM doit analyser la liste de labels et retourner un JSON contenant les axes et le contexte général.
  extraction_prompt <- list(
    list(
      role = "system",
      content = paste(
        "Tu es un assistant budgétaire.",
        "Analyse la liste de labels suivants et identifie les axes pertinents et leur contexte pour une extraction budgétaire.",
        "Voici les labels : ", labels_text,
        "Retourne un objet JSON contenant deux éléments :",
        "{",
        "  \"axes\": [",
        "    {\"axe\": \"Nom de l'axe\", \"description\": \"Description de ce que cet axe représente\"},",
        "    {\"axe\": \"Nom de l'axe 2\", \"description\": \"Description de cet axe\"}",
        "  ],",
        "  \"contexte_general\": \"Fournis ici un contexte général expliquant ce que ces axes représentent dans le cadre de l'analyse budgétaire.\"",
        "}",
        "NE FOURNIS PAS D'EXPLICATION EN DEHORS DU JSON."
      )
    )
  )
  
  # Envoi de la requête POST à l'API
  response <- httr::POST(
    url = api_url,
    httr::add_headers(
      `Authorization` = paste("Bearer", api_key),
      `Content-Type` = "application/json"
    ),
    body = list(
      model = "mistral-small-latest",
      messages = extraction_prompt
    ),
    encode = "json"
  )
  
  # En cas d'erreur de réponse (code différent de 200), afficher un message d'erreur et retourner NULL.
  if (httr::status_code(response) != 200) {
    message("❌ Erreur API LLM : ", httr::content(response, "text"))
    return(NULL)
  }
  
  # Extraction du contenu brut de la réponse
  raw_output <- httr::content(response, "parsed")$choices[[1]]$message$content
  
  # Affichage du contenu brut pour le débogage
  print("🔍 Réponse brute du LLM :")
  print(raw_output)
  
  # Nettoyage du texte reçu : suppression des balises markdown (```json et ```)
  cleaned_text <- raw_output %>%
    gsub("```json", "", ., fixed = TRUE) %>%
    gsub("```", "", ., fixed = TRUE) %>%
    trimws()
  
  # Tentative de conversion du texte nettoyé en JSON via jsonlite::fromJSON
  tryCatch({
    parsed_data <- jsonlite::fromJSON(cleaned_text)
    return(parsed_data)
  }, error = function(e) {
    message("❌ Erreur lors de la conversion de la réponse en JSON : ", e$message)
    return(NULL)
  })
}

# ----------------------------------------------------------
# Fonction get_budget_data
# ----------------------------------------------------------
# Cette fonction interroge l'API Mistral pour extraire les données budgétaires d'un texte fourni.
# Elle accepte un paramètre obligatoire "content_text" (texte à analyser)
# et un paramètre optionnel "axes" qui, s'il est fourni, est utilisé pour construire un prompt plus précis.
get_budget_data <- function(content_text, axes = NULL) {
  
  # Définition de l'URL de l'API et récupération de la clé API
  api_url <- "https://api.mistral.ai/v1/chat/completions"
  api_key <- mistral_api_key  # La clé API a été définie dans global_chatbot.R
  if (FALSE) {
    # Construction du prompt système en fonction de la présence ou non d'axes
    if (!is.null(axes)) {
      # Si l'argument axes est un data.frame
      if (is.data.frame(axes)) {
        axes_df <- axes
        # Mise en minuscules des noms de colonnes pour la cohérence
        colnames(axes_df) <- tolower(colnames(axes_df))
        # Construction d'une chaîne de texte listant les axes et leur description
        axes_text <- paste(apply(axes_df, 1, function(row) {
          paste0(as.character(row["axe"]), ": ", as.character(row["description"]))
        }), collapse = "; ")
        # Construction d'un exemple de tableau JSON avec les axes
        sample_entries <- apply(axes_df, 1, function(row) {
          paste0(
            "{",
            "\"Axe\": \"", as.character(row["axe"]), "\", ",
            "\"Description\": \"", as.character(row["description"]), "\", ",
            "\"Montant\": 0, ",
            "\"Unité\": \"€\", ",
            "\"Probabilite\": , ",
            "\"Nature\": \"\"",
            "}"
          )
        })
        sample_json <- paste0("[", paste(sample_entries, collapse = ", "), "]")
        
      } else if (is.list(axes)) {
        # Si axes est une liste d'objets
        axes_text <- paste(sapply(axes, function(x) {
          axe_val <- if (!is.null(x$axe)) as.character(x$axe) else if (!is.null(x$Axe)) as.character(x$Axe) else ""
          desc_val <- if (!is.null(x$description)) as.character(x$description) else if (!is.null(x$Description)) as.character(x$Description) else ""
          paste0(axe_val, ": ", desc_val)
        }), collapse = "; ")
        sample_entries <- sapply(axes, function(x) {
          axe_val <- if (!is.null(x$axe)) as.character(x$axe) else if (!is.null(x$Axe)) as.character(x$Axe) else ""
          desc_val <- if (!is.null(x$description)) as.character(x$description) else if (!is.null(x$Description)) as.character(x$Description) else ""
          paste0(
            "{",
            "\"Axe\": \"", axe_val, "\", ",
            "\"Description\": \"", desc_val, "\", ",
            "\"Montant\": 0, ",
            "\"Unité\": \"€\", ",
            "\"Probabilite\": , ",
            "\"Nature\": \"\"",
            "}"
          )
        })
        sample_json <- paste0("[", paste(sample_entries, collapse = ", "), "]")
        
      } else {
        # Si axes est fourni sous une autre forme (par exemple, une chaîne)
        axes_text <- as.character(axes)
        sample_json <- "[]"  
      }
    }
    # Construction du prompt système incluant les axes et un exemple de format attendu
    system_prompt <- paste(
      "Tu es un assistant budgétaire.",
      "En te basant sur les axes d'analyse suivants :", axes_text,
      "Analyse le texte fourni et retourne UNIQUEMENT un tableau JSON avec les données budgétaires détectées.",
      "Voici un exemple de format attendu basé sur les axes détectés :", sample_json,
      "La nature peut prendre deux valeurs : Flux d'effectifs ou Mesure catégorielle.",
      "NE FOURNIS PAS D'EXPLICATION EN DEHORS DU JSON."
    )
  } else {
    # Prompt par défaut lorsque aucun axe n'est fourni
    system_prompt <- paste(
      "Tu es un assistant budgétaire.",
      "Analyse le texte fourni et retourne UNIQUEMENT un tableau JSON avec les données budgétaires détectées,",
      "au format suivant :",
      "[",
      "{",
      "\"Axe\": \"\",",
      "\"Description\": \"\",",
      "\"Montant\": 0,",
      "\"Unité\": \"€\",",
      "\"Probabilite\": 0.9,",
      "\"Nature\": \"\"",
      "}",
      "]",
      "La nature peut prendre deux valeurs : Flux d'effectifs ou Mesure catégorielle.",
      "NE FOURNIS PAS D'EXPLICATION EN DEHORS DU JSON."
    )
  }
  
  # Construction du prompt complet sous forme de liste de messages pour l'API :
  # On inclut le message système construit ci-dessus et un message utilisateur contenant le texte à analyser.
  extraction_prompt <- list(
    list(
      role = "system",
      content = system_prompt
    ),
    list(
      role = "user",
      content = paste("Analyse ce texte budgétaire :\n\n", content_text)
    )
  )
  
  # Envoi de la requête POST à l'API Mistral via httr
  response <- httr::POST(
    url = api_url,
    httr::add_headers(
      `Authorization` = paste("Bearer", api_key),
      `Content-Type` = "application/json"
    ),
    body = list(
      model = "mistral-small-latest",
      messages = extraction_prompt
    ),
    encode = "json"
  )
  
  # Si le code de statut HTTP n'est pas 200 (succès), afficher un message d'erreur et retourner NULL
  if (httr::status_code(response) != 200) {
    message("Erreur API LLM 1 : ", httr::content(response, "text"))
    return(NULL)
  }
  
  # Extraction du contenu de la réponse : on récupère le texte renvoyé par l'API dans le premier choix
  raw_output <- httr::content(response, "parsed")$choices[[1]]$message$content
  
  # Nettoyage de la réponse : suppression des balises markdown utilisées pour formater le JSON (```json et ```)
  cleaned_text <- raw_output %>%
    gsub("```json", "", ., fixed = TRUE) %>%
    gsub("```", "", ., fixed = TRUE) %>%
    trimws()
  
  # Extraction du tableau JSON via la fonction auxiliaire extract_json_array
  json_candidate <- extract_json_array(cleaned_text)
  
  # Si aucun bloc JSON n'est détecté, afficher un message et retourner NULL
  if (is.null(json_candidate)) {
    message("Aucune donnée budgétaire détectée.")
    return(NULL)
  }
  
  # Conversion du JSON extrait en objet R (par exemple, data.frame ou liste)
  parsed_data <- jsonlite::fromJSON(json_candidate)
  
  # Retourne les données extraites
  return(parsed_data)
}

# -----------------------------------------------------------------------------
# Sélectionne, via LLM, le tag le plus adapté à une ligne de budget_data
# -----------------------------------------------------------------------------
select_tag_for_entry <- function(entry, tags_json) {
  # 1) Construire un mini-liste de tags sous forme textuelle
  tags_text <- paste0(
    seq_along(tags_json), ") ",
    sapply(tags_json, function(tag) {
      paste0("[cell=", tag$cell_address, "] labels=", paste(tag$labels, collapse = "; "))
    }),
    collapse = "\n"
  )
  
  # 2) Construire le prompt utilisateur
  user_prompt <- paste0(
    "Parmi la liste de tags ci‑dessous, choisis **le numéro** du tag qui correspond le mieux ",
    "à l'axe '", as.character(entry$Axe), "'",
    if (!is.null(entry$Description)) paste0(" et à la description '", as.character(entry$Description), "'") else "",
    ".\n\nTags disponibles :\n", tags_text,
    "\n\nRéponds **uniquement** par ce numéro (ex. 3)."
  )
  
  # 3) Appel LLM
  llm_resp <- get_mistral_response(
    append(list(preprompt),
           list(list(role = "user", content = user_prompt)))
  )
  
  # 4) Extraire l’indice numérique dans la réponse
  idx <- as.integer(gsub("\\D", "", llm_resp))
  if (!is.na(idx) && idx >= 1 && idx <= length(tags_json)) {
    return(tags_json[[idx]]$cell_address)
  } else {
    # fallback en cas d’erreur de parsing LLM
    return(NA_character_)
  }
}

