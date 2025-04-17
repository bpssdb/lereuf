# helpers/server_utils.R

# 🔹 Gère l’arrivée d’un fichier utilisateur : lecture, stockage, synthèse via LLM, mise à jour de l'historique
handle_file_message <- function(file_input, chat_history, dernier_fichier_contenu, session, typing) {
  file_name <- file_input$name
  file_path <- file_input$datapath
  
  file_content <- read_file_content(file_path, file_name)
  dernier_fichier_contenu(file_content)
  
  messages <- chat_history()
  messages <- append(messages, list(list(role = "user", content = paste("Fichier envoyé :", file_name))))
  chat_history(messages)
  
  session$sendCustomMessage(type = 'scrollToBottom', message = list())
  typing(TRUE)
  
  # 🔒 Nettoyage : on retire les champs non autorisés (comme "type")
  clean_messages <- lapply(messages, function(msg) {
    list(role = msg$role, content = msg$content)
  })
  
  messages_for_api <- append(clean_messages, list(
    list(role = "user", content = file_content),
    list(role = "system", content = "L'utilisateur a envoyé un fichier. Propose une synthèse en deux lignes et demande ce qu'il attend de cet envoi.")
  ))
  
  
  future({ get_mistral_response(messages_for_api) }) %...>% (function(response) {
    typing(FALSE)
    messages <- chat_history()
    messages <- append(messages, list(list(role = "assistant", content = response)))
    chat_history(messages)
    session$sendCustomMessage(type = 'scrollToBottom', message = list())
  }) %...!% (function(err) {
    typing(FALSE)
    error_message <- "Une erreur s'est produite lors du traitement du fichier."
    messages <- chat_history()
    messages <- append(messages, list(list(role = "assistant", content = error_message)))
    chat_history(messages)
    session$sendCustomMessage(type = 'scrollToBottom', message = list())
  })
}

# 🔹 Génère dynamiquement la liste des fichiers envoyés par l’utilisateur
render_file_list_ui <- function(messages) {
  files <- sapply(messages, function(m) {
    if (m$role == "user" && grepl("^Fichier envoyé :", m$content)) {
      gsub("^Fichier envoyé :\\s*", "", m$content)
    } else {
      NULL
    }
  })
  
  files <- files[!sapply(files, is.null)]
  
  if (length(files) == 0) {
    return(HTML("<p style='color: #888;'>Aucun fichier envoyé</p>"))
  }
  
  file_list_html <- paste0(
    "<ul style='list-style: none; padding: 0;'>",
    paste0(
      "<li style='display: flex; align-items: center; padding: 8px 0; border-bottom: 1px solid #eee;'>",
      "<i class='fas fa-file-alt' style='margin-right: 10px; font-size: 18px; color: #0055A4;'></i>",
      "<span style='font-size: 14px;'>", files, "</span>",
      "</li>",
      collapse = ""
    ),
    "</ul>"
  )
  
  HTML(file_list_html)
}

handle_user_input <- function(input, session, chat_history, typing, bpss_prompt_active) {
  req(input$user_input != "", nchar(input$user_input) > 1)
  
  messages <- chat_history()
  user_message <- input$user_input
  
  new_user_msg <- list(role = "user", content = user_message)
  updated_messages <- append(messages, list(new_user_msg))
  chat_history(updated_messages)
  
  updateTextAreaInput(session, "user_input", value = "")
  typing(TRUE)
  
  # 🔒 Nettoyage des messages avant appel API
  api_messages <- lapply(updated_messages, function(msg) {
    list(role = msg$role, content = msg$content)
  })
  
  future({
    get_mistral_response(api_messages)
  }) %...>% (function(response) {
    typing(FALSE)
    
    updated_messages <- chat_history()
    new_bot_msg <- list(role = "assistant", content = response)
    updated_messages <- append(updated_messages, list(new_bot_msg))
    
    if (grepl("outil BPSS|remplir.*outil|fichier.*PPES|BUD45|excel.*outil|générer fichier final", user_message, ignore.case = TRUE)) {
      prompt_msg <- list(
        role = "assistant",
        content = "Souhaitez-vous lancer l'outil BPSS ?",
        type = "bpss_prompt"
      )
      updated_messages <- append(updated_messages, list(prompt_msg))
      bpss_prompt_active(TRUE)
    }
    
    chat_history(updated_messages)
    session$sendCustomMessage(type = 'scrollToBottom', message = list())
    
  }) %...!% (function(err) {
    typing(FALSE)
    showNotification(paste("Erreur lors de l'appel à l'API :", err$message), type = "error")
  })
}

get_user_message_for_extraction <- function(index_of_bot, chat_history, dernier_fichier_contenu) {
  messages <- chat_history()
  
  user_index <- NA
  for (j in seq(index_of_bot - 1, 1)) {
    if (messages[[j]]$role == "user") {
      user_index <- j
      break
    }
  }
  
  if (is.na(user_index)) return(NULL)
  
  user_message <- if (grepl("^Fichier envoyé", messages[[user_index]]$content)) {
    dernier_fichier_contenu()
  } else {
    messages[[user_index]]$content
  }
  
  if (grepl("^Fichier envoyé :", user_message)) {
    file_content_index <- NA
    for (k in seq(user_index - 1, 1)) {
      if (!is.null(messages[[k]]$meta) && messages[[k]]$meta == "fichier_content") {
        file_content_index <- k
        break
      }
    }
    if (!is.na(file_content_index)) {
      user_message <- messages[[file_content_index]]$content
    } else {
      return(NULL)
    }
  }
  
  user_message
}

find_free_port <- function() {
  sample(3000:9000, 1)
}