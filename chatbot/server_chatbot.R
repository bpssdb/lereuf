library(shiny)
library(dplyr)


server <- function(input, output, session) {
  # RÉACTIFS INITIAUX
  chat_history <- reactiveVal(list())
  dernier_fichier_contenu <- reactiveVal(NULL)
  typing <- reactiveVal(FALSE)
  bpss_prompt_active <- reactiveVal(FALSE)
  
  rv <- reactiveValues(
    imported_json    = NULL,
    extracted_labels = NULL,
    axes             = NULL,
    context          = NULL
  )
  
  donnees_extraites <- reactiveVal(
    data.frame(
      Axe           = character(),
      Description   = character(),
      Montant       = numeric(),
      Unité         = character(),
      Probabilite   = numeric(),
      Feuille_excel = character(),
      SourcePhrase  = character(),
      CelluleCible  = character(),
      stringsAsFactors = FALSE
    )
  )
  
  # MODULES EXISTANTS
  mod_mesures_cat_server("cat1", rv, on_analysis_summary = function(summary) {
    msgs <- chat_history()
    msgs <- append(msgs, list(list(role = "assistant", content = summary)))
    chat_history(msgs)
    session$sendCustomMessage("scrollToBottom", list())
  })
  mod_outil_bpss_server("bpss1")
  mod_json_helper_server("json_helper", rv)
  setupBudgetExtraction(input, output, session,
                        rv, chat_history,
                        dernier_fichier_contenu,
                        donnees_extraites)
  mod_videoTranscriberServer(
    "vt1",
    trigger = reactive(input$show_video_modal)
  )
  
  
  # ENVOI UTILISATEUR
  observeEvent(input$send_btn, {
    handle_user_input(input, session, chat_history, typing, bpss_prompt_active)
  })
  
  
  # UPLOAD FICHIER
  observeEvent(input$file_input, {
    req(input$file_input)
    handle_file_message(input$file_input, chat_history, dernier_fichier_contenu, session, typing)
  })
  

  observeEvent(input$show_budget_modal, {
    showModal(modalDialog(
      title     = "🗺️ Mapping budgétaire",
      size      = "l",
      easyClose = TRUE,
      footer    = modalButton("Fermer"),
      mod_budget_mapping_ui("mapping1")
    ))
  })
  
  
  # MODAL BPSS
  observeEvent(input$toggle_bpss_ui, {
    showModal(modalDialog(title = "🛠️ Outil BPSS", size = "l", easyClose = TRUE,
                          footer = modalButton("Fermer"), mod_outil_bpss_ui("bpss1")))
  })
  
   # IMPORT JSON
  observeEvent(rv$imported_json, {
    req(rv$imported_json)
    showNotification("📂 JSON importé avec succès !", type = "message")
    msgs <- chat_history()
    msgs <- append(msgs, list(list(role = "assistant", content = "✅ J'ai reçu la structure JSON de votre classeur.")))
    chat_history(msgs)
    session$sendCustomMessage("scrollToBottom", list())
  })
  
  # ANALYSE DES LABELS
  observeEvent(rv$extracted_labels, {
    req(rv$extracted_labels)
    labs <- unique(unlist(rv$extracted_labels))
    analysis <- analyze_labels(labs)
    if (!is.list(analysis) || !all(c("axes","contexte_general") %in% names(analysis))) {
      showNotification("❌ Échec de l'analyse des labels.", type = "error")
      return()
    }
    rv$axes    <- analysis$axes
    rv$context <- analysis$contexte_general
    axes_info <- if (is.data.frame(rv$axes)) {
      paste0("- ", rv$axes$axe, ": ", rv$axes$description, collapse = "\n")
    } else {
      paste(sapply(rv$axes, function(x) paste0("- ", x$axe, ": ", x$description)), collapse = "\n")
    }
    recap <- paste0("✅ Axes détectés :\n", axes_info, "\n\nContexte : ", rv$context)
    msgs <- chat_history()
    msgs <- append(msgs, list(list(role = "assistant", content = recap)))
    chat_history(msgs)
    session$sendCustomMessage("scrollToBottom", list())
  })
  
  # TÉLÉCHARGEMENT & HISTORIQUE
  output$download_chat <- downloadHandler(
    filename = function() paste0("chat_", Sys.Date(), ".txt"),
    content = function(file) {
      msgs <- chat_history()
      writeLines(sapply(msgs, function(m) paste0(m$role, ": ", m$content)), con = file)
    }
  )
  
  # RENDU CHAT HISTORY AVEC QUICK REPLIES
  output$chat_history <- renderUI({
    msgs <- chat_history()
    if (length(msgs) == 0) return(NULL)
    rendered <- lapply(seq_along(msgs), function(i) {
      msg <- msgs[[i]]
      if (!is.null(msg$meta) && msg$meta == "fichier_content") return(NULL)
      # Type BPSS
      if (!is.null(msg$type) && msg$type == "bpss_prompt") {
        return(
          div(class = "chat-bubble-container bot-message-container",
              div(class = "chat-sender", "BudgiBot"),
              div(class = "chat-message bot-message", "Souhaitez-vous lancer l'outil BPSS ?"),
              div(class = "quick-replies",
                  actionButton("lancer_outil_bpss", "🛠️ Lancer l’outil Excel BPSS", class = "btn btn-success")
              )
          )
        )
      }
      if (msg$role == "assistant") {
        return(
          div(class = "chat-bubble-container bot-message-container",
              div(class = "chat-sender", "BudgiBot"),
              div(class = "chat-message bot-message", HTML(msg$content)),
              div(class = "quick-replies",
                  actionButton(paste0("btn_detail_", i), "Peux-tu détailler ?", onclick =
                                 "Shiny.setInputValue('user_input', 'Peux-tu détailler ?', {priority:'event'}); $('#send_btn').click();"),
                  actionButton(paste0("btn_example_", i), "Donne-moi un exemple", onclick =
                                 "Shiny.setInputValue('user_input', 'Donne-moi un exemple', {priority:'event'}); $('#send_btn').click();"),
                  actionButton(paste0("btn_resume_", i), "Résume", onclick =
                                 "Shiny.setInputValue('user_input', 'Résume', {priority:'event'}); $('#send_btn').click();"),
                  actionButton(paste0("btn_extract_budget_", i), "Extrait les données budgétaires", onclick =
                                 paste0("Shiny.setInputValue('extract_budget_under_bot_clicked', ", i, ", {priority:'event'});")
                  )
              )
          )
        )
      } else if (msg$role == "user") {
        return(
          div(class = "chat-bubble-container user-message-container",
              div(class = "chat-sender", "Vous"),
              div(class = "chat-message user-message", msg$content)
          )
        )
      }
      NULL
    })
    tagList(Filter(Negate(is.null), rendered))
  })
}
