library(shiny)
library(DT)
library(shinycssloaders)
library(logger)

# Options globales
options(shiny.maxRequestSize = 30 * 1024^2)

server <- function(input, output, session) {
  # -- Initialisation du logger --
  log_appender(appender_console)
  log_threshold(INFO)
  
  # -- RÉACTIFS GLOBAUX --
  chat_history <- reactiveVal(list())
  dernier_fichier_contenu <- reactiveVal(NULL)
  outil_process <- reactiveVal(NULL)
  bpss_prompt_active <- reactiveVal(FALSE)
  typing <- reactiveVal(FALSE)
  
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
  
  # -- MODULE : Traitement de fichiers (PDF, Word, Excel, JSON) --
  file_data <- callModule(
    mod_file_processor_server, "fileproc",
    input, session,
    chat_history, dernier_fichier_contenu, typing
  )
  
  # -- MODULE : Aide JSON --
  callModule(
    mod_json_helper_server, "json_helper",
    rv
  )
  
  # -- MODULE : Analyse des labels & détection d'axes --
  label_results <- callModule(
    mod_label_analysis_server, "labelana",
    labels       = rv$extracted_labels,
    chat_history = chat_history
  )
  
  # -- MODULE : Outil BPSS --
  callModule(
    mod_outil_bpss_server, "bpss1",
    bpss_prompt_active
  )
  
  # -- EXTRACTION BUDGÉTAIRE --
  budget_data <- eventReactive(label_results(), {
    withProgress(message = "Extraction budgétaire en cours...", {
      req(label_results()$axes, label_results()$contexte)
      df <- setupBudgetExtraction(
        input, output, session,
        rv, chat_history,
        dernier_fichier_contenu,
        donnees_extraites
      )
      validate(
        need(is.data.frame(df), "L'extraction n'a pas renvoyé de data.frame valide.")
      )
      log_info("Extraction budgétaire effectuée : {nrow(df)} lignes.")
      df
    })
  })
  
  # -- AFFICHAGE DU TABLEAU BUDGÉTAIRE --
  output$budget_table <- renderDT({
    shinycssloaders::withSpinner(
      datatable(
        budget_data(),
        options = list(scrollX = TRUE),
        editable = TRUE
      )
    )
  })
  
  observeEvent(input$budget_table_cell_edit, {
    info <- input$budget_table_cell_edit
    df   <- budget_data()
    df[info$row, info$col + 1] <- DT::coerceValue(
      info$value, df[info$row, info$col + 1]
    )
    donnees_extraites(df)
    log_debug("Cellule modifiée ligne={info$row}, col={info$col}")
  })
  
  observeEvent(input$save_budget_changes, {
    removeModal()
  })
  
  # -- GESTION DES MESSAGES UTILISATEUR --
  observeEvent(input$send_btn, {
    req(input$user_input)
    handle_user_input(
      input, session,
      chat_history, typing,
      bpss_prompt_active
    )
  })
  
  # -- GESTION DU CHARGEMENT DE FICHIER --
  observeEvent(input$file_input, {
    req(input$file_input)
    handle_file_message(
      input$file_input,
      chat_history,
      dernier_fichier_contenu,
      session,
      typing
    )
  })
  
  # -- AFFICHAGE DU CHAT --
  output$chat_history <- renderUI({
    msgs <- chat_history()
    if (length(msgs) == 0) return(NULL)
    lapply(seq_along(msgs), function(i) {
      msg <- msgs[[i]]
      # Ignorer le contenu brut de fichier
      if (!is.null(msg$meta) && msg$meta == "fichier_content") return(NULL)
      
      # Prompt BPSS
      if (!is.null(msg$type) && msg$type == "bpss_prompt") {
        return(
          div(class = "bot-message-container",
              div(class = "chat-sender", "BudgiBot"),
              div(class = "chat-message bot-message",
                  "Souhaitez-vous lancer l'outil BPSS ?"),
              div(class = "quick-replies",
                  actionButton(
                    "lancer_outil_bpss",
                    "🛠️ Lancer l’outil Excel BPSS",
                    class = "btn btn-success"
                  )
              )
          )
        )
      }
      
      # Messages assistant / utilisateur
      if (msg$role == "assistant") {
        div(class = "bot-message-container",
            div(class = "chat-sender", "BudgiBot"),
            div(class = "chat-message bot-message",
                HTML(msg$content)),
            div(class = "quick-replies",
                actionButton(paste0("btn_detail_", i), "Peux-tu détailler ?", onclick =
                               "Shiny.setInputValue('user_input', 'Peux-tu détailler ?'); $('#send_btn').click();"),
                actionButton(paste0("btn_example_", i), "Donne-moi un exemple", onclick =
                               "Shiny.setInputValue('user_input', 'Donne-moi un exemple'); $('#send_btn').click();"),
                actionButton(paste0("btn_resume_", i), "Résume", onclick =
                               "Shiny.setInputValue('user_input', 'Résume'); $('#send_btn').click();"),
                actionButton(paste0("btn_extract_budget_", i),
                             "Extrait les données budgétaires", onclick =
                               paste0("Shiny.setInputValue('extract_budget_under_bot_clicked', ", i, ", {priority: 'event'});")
                )
            )
        )
      } else if (msg$role == "user") {
        div(class = "user-message-container",
            div(class = "chat-sender", "Utilisateur"),
            div(class = "chat-message user-message", msg$content)
        )
      }
    }) |> tagList()
  })
  
  # -- MODULE BPSS UI & JS --
  observeEvent(input$toggle_bpss_ui, {
    showModal(
      modalDialog(
        title  = "🛠️ Outil BPSS",
        size   = "l",
        easyClose = TRUE,
        footer = modalButton("Fermer"),
        mod_outil_bpss_ui("bpss1")
      )
    )
  })
  observeEvent(input$lancer_outil_bpss, {
    session$sendCustomMessage(
      type    = "jsCode",
      message = list(code = "$('#toggle_bpss_ui').click();")
    )
  })
  
  # -- IMPORT JSON --
  observeEvent(rv$imported_json, {
    req(rv$imported_json)
    showNotification("📂 JSON importé avec succès !", type = "message")
    chat_history(append(chat_history(), list(
      list(role = "assistant", content =
             "✅ J'ai bien reçu un fichier JSON décrivant la structure de votre classeur Excel.")
    )))
    session$sendCustomMessage(type = 'scrollToBottom', message = list())
  })
  
  # -- ANALYSE LABELS --
  observeEvent(rv$extracted_labels, {
    req(rv$extracted_labels)
    log_debug("Labels extraits : {paste(rv$extracted_labels, collapse=', ')}")
    result <- analyze_labels(rv$extracted_labels)
    if (is.null(result) || !is.list(result) ||
        !all(c("axes", "contexte_general") %in% names(result))) {
      showNotification("❌ Échec de l'analyse des labels.", type = "error")
      return()
    }
    rv$axes    <- result$axes
    rv$context <- result$contexte_general
    # Message récapitulatif
    axes_info <- if (is.data.frame(rv$axes)) {
      paste0("- ", rv$axes$axe, ": ", rv$axes$description, collapse="\n")
    } else {
      paste0(sapply(rv$axes, function(x) paste0("- ", x$axe, ": ", x$description)), collapse="\n")
    }
    recap <- paste0(
      "✅ Axes détectés :\n", axes_info,
      "\n\nContexte :\n", rv$context,
      "\n\nSouhaitez-vous analyser un texte budgétaire ?"
    )
    chat_history(append(chat_history(), list(list(role="assistant", content=recap))))
    session$sendCustomMessage(type = 'scrollToBottom', message = list())
  })
  
  # -- TÉLÉCHARGER L'HISTORIQUE --
  output$download_chat <- downloadHandler(
    filename = function() paste0("chat_", Sys.Date(), ".txt"),
    content  = function(file) {
      msgs <- chat_history()
      writeLines(
        sapply(msgs, function(m) paste0(m$role, ": ", m$content)),
        con = file
      )
      log_info("Historique exporté vers {file}")
    }
  )
}

shinyApp(ui = ui, server = server)
