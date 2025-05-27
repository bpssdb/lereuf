# modules/mod_mesures_cat_combo_v3.R
# Module Excel refactorisé avec parseur optimisé v3

library(shiny)
library(readxl)
library(reactable)
library(openxlsx2)
library(rhandsontable)
library(shinycssloaders)
library(DT)

# ============================
# MODULE UI (inchangé)
# ============================
mod_mesures_cat_ui <- function(id) {
  ns <- NS(id)
  tagList(
    # CSS pour supprimer marges par défaut
    tags$style(HTML(paste0(
      "#", ns("upload_file"), " {margin-bottom:0!important; padding-bottom:0!important;}
",
      "#", ns("selected_sheet"), " {margin-bottom:0!important; padding-bottom:0!important;}
",
      "#", ns("table_container"), " {margin-top:0!important; padding-top:0!important;}
",
      "#", ns("buttons_container"), " {margin-top:0!important; padding-top:0!important;}"
    ))),
    # 1️⃣ Upload + Sélecteur côte à côte
    fluidRow(style = "margin-bottom:0; padding-bottom:0;",
             column(6,
                    fileInput(ns("upload_file"),
                              "📂 Charger un fichier Excel (.xlsx)",
                              accept = ".xlsx",
                              buttonLabel = "Parcourir…",
                              placeholder = "Aucun fichier sélectionné",
                              width = "100%"
                    )
             ),
             column(6,
                    uiOutput(ns("sheet_selector"))
             )
    ),
    # 2️⃣ Tableau conditionnel
    uiOutput(ns("table_container")),
    # 3️⃣ Boutons conditionnels
    uiOutput(ns("buttons_container"))
  )
}

# ============================
# MODULE SERVER (mise à jour v3)
# ============================

mod_mesures_cat_server <- function(id, rv, on_analysis_summary = NULL) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    
    # === Réactifs internes ===
    rv_path        <- reactiveVal(NULL)
    rv_path_script <- reactiveVal(NULL)
    rv_excel       <- reactiveVal(NULL)
    rv_sheets      <- reactiveVal(NULL)
    rv_table       <- reactiveVal(NULL)
    rv_selected    <- reactiveVal(NULL)
    
    # === Chargement initial du fichier Excel AVEC PARSEUR V3 ===
    observeEvent(input$upload_file, {
      req(input$upload_file)
      path <- input$upload_file$datapath
      ext  <- tools::file_ext(input$upload_file$name)
      
      if (ext != "xlsx") {
        showNotification("❌ Veuillez charger un fichier Excel (.xlsx)", type = "error")
        return()
      }
      
      showNotification("🔄 Chargement avec le parseur optimisé v3...", type = "message", duration = 3)
      
      wb <- tryCatch(openxlsx2::wb_load(path), error = function(e) NULL)
      if (inherits(wb, "wbWorkbook")) {
        rv_path(path)
        rv_excel(wb)
        rv$fichier_excel <- wb
        
        sheets <- tryCatch(openxlsx2::wb_get_sheet_names(wb), error = function(e) NULL)
        rv_sheets(sheets)
        updateSelectInput(session, ns("selected_sheet"), choices = sheets, selected = sheets[1])
        
        # === NOUVEAU: Parsing avec le parseur v3 ===
        withProgress(message = "Analyse avancée des formules...", value = 0, {
          tryCatch({
            incProgress(0.2, detail = "Initialisation du parseur v3")
            
            # Source le parseur v3 s'il n'est pas déjà chargé
            if (!exists("parse_excel_formulas_v3")) {
              source("~/work/lereuf/chatbot/modules/parseur_excel/parseur_excel_v3.R")
            }
            
            incProgress(0.4, detail = "Extraction et conversion des formules")
            
            # Utilisation du nouveau parseur optimisé
            parsing_result <- parse_excel_formulas_v3(rv_path(), emit_script = TRUE)
            
            incProgress(0.8, detail = "Génération du script optimisé")
            
            if (!is.null(parsing_result$script_file)) {
              rv_path_script(parsing_result$script_file)
              
              # Stocker les résultats du parsing dans rv pour utilisation ultérieure
              rv$parsing_result <- parsing_result
              
              # Message de succès détaillé
              success_msg <- sprintf(
                "✅ Parsing v3 terminé: %d formules analysées en %.2fs (%.1f%% succès)", 
                parsing_result$statistics$total,
                parsing_result$processing_time,
                parsing_result$statistics$success_rate
              )
              showNotification(success_msg, type = "message", duration = 5)
              
              # Log des performances
              message("=== PERFORMANCES PARSEUR V3 ===")
              message(sprintf("Fichier: %s", basename(path)))
              message(sprintf("Formules: %d total, %d succès, %d erreurs", 
                              parsing_result$statistics$total,
                              parsing_result$statistics$success,
                              parsing_result$statistics$errors))
              message(sprintf("Temps: %.2f secondes", parsing_result$processing_time))
              message(sprintf("Config: chunks=%d, workers=%d", 
                              parsing_result$config_used$chunk_size,
                              parsing_result$config_used$parallel_workers))
              
            } else {
              showNotification("ℹ️ Aucune formule trouvée dans le fichier", type = "message")
            }
            
            incProgress(1, detail = "Terminé")
            
          }, error = function(e) {
            error_msg <- paste("❌ Erreur parsing v3:", e$message)
            showNotification(error_msg, type = "error", duration = 10)
            message("ERREUR PARSEUR V3: ", e$message)
          })
        })
        
      } else {
        showNotification("❌ Échec du chargement du fichier.", type = "error")
      }
    })
    
    # === Rendu du sélecteur de feuilles ===
    output$sheet_selector <- renderUI({
      req(rv_sheets())
      selectInput(ns("selected_sheet"), "🗂️ Choisir une feuille", choices = rv_sheets(), width = "100%")
    })
    
    # === Lecture de la feuille sélectionnée ===
    observeEvent(input$selected_sheet, {
      req(rv_excel(), input$selected_sheet)
      df <- tryCatch(
        openxlsx2::wb_to_df(rv_excel(), sheet = input$selected_sheet),
        error = function(e) {
          showNotification(paste("❌ Erreur lecture feuille :", e$message), type = "error")
          NULL
        }
      )
      req(df)
      rv_selected(input$selected_sheet)
      rv_table(as.data.frame(df))
      rv$excel_data  <- df
      rv$excel_sheet <- input$selected_sheet
    })
    
    # === Recharge forcée si excel modifié ailleurs ===
    observeEvent(rv$excel_updated, {
      wb <- rv$fichier_excel
      req(inherits(wb, "wbWorkbook"))
      
      rv_excel(wb)
      feuilles <- tryCatch(openxlsx2::wb_get_sheet_names(wb), error = function(e) NULL)
      rv_sheets(feuilles)
      
      sheet <- rv_selected()
      if (!is.null(sheet) && sheet %in% feuilles) {
        df <- tryCatch(openxlsx2::wb_to_df(wb, sheet = sheet), error = function(e) NULL)
        if (!is.null(df)) {
          rv_table(as.data.frame(df))
        }
      }
    })
    
    # === Table UI ===
    output$table_container <- renderUI({
      req(rv_table())
      fluidRow(
        column(12,
               shinycssloaders::withSpinner(
                 reactableOutput(ns("reactable_table")),
                 type = 4, color = "#0055A4"
               )
        )
      )
    })
    
    output$reactable_table <- renderReactable({
      df <- rv_table()
      if (is.null(df) || !is.data.frame(df)) {
        return(reactable(data.frame(Message = "📭 Aucune donnée à afficher"), bordered = TRUE))
      }
      
      if (is.null(names(df)) || any(is.na(names(df))) || any(names(df) == "")) {
        names(df) <- paste0("Colonne_", seq_len(ncol(df)))
      }      
      
      reactable(
        df,
        searchable = TRUE, resizable = TRUE, highlight = TRUE, bordered = TRUE,
        striped = TRUE, pagination = FALSE,
        defaultColDef = colDef(minWidth = 100, style = list(whiteSpace = "pre-wrap")),
        style = list(maxHeight = "70vh", overflowY = "auto")
      )
    })
    
    # === Boutons AVEC INFORMATIONS PARSEUR V3 ===
    output$buttons_container <- renderUI({
      req(rv_table())
      
      # Texte dynamique pour le bouton selon les résultats du parsing
      formula_btn_text <- if(!is.null(rv$parsing_result)) {
        sprintf("🖋️ Appliquer formules (%d)", rv$parsing_result$statistics$success)
      } else {
        "🖋️ Appliquer les formules"
      }
      
      fluidRow(
        column(12,
               div(class = "d-flex",
                   actionButton(ns("open_full_editor"), "🖋️ Modifier la feuille", class = "btn btn-secondary mr-2"),
                   actionButton(ns("apply_formulas"), formula_btn_text, class = "btn btn-secondary mr-2"),
                   actionButton(ns("show_parsing_stats"), "📊 Stats parsing v3", class = "btn btn-info mr-2"),
                   downloadButton(ns("download_table"), "💾 Exporter le tableau", class = "btn btn-success")
               )
        )
      )
    })
    
    # === NOUVEAU: Modal d'affichage des statistiques de parsing ===
    observeEvent(input$show_parsing_stats, {
      req(rv$parsing_result)
      
      stats <- rv$parsing_result$statistics
      config <- rv$parsing_result$config_used
      
      showModal(modalDialog(
        title = "📊 Statistiques du Parseur Excel v3",
        size = "l",
        easyClose = TRUE,
        tagList(
          fluidRow(
            column(6,
                   h4("🎯 Résultats"),
                   tags$table(class = "table table-striped",
                              tags$tr(tags$td("Total formules:"), tags$td(tags$strong(stats$total))),
                              tags$tr(tags$td("Conversions réussies:"), tags$td(tags$strong(stats$success))),
                              tags$tr(tags$td("Erreurs:"), tags$td(tags$strong(stats$errors))),
                              tags$tr(tags$td("Taux de succès:"), tags$td(tags$strong(paste0(stats$success_rate, "%")))),
                              tags$tr(tags$td("Temps de traitement:"), tags$td(tags$strong(sprintf("%.2f sec", rv$parsing_result$processing_time))))
                   )
            ),
            column(6,
                   h4("⚙️ Configuration"),
                   tags$table(class = "table table-striped",
                              tags$tr(tags$td("Taille des chunks:"), tags$td(config$chunk_size)),
                              tags$tr(tags$td("Workers parallèles:"), tags$td(config$parallel_workers)),
                              tags$tr(tags$td("Cache activé:"), tags$td(if(config$cache_enabled) "✅ Oui" else "❌ Non")),
                              tags$tr(tags$td("Limite mémoire:"), tags$td(paste(config$max_memory_mb, "MB"))),
                              tags$tr(tags$td("Progression:"), tags$td(if(config$progress_enabled) "✅ Oui" else "❌ Non"))
                   )
            )
          ),
          if (!is.null(rv$parsing_result$errors) && nrow(rv$parsing_result$errors) > 0) {
            tagList(
              hr(),
              h4("⚠️ Formules non converties"),
              DT::DTOutput(ns("errors_table"))
            )
          } else {
            tagList(
              hr(),
              div(class = "alert alert-success", 
                  icon("check-circle"), " Toutes les formules ont été converties avec succès !")
            )
          }
        ),
        footer = modalButton("Fermer")
      ))
      
      # Affichage du tableau d'erreurs si présent
      if (!is.null(rv$parsing_result$errors) && nrow(rv$parsing_result$errors) > 0) {
        output$errors_table <- DT::renderDT({
          rv$parsing_result$errors %>%
            select(sheet, address, formula) %>%
            mutate(formula = substr(formula, 1, 100)) %>%  # Tronquer les formules longues
            DT::datatable(
              options = list(
                pageLength = 10, 
                scrollX = TRUE,
                columnDefs = list(list(width = "50%", targets = 2))  # Largeur colonne formule
              ),
              colnames = c("Feuille", "Cellule", "Formule")
            )
        })
      }
    })
    
    # === Application des formules avec le script v3 ===
    observeEvent(input$apply_formulas, {
      script_path <- rv_path_script()
      
      if (!is.null(script_path) && file.exists(script_path)) {
        
        withProgress(message = "Application des formules optimisées...", value = 0, {
          tryCatch({
            incProgress(0.2, detail = "Chargement du script v3")
            
            # Source le script généré
            source(script_path)
            
            incProgress(0.5, detail = "Exécution des calculs")
            
            # Exécution avec gestion d'erreur
            script(rv, rv_path)
            
            incProgress(1, detail = "Mise à jour terminée")
            
            # Message de succès avec détails
            if (!is.null(rv$parsing_result)) {
              success_msg <- sprintf(
                "✅ %d formules appliquées avec succès", 
                rv$parsing_result$statistics$success
              )
              showNotification(success_msg, type = "message", duration = 5)
            }
            
          }, error = function(e) {
            error_msg <- paste("❌ Erreur application formules:", e$message)
            showNotification(error_msg, type = "error", duration = 10)
            message("ERREUR APPLICATION FORMULES: ", e$message)
          })
        })
        
      } else {
        showNotification("❌ Script de formules introuvable. Rechargez le fichier.", type = "error")
      }
    })
    
    # === Éditeur plein écran ===
    observeEvent(input$open_full_editor, {
      req(rv_table())
      showModal(modalDialog(
        title     = "🖋️ Édition plein écran",
        size      = "l",
        easyClose = TRUE,
        footer    = tagList(
          modalButton("❌ Fermer"),
          actionButton(ns("save_edits"), "💾 Enregistrer", class = "btn btn-primary")
        ),
        rHandsontableOutput(ns("hot_table"), height = "70vh"),
        tags$script(HTML("$('.modal-dialog').css('width','95vw')"))
      ))
    })
    
    output$hot_table <- renderRHandsontable({
      df <- rv_table()
      req(df)
      
      if (!is.data.frame(df)) {
        showNotification("❌ Le tableau n'est pas un data.frame.", type = "error")
        return(NULL)
      }
      
      nc <- suppressWarnings(ncol(df))
      if (is.null(nc) || is.na(nc) || nc < 1) {
        showNotification("❌ Le tableau est vide ou mal formé.", type = "error")
        return(NULL)
      }
      
      # Corriger les noms de colonnes si besoin
      if (is.null(names(df)) || any(is.na(names(df))) || any(names(df) == "")) {
        names(df) <- paste0("Colonne_", seq_len(nc))
      }
      
      rhandsontable(df, useTypes = TRUE, stretchH = "all") %>%
        hot_cols(colWidths = rep(120, nc))
    })
    
    observeEvent(input$save_edits, {
      req(input$hot_table)
      rv_table(hot_to_r(input$hot_table))
      removeModal()
    })
    
    # === Téléchargement avec nom v3 ===
    output$download_table <- downloadHandler(
      filename = function() {
        paste0("sortie_budgibot_v3_", Sys.Date(), ".xlsx")
      },
      content = function(file) {
        wb <- rv$fichier_excel
        req(inherits(wb, "wbWorkbook"))
        
        openxlsx2::wb_save(wb, file = file, overwrite = TRUE)
        
        # Message avec infos parsing si disponible
        if (!is.null(rv$parsing_result)) {
          success_msg <- sprintf("✅ Export terminé - %d formules intégrées", 
                                 rv$parsing_result$statistics$success)
          showNotification(success_msg, type = "message")
        } else {
          showNotification("✅ Export Excel terminé", type = "message")
        }
      }
    )
    
  })
}

message("✅ Module mod_mesures_cat v3 chargé avec parseur optimisé")