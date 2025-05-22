library(shiny)
library(shinycssloaders)
library(DT)

source("~/work/lereuf/chatbot/modules/logique_excel_outil_bpss.R")
#NB : La détection des programmes est faite via un sous-string dans le nom. Il faut toujours indiquer 3 chiffres dans le programems. 

read_xlsx_with_recovery <- function(file_path, sheet = NULL) {
  tryCatch({
    readxl::read_excel(file_path, sheet = sheet)
  }, error = function(e) {
    warning(paste("Erreur de lecture du fichier Excel :", e$message))
    data.frame()
  })
}


mod_outil_bpss_ui <- function(id) {
  ns <- NS(id)
  tagList(
    fluidRow(
      # Colonne de gauche (inputs et actions)
      column(
        width = 4,
        wellPanel(
          h4("📊 Paramètres & Fichiers"),
          numericInput( ns("annee"),           "Année",          value = 2025, min = 2000, max = 2100),
          textInput(    ns("code_ministere"),  "Code Ministère", value = "38"),
          textInput(    ns("code_programme"),  "Code Programme", value = "150"),
          fileInput(    ns("ppes_file"),       "PP‑E‑S (.xlsx)", accept = ".xlsx"),
          fileInput(    ns("dpp18_file"),      "DPP 18 (.xlsx)", accept = ".xlsx"),
          fileInput(    ns("bud45_file"),      "BUD 45 (.xlsx)", accept = ".xlsx"),
          #fileInput(    ns("final_file"),      "Template Final", accept = ".xlsx"),
          actionButton( ns("process_button"), "✅ Générer",      class = "btn-primary w-100"),
          br(), br(),
          #downloadButton(ns("download_final"),"⬇️ Télécharger",   class = "btn-success w-100"),
          br(), br(),
          textOutput(   ns("status"))
        )
      ),
      # Colonne de droite (aperçus empilés)
      column(
        width = 8,
        # 1️⃣ PP‑E‑S
        h4("Aperçu PP‑E‑S"),
        shinycssloaders::withSpinner(DT::DTOutput(ns("table_ppes")), type = 4),
        hr(),
        # 2️⃣ DPP 18
        h4("Aperçu DPP 18"),
        shinycssloaders::withSpinner(DT::DTOutput(ns("table_dpp18")), type = 4),
        hr(),
        # 3️⃣ BUD 45
        h4("Aperçu BUD 45"),
        shinycssloaders::withSpinner(DT::DTOutput(ns("table_bud45")), type = 4)
      )
    )
  )
}



mod_outil_bpss_server <- function(id, rv) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    #final_file_path <- reactiveVal(NULL)
    

    observeEvent(input$process_button, {
      req(input$ppes_file, input$dpp18_file, input$bud45_file, rv$fichier_excel)
      shinyjs::disable("process_button")
      output$status <- renderText("⚙️ Démarrage du traitement…")
      
      withProgress(message = "Traitement en cours", value = 0, {
        incProgress(0.1)
        tryCatch({
          #Récupération des paramètres
          ppes_path   <- input$ppes_file$datapath
          dpp18_path  <- input$dpp18_file$datapath
          bud45_path  <- input$bud45_file$datapath
          code_min    <- input$code_ministere
          code_prog   <- input$code_programme
          wb          <- rv$fichier_excel
          
          incProgress(0.2)
          
          if (is.null(wb) || !inherits(wb, "wbWorkbook")) {
            showNotification("❌ Le workbook Excel n'est pas valide.", type = "error")
            output$status <- renderText("❌ Workbook Excel manquant ou invalide.")
            return()
          }
          
          #Lecture des feuilles internes
          feuilles <- list(
            pp_categ = paste0("MIN_", code_min, "_DETAIL_Prog_PP_CATEG"),
            entrants = paste0("MIN_", code_min, "_DETAIL_Prog_Entrants"),
            sortants = paste0("MIN_", code_min, "_DETAIL_Prog_Sortants")
          )
          
          df_pp_categ <- read_xlsx_with_recovery(ppes_path, feuilles$pp_categ)
          df_entrants <- read_xlsx_with_recovery(ppes_path, feuilles$entrants)
          df_sortants <- read_xlsx_with_recovery(ppes_path, feuilles$sortants)
          df_dpp18    <- read_xlsx_with_recovery(dpp18_path)
          df_bud45    <- read_xlsx_with_recovery(bud45_path)
          incProgress(0.2)
          
          wb <- load_ppes_data(wb, df_pp_categ, df_entrants, df_sortants, code_prog)
          incProgress(0.1)
          wb <- load_inf_dpp18(wb, df_dpp18, code_prog)
          incProgress(0.1)
          wb <- load_inf_bud45(wb, df_bud45, code_prog)
          incProgress(0.1)
          
          rv$fichier_excel <- wb
          rv$excel_updated <- Sys.time()
          incProgress(0.2)
          
          output$status <- renderText("✅ Terminé, cliquez pour télécharger.")
          showNotification("📄 Fichier complété avec succès !", type = "message")
          
        }, error = function(e) {
          showNotification(paste0("❌ Erreur : ", e$message), type = "error")
          output$status <- renderText("❌ Une erreur est survenue.")
        })
      })
      
      shinyjs::enable("process_button")
    })
    
    # — Previews avec DT
    output$table_ppes <- DT::renderDT({
      req(input$ppes_file)
      read_xlsx_with_recovery(input$ppes_file$datapath, 
                sheet = paste0("MIN_", input$code_ministere, "_DETAIL_Prog_PP_CATEG")) %>%
        dplyr::filter(substr(as.character(nom_prog),1,3)==substr(input$code_programme,1,3))
    }, options=list(pageLength=3, scrollX=TRUE))
    
    output$table_dpp18 <- DT::renderDT({
      req(input$dpp18_file)
      read_xlsx_with_recovery(input$dpp18_file$datapath) %>%
        dplyr::filter(str_detect(.[[1]], fixed(substr(input$code_programme,1,3))))
    }, options=list(pageLength=3, scrollX=TRUE))
    
    output$table_bud45 <- DT::renderDT({
      req(input$bud45_file)
      read_xlsx_with_recovery(input$bud45_file$datapath) %>%
        dplyr::filter(str_detect(.[[2]], fixed(substr(input$code_programme,1,3))))
    }, options=list(pageLength=3, scrollX=TRUE))
  })
}

