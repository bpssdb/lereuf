library(shiny)
library(shinycssloaders)
library(DT)

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
          textInput(    ns("code_programme"),  "Code Programme", value = "123"),
          fileInput(    ns("ppes_file"),       "PP‑E‑S (.xlsx)", accept = ".xlsx"),
          fileInput(    ns("dpp18_file"),      "DPP 18 (.xlsx)", accept = ".xlsx"),
          fileInput(    ns("bud45_file"),      "BUD 45 (.xlsx)", accept = ".xlsx"),
          fileInput(    ns("final_file"),      "Template Final", accept = ".xlsx"),
          actionButton( ns("process_button"), "✅ Générer",      class = "btn-primary w-100"),
          br(), br(),
          downloadButton(ns("download_final"),"⬇️ Télécharger",   class = "btn-success w-100"),
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



mod_outil_bpss_server <- function(id) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    final_file_path <- reactiveVal(NULL)
    
    # utilitaire de lecture robuste
    read_safe <- function(path, sheet=NULL) {
      tryCatch(readxl::read_excel(path, sheet=sheet), error = function(e) data.frame())
    }
    
    observeEvent(input$process_button, {
      # 1) Validation
      req(input$ppes_file, input$dpp18_file, input$bud45_file, input$final_file)
      shinyjs::disable("process_button")                    # désactive bouton
      output$status <- renderText("⚙️ Démarrage du traitement…")
      
      # 2) withProgress pour donner du feedback
      withProgress(message = "Traitement en cours", value = 0, {
        incProgress(0.1)
        tryCatch({
          # chemins
          ppes_path <- input$ppes_file$datapath
          dpp18_path <- input$dpp18_file$datapath
          bud45_path <- input$bud45_file$datapath
          final_template <- input$final_file$datapath
          incProgress(0.1, detail="Lecture des fichiers source…")
          
          # lecture PP‑E‑S
          sheets <- paste0("MIN_", input$code_ministere, "_DETAIL_Prog_",
                           c("PP_CATEG","Entrants","Sortants"))
          dfs <- lapply(sheets, read_safe, path=ppes_path)
          names(dfs) <- c("categ","entrants","sortants")
          incProgress(0.2, detail="Filtrage et formattage…")
          
          # filtrage programme
          prog3 <- substr(as.character(input$code_programme),1,3)
          dfs <- lapply(dfs, function(df) {
            df[df$nom_prog %>% as.character() %>% substr(1,3) == prog3, ]
          })
          
          # chargement du template
          wb <- openxlsx2::wb_load(final_template)
          incProgress(0.2, detail="Insertion dans le classeur…")
          
          # insertion automatisée
          openxlsx2::wb_add_data(wb, "Données PP‑E‑S",  dfs$categ,    startCol=3, startRow=7)
          openxlsx2::wb_add_data(wb, "Données PP‑E‑S",  dfs$entrants, startCol=3, startRow=113)
          openxlsx2::wb_add_data(wb, "Données PP‑E‑S",  dfs$sortants, startCol=3, startRow=213)
          incProgress(0.2)
          
          # DPP18 / BUD45
          df18  <- read_safe(dpp18_path) %>% dplyr::filter(str_detect(.[[1]], fixed(prog3)))
          df45  <- read_safe(bud45_path) %>% dplyr::filter(str_detect(.[[1]], fixed(prog3)))
          openxlsx2::wb_add_data(wb, "INF DPP 18", df18, startCol=2, startRow=6)
          openxlsx2::wb_add_data(wb, "INF BUD 45", df45, startCol=2, startRow=6)
          incProgress(0.2)
          
          # calculs et indices… (pareil que toi)
          # …
          incProgress(0.1, detail="Finalisation…")
          
          # on sauve et on stocke
          outf <- tempfile(fileext=".xlsx")
          openxlsx2::wb_save(wb, outf)
          final_file_path(outf)
          
          output$status <- renderText("✅ Terminé, cliquez pour télécharger.")
        }, error = function(e) {
          showNotification(paste0("❌ Erreur : ", e$message), type="error")
          output$status <- renderText("❌ Une erreur est survenue.")
        })
      })
      
      shinyjs::enable("process_button")
    })
    
    # — Previews avec DT
    output$table_ppes <- DT::renderDT({
      req(input$ppes_file)
      read_safe(input$ppes_file$datapath, 
                sheet = paste0("MIN_", input$code_ministere, "_DETAIL_Prog_PP_CATEG")) %>%
        dplyr::filter(substr(as.character(nom_prog),1,3)==substr(input$code_programme,1,3))
    }, options=list(pageLength=3, scrollX=TRUE))
    
    output$table_dpp18 <- DT::renderDT({
      req(input$dpp18_file)
      read_safe(input$dpp18_file$datapath) %>%
        dplyr::filter(str_detect(.[[1]], fixed(substr(input$code_programme,1,3))))
    }, options=list(pageLength=3, scrollX=TRUE))
    
    output$table_bud45 <- DT::renderDT({
      req(input$bud45_file)
      read_safe(input$bud45_file$datapath) %>%
        dplyr::filter(str_detect(.[[1]], fixed(substr(input$code_programme,1,3))))
    }, options=list(pageLength=3, scrollX=TRUE))
    
    # — Download final
    output$download_final <- downloadHandler(
      filename = function() paste0("Budget_", input$code_ministere, "_", Sys.Date(), ".xlsx"),
      content  = function(file) file.copy(final_file_path(), file)
    )
  })
}

