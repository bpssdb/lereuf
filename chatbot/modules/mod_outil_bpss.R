library(shiny)
library(shinycssloaders)
library(DT)

#NB : La détection des programmes est faite via un sous-string dans le nom. Il faut toujours indiquer 3 chiffres dans le programems. 


insert_data_to_sheet <- function(wb, sheet_name, data, startCol = 1, startRow = 1) {
  if (!inherits(wb, "wbWorkbook")) stop("❌ wb n'est pas un workbook openxlsx2.")
  if (!sheet_name %in% openxlsx2::wb_get_sheet_names(wb)) {
    wb <- openxlsx2::wb_add_worksheet(wb, sheet_name)

  }
  wb <- openxlsx2::wb_add_data(wb, sheet = sheet_name, x = data, startCol = startCol, startRow = startRow)
}

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
      # 1) Validation
      req(input$ppes_file, input$dpp18_file, input$bud45_file, rv$fichier_excel)
      shinyjs::disable("process_button")                    # désactive bouton
      output$status <- renderText("⚙️ Démarrage du traitement…")
      
      # 2) withProgress pour donner du feedback
      withProgress(message = "Traitement en cours", value = 0, {
        incProgress(0.1)
        tryCatch({
          # ⚙️ Paramètres
          ppes_path <- input$ppes_file$datapath
          dpp18_path <- input$dpp18_file$datapath
          bud45_path <- input$bud45_file$datapath
          
          code_ministere <- input$code_ministere
          code_programme <- input$code_programme
          
          wb <- rv$fichier_excel
          if (is.null(wb) || !inherits(wb, "wbWorkbook")) {
            showNotification("❌ Le workbook Excel n'est pas valide.", type = "error")
            output$status <- renderText("❌ Workbook Excel manquant ou invalide.")
            return()
          }
          
          # Exemple pour lire les feuilles internes
          feuilles <- list(
            pp_categ = paste0("MIN_", code_ministere, "_DETAIL_Prog_PP_CATEG"),
            entrants = paste0("MIN_", code_ministere, "_DETAIL_Prog_Entrants"),
            sortants = paste0("MIN_", code_ministere, "_DETAIL_Prog_Sortants")
          )
          
          df_pp_categ <- read_xlsx_with_recovery(ppes_path, feuilles$pp_categ)
          df_entrants <- read_xlsx_with_recovery(ppes_path, feuilles$entrants)
          df_sortants <- read_xlsx_with_recovery(ppes_path, feuilles$sortants)
          print("feuilles lues")
          
          # 🎯 Filtrage
          filtered_pp_categ <- df_pp_categ %>% 
            dplyr::filter(substr(as.character(nom_prog),1,3)==substr(input$code_programme,1,3))
          print("premier filtrage")
          filtered_entrants <- df_entrants %>% 
            filter(substr(as.character(nom_prog), 1, 3) == code_programme)
          print("deuxieme filtrage")
          filtered_sortants <- df_sortants %>% 
            filter(substr(as.character(nom_prog), 1, 3) == code_programme) 
          print("troisieme filtrage")
          # 📊 Ajout dans le workbook
          insert_data_to_sheet(wb, "Données PP-E-S", filtered_pp_categ,  startCol = 3, startRow = 7)
          insert_data_to_sheet(wb, "Données PP-E-S", filtered_entrants, startCol = 3, startRow = 113)
          insert_data_to_sheet(wb, "Données PP-E-S", filtered_sortants, startCol = 3, startRow = 213)
          print("Ajout dans le wb")
          
          df_dpp18 <- read_xlsx_with_recovery(dpp18_path) %>% 
            filter(str_detect(.[[1]], fixed(code_programme)))
          df_bud45 <- read_xlsx_with_recovery(bud45_path) %>% 
            filter(str_detect(.[[1]], fixed(code_programme)))
          
          insert_data_to_sheet(wb, "INF DPP 18", df_dpp18, startCol = 2, startRow = 6)
          insert_data_to_sheet(wb, "INF BUD 45", df_bud45, startCol = 2, startRow = 6)
          
          # 👤 Indicié dans Accueil
          df_pp_indicie <- filtered_pp_categ %>% filter(marqueur_masse_indiciaire == "Indicié")
          wb <- wb_add_data(wb, "Accueil", "", dims = paste0("B43:B", 43 + nrow(df_pp_indicie)))
          wb <- wb_add_data(wb, "Accueil", "", dims = paste0("C43:C", 43 + nrow(df_pp_indicie)))
          wb <- wb_add_data(wb, "Accueil", df_pp_indicie[[2]], start_col = 2, start_row = 43)
          wb <- wb_add_data(wb, "Accueil", df_pp_indicie[[3]], start_col = 3, start_row = 43)
          
          # 📈 Calculs Excel
          data_socle <- wb_read(wb, "I - Socle exécution n-1", col_names = FALSE)
          data_hyp   <- wb_read(wb, "III - Hyp. salariales", col_names = FALSE)
          
          data_socle[, 3:5] <- lapply(data_socle[, 3:5], as.numeric)
          data_hyp[, 5]     <- as.numeric(data_hyp[, 5])
          data_hyp[, 7:10]  <- lapply(data_hyp[, 7:10], as.numeric)
          
          val_C67 <- sum(data_socle[c(34, 44, 46, 49), 3], na.rm = TRUE)
          val_C68 <- sum(data_socle[c(35, 45, 47), 3], na.rm = TRUE)
          val_D68 <- sum(data_socle[c(35, 45, 47), 4], na.rm = TRUE)
          val_E68 <- sum(data_socle[c(35, 45, 47), 5], na.rm = TRUE)
          
          val_E40     <- sum(data_hyp[c(113, 114, 115, 116), 5], na.rm = TRUE)
          vals_GHIJ40 <- sapply(7:10, function(col) sum(data_hyp[c(113, 114, 115, 116), col], na.rm = TRUE))
          
          wb <- wb_add_data(wb, "I - Socle exécution n-1", val_C67, dims = "C67")
          wb <- wb_add_data(wb, "I - Socle exécution n-1", val_C68, dims = "C68")
          wb <- wb_add_data(wb, "I - Socle exécution n-1", val_D68, dims = "D68")
          wb <- wb_add_data(wb, "I - Socle exécution n-1", val_E68, dims = "E68")
          
          wb <- wb_add_data(wb, "VI - Facteurs d'évolution MS", val_E40, dims = "E40")
          lapply(seq_along(vals_GHIJ40), function(i) {
            wb <<- wb_add_data(wb, "VI - Facteurs d'évolution MS", vals_GHIJ40[i], dims = paste0(LETTERS[7 + i - 1], "40"))
          })
          
          rv$fichier_excel <- wb
          rv$excel_updated <- Sys.time()  # force un signal de changement

          output$status <- renderText("✅ Terminé, cliquez pour télécharger.")
          showNotification("📄 Fichier complété avec succès !", type = "message")
          
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
    
    # — Download final
    #$download_final <- downloadHandler(
    #  filename = function() paste0("Budget_", input$code_ministere, "_", Sys.Date(), ".xlsx"),
    #  content  = function(file) file.copy(final_file_path(), file)
    #)
  })
}

