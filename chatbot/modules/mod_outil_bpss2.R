library(stringr)


mod_outil_bpss_ui <- function(id) {
  ns <- NS(id)
  
  fluidPage(
    titlePanel("Application Budgétaire - Chargement et Calculs Automatisés"),
    
    sidebarLayout(
      sidebarPanel(
        h4("Étape 1 - Paramètres généraux"),
        textInput(ns("annee"), "Année :", value = "2025", placeholder = "e.g., 2025"),
        textInput(ns("code_ministere"), "Code Ministère :", value = "38", placeholder = "e.g., 38"),
        textInput(ns("code_programme"), "Code Programme (3 chiffres) :", value = "123", placeholder = "e.g., 123"),
        
        hr(),
        h4("Étape 2 - Déposez les fichiers sources"),
        fileInput(ns("ppes_file"), "Fichier PP-E-S", accept = ".xlsx"),
        fileInput(ns("dpp18_file"), "Fichier DPP 18", accept = ".xlsx"),
        fileInput(ns("bud45_file"), "Fichier BUD 45", accept = ".xlsx"),
        
        hr(),
        h4("Étape 3 - Déposez le fichier final (vierge)"),
        fileInput(ns("final_file"), "Fichier Final à Compléter", accept = ".xlsx"),
        actionButton(ns("process_button"), "Générer le fichier complété", class = "btn-primary"),
        downloadButton(ns("download_final"), "Télécharger le fichier final", class = "btn-success"),
        hr(),
        textOutput(ns("status"))
      ),
      
      mainPanel(
        h4("Aperçu - PP-E-S"),
        dataTableOutput(ns("table_ppes")),
        h4("Aperçu - DPP 18"),
        dataTableOutput(ns("table_dpp18")),
        h4("Aperçu - BUD 45"),
        dataTableOutput(ns("table_bud45"))
      )
    )
  )
}


mod_outil_bpss_server <- function(id) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    
    final_file_path <- reactiveVal(NULL)
    
    read_xlsx_with_recovery <- function(file_path, sheet = NULL) {
      tryCatch({
        readxl::read_excel(file_path, sheet = sheet)
      }, error = function(e) {
        warning("Erreur de lecture du fichier Excel.")
        data.frame()
      })
    }
    
    observeEvent(input$process_button, {
      req(input$ppes_file, input$dpp18_file, input$bud45_file, input$final_file)
      
      output$status <- renderText("Traitement en cours...")
      
      tryCatch({
        # 🗂️ Chemins
        ppes_path <- input$ppes_file$datapath
        dpp18_path <- input$dpp18_file$datapath
        bud45_path <- input$bud45_file$datapath
        final_template_path <- input$final_file$datapath
        
        code_ministere <- input$code_ministere
        code_programme <- input$code_programme
        
        # 🧾 Noms de feuilles dynamiques
        nomFeuilleSource1 <- paste0("MIN_", code_ministere, "_DETAIL_Prog_PP_CATEG")
        nomFeuilleSource2 <- paste0("MIN_", code_ministere, "_DETAIL_Prog_Entrants")
        nomFeuilleSource3 <- paste0("MIN_", code_ministere, "_DETAIL_Prog_Sortants")
        
        # 📥 Lecture
        df_pp_categ <- read_xlsx_with_recovery(ppes_path, nomFeuilleSource1)
        df_entrants <- read_xlsx_with_recovery(ppes_path, nomFeuilleSource2)
        df_sortants <- read_xlsx_with_recovery(ppes_path, nomFeuilleSource3)
        
        filtered_pp_categ <- df_pp_categ %>% filter(substr(as.character(nom_prog), 1, 3) == code_programme) %>% select(1:33)
        filtered_entrants <- df_entrants %>% filter(substr(as.character(nom_prog), 1, 3) == code_programme) %>% select(1:47)
        filtered_sortants <- df_sortants %>% filter(substr(as.character(nom_prog), 1, 3) == code_programme) %>% select(1:47)
        
        # 📊 Charger le modèle
        wb <- openxlsx2::wb_load(final_template_path)
        
        wb <- wb_add_data(wb, "Données PP-E-S", filtered_pp_categ, start_col = 3, start_row = 7)
        wb <- wb_add_data(wb, "Données PP-E-S", filtered_entrants, start_col = 3, start_row = 113)
        wb <- wb_add_data(wb, "Données PP-E-S", filtered_sortants, start_col = 3, start_row = 213)
        
        df_dpp18 <- read_xlsx_with_recovery(dpp18_path) %>% filter(str_detect(.[[1]], fixed(code_programme)))
        df_bud45 <- read_xlsx_with_recovery(bud45_path) %>% filter(str_detect(.[[1]], fixed(code_programme)))
        
        wb <- wb_add_data(wb, "INF DPP 18", df_dpp18, start_col = 2, start_row = 6)
        wb <- wb_add_data(wb, "INF BUD 45", df_bud45, start_col = 2, start_row = 6)
        
        # 👤 Indicié
        df_pp_indicie <- filtered_pp_categ %>% filter(marqueur_masse_indiciaire == "Indicié")
        
        wb <- wb_add_data(wb, "Accueil", "", dims = paste0("B43:B", 43 + nrow(df_pp_indicie)))
        wb <- wb_add_data(wb, "Accueil", "", dims = paste0("C43:C", 43 + nrow(df_pp_indicie)))
        
        wb <- wb_add_data(wb, "Accueil", df_pp_indicie[[2]], start_col = 2, start_row = 43)
        wb <- wb_add_data(wb, "Accueil", df_pp_indicie[[3]], start_col = 3, start_row = 43)
        
        # 📈 Calculs
        data_socle <- wb_read(wb, "I - Socle exécution n-1", col_names = FALSE)
        data_hyp <- wb_read(wb, "III - Hyp. salariales", col_names = FALSE)
        
        data_socle[, 3:5] <- lapply(data_socle[, 3:5], as.numeric)
        data_hyp[, 5] <- as.numeric(data_hyp[, 5])
        data_hyp[, 7:10] <- lapply(data_hyp[, 7:10], as.numeric)
        
        val_C67 <- sum(data_socle[c(34, 44, 46, 49), 3], na.rm = TRUE)
        val_C68 <- sum(data_socle[c(35, 45, 47), 3], na.rm = TRUE)
        val_D68 <- sum(data_socle[c(35, 45, 47), 4], na.rm = TRUE)
        val_E68 <- sum(data_socle[c(35, 45, 47), 5], na.rm = TRUE)
        
        val_E40 <- sum(data_hyp[c(113, 114, 115, 116), 5], na.rm = TRUE)
        vals_GHIJ40 <- sapply(7:10, function(col) sum(data_hyp[c(113, 114, 115, 116), col], na.rm = TRUE))
        
        wb <- wb_add_data(wb, "I - Socle exécution n-1", val_C67, dims = "C67")
        wb <- wb_add_data(wb, "I - Socle exécution n-1", val_C68, dims = "C68")
        wb <- wb_add_data(wb, "I - Socle exécution n-1", val_D68, dims = "D68")
        wb <- wb_add_data(wb, "I - Socle exécution n-1", val_E68, dims = "E68")
        
        wb <- wb_add_data(wb, "VI - Facteurs d'évolution MS", val_E40, dims = "E40")
        lapply(seq_along(vals_GHIJ40), function(i) {
          wb <<- wb_add_data(wb, "VI - Facteurs d'évolution MS", vals_GHIJ40[i], dims = paste0(LETTERS[7 + i - 1], "40"))
        })
        
        # 💾 Export
        output_file <- tempfile(fileext = ".xlsx")
        wb_save(wb, output_file)
        final_file_path(output_file)
        
        output$status <- renderText("✅ Traitement terminé ! Téléchargez le fichier final.")
      }, error = function(e) {
        output$status <- renderText(paste("❌ Erreur :", e$message))
      })
    })
    
    # Previews
    output$table_ppes <- renderDataTable({
      req(input$ppes_file)
      df <- read_xlsx_with_recovery(input$ppes_file$datapath,
                                    sheet = paste0("MIN_", input$code_ministere, "_DETAIL_Prog_PP_CATEG"))
      df %>% filter(substr(as.character(nom_prog), 1, 3) == input$code_programme)
    }, options = list(pageLength = 2))
    
    output$table_dpp18 <- renderDataTable({
      req(input$dpp18_file)
      read_xlsx_with_recovery(input$dpp18_file$datapath) %>%
        filter(str_detect(.[[1]], fixed(input$code_programme)))
    }, options = list(pageLength = 2))
    
    output$table_bud45 <- renderDataTable({
      req(input$bud45_file)
      read_xlsx_with_recovery(input$bud45_file$datapath) %>%
        filter(str_detect(.[[1]], fixed(input$code_programme)))
    }, options = list(pageLength = 2))
    
    output$download_final <- downloadHandler(
      filename = function() paste0("Fichier_Final_Budget_", Sys.Date(), ".xlsx"),
      content = function(file) file.copy(final_file_path(), file)
    )
  })
}