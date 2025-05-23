# mod_mesures_cat.R: Module Excel refactorisé, UI épurée et sans espace vide

library(shiny)
library(readxl)
library(reactable)
library(openxlsx2)
library(rhandsontable)
library(shinycssloaders)

# ============================
# MODULE UI
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
# MODULE SERVER
# ============================

mod_mesures_cat_server <- function(id, rv, on_analysis_summary = NULL) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    
    # === Réactifs internes ===
    rv_path        <- reactiveVal(NULL)
    rv_path_script <- reactiveVal(NULL)
    rv_sheets      <- reactiveVal(NULL)
    rv_table       <- reactiveVal(NULL)
    rv_selected    <- reactiveVal(NULL)
    
    # === Chargement initial du fichier Excel ===
    observeEvent(input$upload_file, {
      req(input$upload_file)
      path <- input$upload_file$datapath
      ext  <- tools::file_ext(input$upload_file$name)
      
      if (ext != "xlsx") {
        showNotification("❌ Veuillez charger un fichier Excel (.xlsx)", type = "error")
        return()
      }
      
      wb <- tryCatch(openxlsx2::wb_load(path), error = function(e) NULL)
      if (inherits(wb, "wbWorkbook")) {
        rv_path(path)
        rv$fichier_excel <- wb
        
        sheets <- tryCatch(openxlsx2::wb_get_sheet_names(wb), error = function(e) NULL)
        rv_sheets(sheets)
        updateSelectInput(session, ns("selected_sheet"), choices = sheets, selected = sheets[1])
      } else {
        showNotification("❌ Échec du chargement du fichier.", type = "error")
      }
      # === Création du fichier des formules sous-jacentes de l'excel ===
      parse_excel_formulas(rv_path(), emit_script = TRUE)
      rv_path_script(paste0(tools::file_path_sans_ext(basename(rv_path())), "_converted_formulas.R"))
    })
    
    
    # === Rendu du sélecteur de feuilles ===
    output$sheet_selector <- renderUI({
      req(rv_sheets())
      selectInput(ns("selected_sheet"), "🗂️ Choisir une feuille", choices = rv_sheets(), width = "100%")
    })
    
    # Quand l’utilisateur change de feuille
    observeEvent(input$selected_sheet, {
      req(rv$fichier_excel)
      sheet <- input$selected_sheet
      df    <- openxlsx2::wb_to_df(rv$fichier_excel, sheet = sheet)
      rv_selected(sheet)
      rv_table(df)
    }, ignoreNULL = TRUE)
    
    # Quand on a écrit dans rv$fichier_excel
    observeEvent(rv$excel_updated, {
      print("Jerentrela")
      req(rv$fichier_excel, rv_selected())
      wb    <- rv$fichier_excel
      print(rv$fichier_excel)
      
      #  Mettre à jour la liste des feuilles
      sheets <- tryCatch(openxlsx2::wb_get_sheet_names(wb), error = function(e) NULL)
      rv_sheets(sheets)
      # si vous voulez forcer la mise à jour du sélecteur aussi :
      updateSelectInput(session, ns("selected_sheet"),
                        choices = sheets,
                        selected = rv_selected())
      
      # Rafraîchir la table de la feuille active
      sheet <- rv_selected()
      df    <- openxlsx2::wb_to_df(wb, sheet = sheet)
      print(df)
      rv_table(df)
    }, ignoreNULL = FALSE)
    
    
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
        return(reactable(data.frame(Message = "📭 Aucune donnée à afficher ouin"), bordered = TRUE))
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
    
    # === Boutons ===
    output$buttons_container <- renderUI({
      req(rv_table())
      fluidRow(
        column(12,
               div(class = "d-flex",
                   actionButton(ns("open_full_editor"), "🖋️ Modifier la feuille", class = "btn btn-secondary mr-2"),
                   actionButton(ns("apply_formulas"), "🖋 ️Appliquer les formules", class = "btn btn-secondary mr-2"),
                   downloadButton(ns("download_table"), "💾 Exporter le tableau", class = "btn btn-success")
               )
        )
      )
    })
    
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
    
    observeEvent(input$apply_formulas, {
      path_script <- rv_path_script()
      if (!is.null(path_script) && file.exists(path_script)) {
        source(path_script)
        script(rv)
      } else {
        showNotification("❌ Script de formules introuvable ou non généré.", type = "error")
      }
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
      df    <- hot_to_r(input$hot_table)
      sheet <- rv_selected()
      
      # 1) Mise à jour du data.frame local
      rv_table(df)
      
      # 2) Injection dans le workbook global
      wb <- rv$fichier_excel
      wb <- openxlsx2::wb_remove_worksheet(wb, sheet)
      wb <- openxlsx2::wb_add_worksheet(wb, sheet)
      wb <- openxlsx2::wb_add_data(wb, sheet = sheet, x = df)
      
      rv$fichier_excel <- wb
      rv$excel_updated <- Sys.time()
      removeModal()
    })
    
    
    output$download_table <- downloadHandler(
      filename = function() {
        paste0("sortie_budgibot_", Sys.Date(), ".xlsx")
      },
      content = function(file) {
        wb <- rv$fichier_excel # ⚡ Ici on récupère directement le wb en mémoire
        req(inherits(wb, "wbWorkbook"))
        
        # ⚡ Enregistre directement le workbook que l'utilisateur manipule
        openxlsx2::wb_save(wb, file = file, overwrite = TRUE)
        
        showNotification("✅ Export Excel terminé", type = "message")
      }
    )
    
  })
}
