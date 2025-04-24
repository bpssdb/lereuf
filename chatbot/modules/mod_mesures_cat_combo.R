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
    rv_path     <- reactiveVal(NULL)
    rv_excel    <- reactiveVal(NULL)
    rv_sheets   <- reactiveVal(NULL)
    rv_table    <- reactiveVal(NULL)
    rv_selected <- reactiveVal(NULL)
    
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
        rv_excel(wb)
        rv$fichier_excel <- wb
        
        sheets <- tryCatch(openxlsx2::wb_get_sheet_names(wb), error = function(e) NULL)
        rv_sheets(sheets)
        updateSelectInput(session, ns("selected_sheet"), choices = sheets, selected = sheets[1])
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
        wb_read(rv_excel(), sheet = input$selected_sheet, col_names = TRUE),
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
        df <- tryCatch(wb_read(wb, sheet = sheet, col_names = TRUE), error = function(e) NULL)
        if (!is.null(df)) rv_table(as.data.frame(df))
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
    
    output$hot_table <- renderRHandsontable({
      req(rv_table())
      rhandsontable(rv_table(), useTypes = TRUE, stretchH = "all") %>%
        hot_cols(colWidths = rep(120, ncol(rv_table())))
    })
    
    observeEvent(input$save_edits, {
      req(input$hot_table)
      rv_table(hot_to_r(input$hot_table))
      removeModal()
    })
    
    output$download_table <- downloadHandler(
      filename = function() paste0("mesures_cat_", Sys.Date(), ".xlsx"),
      content  = function(file) {
        wb <- wb_workbook()
        wb_add_worksheet(wb, "Mesures")
        wb_add_data(wb, sheet = 1, x = rv_table())
        wb_save(wb, file)
        showNotification("✅ Export Excel terminé", type = "message")
      }
    )
  })
}
