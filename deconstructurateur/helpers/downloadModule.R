# downloadModule.R

# UI du module pour le téléchargement
downloadModuleUI <- function(id) {
  ns <- NS(id)
  tagList(
    actionButton(ns("preview_json"), "Prévisualiser JSON"),
    downloadButton(ns("download_tags"), "Télécharger JSON")
  )
}

# Partie serveur du module pour le téléchargement
downloadModuleServer <- function(id, df, tags_data, selected_sheet) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    
    json_data <- reactive({
      req(df(), tags_data(), selected_sheet())
      tags_export <- lapply(tags_data(), function(tag) {
        # On s'assure que les clés "row", "col", "labels" et "header_cells" (pour les sources) existent
        if (!all(c("row", "col", "labels") %in% names(tag))) return(NULL)
        # Calcul de l'adresse de la cellule tagguée au format Excel (ex : "A1")
        cell_address <- paste0(num_to_excel_col(tag$col), tag$row)
        # On récupère les cellules sources (positions des headers) si elles existent
        source_cells <- if (!is.null(tag$header_cells)) tag$header_cells else character(0)
        list(
          row = tag$row,
          col = tag$col,
          cell_address = cell_address,
          labels = tag$labels,
          source_cells = source_cells
        )
      })
      tags_export <- Filter(Negate(is.null), tags_export)
      export_data <- list(
        sheet_name = selected_sheet(),
        date_export = as.character(Sys.Date()),
        tags = tags_export
      )
      jsonlite::toJSON(export_data, pretty = TRUE, auto_unbox = TRUE)
    })
    
    output$download_tags <- downloadHandler(
      filename = function() {
        paste0("tags_", Sys.Date(), ".json")
      },
      content = function(file) {
        write(json_data(), file)
      }
    )
    
    observeEvent(input$preview_json, {
      showModal(modalDialog(
        title = "Prévisualisation du JSON",
        tags$textarea(
          id = ns("json_preview"),
          json_data(),
          style = "width:100%; height:400px;"
        ),
        footer = modalButton("Fermer"),
        size = "l"
      ))
    })
    
    return(json_data)
  })
}