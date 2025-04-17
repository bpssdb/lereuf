library(shiny)
library(jsonlite)

# ===============================
# UI du Module JSON Helper
# ===============================
mod_json_helper_ui <- function(id) {
  ns <- NS(id)
  tagList(
    fileInput(ns("import_json"), "📂 Importer un fichier JSON", accept = c(".json")),
    actionButton(ns("reload_json"), "🔄 Recharger le JSON"),
    tags$div(
      style = "max-height: 300px; overflow-y: auto; border: 1px solid #ddd; padding: 5px;",
      verbatimTextOutput(ns("json_preview"))
    ),
    actionButton(ns("show_labels"), "Afficher les résultats d'analyse"),
    actionButton(ns("update_tags"), "🛠️ Actualiser les tags à partir des cellules sources")
  )
}

# ===============================
# Server du Module JSON Helper
# ===============================
mod_json_helper_server <- function(id, rv) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    imported_data <- reactiveVal(NULL)
    
    load_json_file <- function(file_path) {
      tryCatch({
        json_data <- fromJSON(file_path, simplifyVector = FALSE)
        imported_data(json_data)
        showNotification("✅ JSON importé avec succès !", type = "message")
      }, error = function(e) {
        showNotification(paste("❌ Erreur lors de l'importation :", e$message), type = "error")
      })
    }
    
    observeEvent(input$import_json, {
      req(input$import_json)
      load_json_file(input$import_json$datapath)
      rv$imported_json <- imported_data()
    })
    
    observeEvent(input$reload_json, {
      req(imported_data())
      load_json_file(input$import_json$datapath)
      rv$imported_json <- imported_data()
    })
    
    output$json_preview <- renderPrint({
      req(imported_data())
      toJSON(imported_data(), pretty = TRUE, auto_unbox = TRUE)
    })
    
    observe({
      rv$imported_json <- imported_data()
    })
    
    observe({
      req(imported_data())
      json_data <- imported_data()
      if (!is.list(json_data) || !("tags" %in% names(json_data))) {
        showNotification("⚠️ Le fichier JSON n'est pas structuré comme prévu (clé 'tags' manquante).", type = "error")
        return(NULL)
      }
      tags_list <- json_data$tags
      if (is.data.frame(tags_list)) {
        labels <- unique(na.omit(tags_list$labels))
      } else if (is.list(tags_list)) {
        labels <- unique(na.omit(unlist(lapply(tags_list, function(tag) tag$labels))))
      } else {
        labels <- character(0)
      }
      rv$extracted_labels <- labels
      showNotification("✅ Labels extraits du JSON avec succès !", type = "message")
    })
    
    observeEvent(input$show_labels, {
      labels <- rv$extracted_labels
      if (is.null(labels) || length(labels) == 0) {
        labels_table <- "Aucun label n'a été extrait."
      } else {
        labels_table <- tags$table(
          class = "table table-striped",
          tags$thead(tags$tr(tags$th("Label"))),
          tags$tbody(
            lapply(labels, function(label) {
              tags$tr(tags$td(paste(as.character(label), collapse = ", ")))
            })
          )
        )
      }
      
      if (!is.null(rv$axes)) {
        if (is.data.frame(rv$axes)) {
          axes_rows <- lapply(seq_len(nrow(rv$axes)), function(i) {
            tags$tr(
              tags$td(as.character(rv$axes$axe[i])),
              tags$td(as.character(rv$axes$description[i]))
            )
          })
        } else if (is.list(rv$axes)) {
          axes_rows <- lapply(rv$axes, function(x) {
            axe_text <- if (!is.null(x$axe)) as.character(x$axe) else ""
            desc_text <- if (!is.null(x$description)) as.character(x$description) else ""
            tags$tr(tags$td(axe_text), tags$td(desc_text))
          })
        } else {
          axes_rows <- NULL
        }
        axes_table <- tags$table(
          class = "table table-striped",
          tags$thead(tags$tr(tags$th("Axe"), tags$th("Description"))),
          tags$tbody(axes_rows)
        )
      } else {
        axes_table <- "Aucun axe n'a été extrait."
      }
      
      modal_content <- tagList(
        h3("Labels extraits"),
        labels_table,
        br(),
        h3("Axes d'analyse"),
        axes_table
      )
      
      showModal(modalDialog(
        title = "Résultats de l'analyse",
        modal_content,
        easyClose = TRUE,
        footer = modalButton("Fermer")
      ))
    })
    
    observeEvent(input$update_tags, {
      req(imported_data(), rv$imported_json, rv$excel_data)
      json_data <- imported_data()
      excel_data <- rv$excel_data
      
      if (!"tags" %in% names(json_data)) {
        showNotification("⚠️ Clé 'tags' absente dans le JSON.", type = "error")
        return()
      }
      if (!is.data.frame(excel_data)) {
        showNotification("⚠️ Le contenu Excel est invalide.", type = "error")
        return()
      }
      
      colnames(excel_data) <- LETTERS[seq_len(ncol(excel_data))]
      
      ajouts <- list()
      updated_tags <- lapply(seq_along(json_data$tags), function(i) {
        tag <- json_data$tags[[i]]
        existing_labels <- if (!is.null(tag$labels)) tag$labels else character(0)
        source_labels <- c()
        
        if (!is.null(tag$source_cells)) {
          source_labels <- sapply(tag$source_cells, function(cell_address) {
            if (grepl("^[A-Z]+[0-9]+$", cell_address)) {
              col_letter <- gsub("[0-9]", "", cell_address)
              row_num <- as.integer(gsub("[A-Z]", "", cell_address))
              col_index <- match(col_letter, colnames(excel_data))
              if (!is.na(col_index) && row_num <= nrow(excel_data)) {
                return(as.character(excel_data[[col_index]][row_num]))
              }
            }
            return(NA)
          })
        }
        
        new_labels <- unique(c(existing_labels, na.omit(source_labels)))
        diff <- setdiff(new_labels, existing_labels)
        
        if (length(diff) > 0) {
          ajouts[[length(ajouts) + 1]] <<- list(
            cell = if (!is.null(tag$cell_address)) tag$cell_address else paste0("Tag ", i),
            added = diff
          )
        }
        
        tag$labels <- new_labels
        tag
      })
      
      json_data$tags <- updated_tags
      imported_data(json_data)
      rv$imported_json <- json_data
      
      showNotification("✅ Tags enrichis avec les labels Excel (sans écraser les existants) !", type = "message")
      
      if (length(ajouts) > 0) {
        resume_lignes <- lapply(ajouts, function(a) {
          tags$p(HTML(paste0("📌 <b>", a$cell, "</b> : + ", paste(shQuote(a$added), collapse = ", "))))
        })
        
        showModal(modalDialog(
          title = "🆕 Labels ajoutés aux tags",
          tagList(
            tags$p(paste0("✅ ", sum(sapply(ajouts, function(a) length(a$added))),
                          " nouveau(x) label(s) ajouté(s) dans ",
                          length(ajouts), " cellule(s).")),
            tags$hr(),
            resume_lignes
          ),
          easyClose = TRUE,
          footer = modalButton("Fermer")
        ))
      } else {
        showNotification("Aucun nouveau label n’a été ajouté.", type = "message")
      }
    })
  })
}
