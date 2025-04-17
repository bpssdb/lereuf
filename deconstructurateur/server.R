server <- function(input, output, session) {
  
  # === Valeurs réactives globales ===
  rv <- reactiveValues(
    all_sheets = NULL,
    df = NULL,
    selected = NULL,             # Zone sélectionnée (coordonnées)
    tags = list(),               # Tags définitifs (fusion des labels auto et des headers manuels)
    current_rows = NULL,
    current_cols = NULL,
    header_highlights = list(),  # En-têtes automatiques détectées (cellules directement au-dessus/gauche)
    manual_trigger = Sys.time(), # Déclencheur de re-rendu (pour actualiser le tableau)
    json_data = NULL,            # Pour l'export JSON
    edited_cells = list()        # Liste des cellules sources (headers manuels) sélectionnées par l'utilisateur
  )
  
  ### 1. Fonctions auxiliaires
  
  # La fonction add_or_update_tags() met à jour ou ajoute un tag pour une cellule,
  # en fusionnant (via union) les nouveaux labels avec ceux déjà existants.
  add_or_update_tags <- function(r, c, new_labels, new_header_cells = NULL, new_type = NULL, new_emoji = NULL) {
    if(length(rv$tags) == 0) {
      idx <- integer(0)
    } else {
      idx <- which(vapply(rv$tags, function(tag) {
        if(is.list(tag) && !is.null(tag$row) && !is.null(tag$col)) {
          return(tag$row == r && tag$col == c)
        } else {
          return(FALSE)
        }
      }, FUN.VALUE = logical(1)))
    }
    
    if(length(idx) > 0) {
      current_labels <- rv$tags[[idx]]$labels
      combined_labels <- unique(c(current_labels, new_labels))
      rv$tags[[idx]]$labels <- combined_labels
      if(!is.null(new_header_cells)) {
        current_header <- if(!is.null(rv$tags[[idx]]$header_cells)) rv$tags[[idx]]$header_cells else character(0)
        rv$tags[[idx]]$header_cells <- unique(c(current_header, new_header_cells))
      }
      if(!is.null(new_type)) rv$tags[[idx]]$type <- new_type
      if(!is.null(new_emoji)) rv$tags[[idx]]$emoji <- new_emoji
      print(paste("🔄 Mise à jour du tag pour cellule [", r, ",", c, "]:", 
                  paste(combined_labels, collapse = "; ")))
    } else {
      rv$tags <- c(rv$tags, list(list(
        row = r,
        col = c,
        labels = new_labels,
        header_cells = new_header_cells,
        type = new_type,
        emoji = new_emoji
      )))
      print(paste("✅ Nouveau tag ajouté pour cellule [", r, ",", c, "]:",
                  paste(new_labels, collapse="; ")))
    }
  }
  
  # Fonction pour mettre à jour rv$selected à partir de la sélection (issue de rhandsontable)
  updateSelectedZone <- function(sel) {
    if(!is.null(sel) &&
       !is.null(sel$r) && !is.null(sel$r2) &&
       !is.null(sel$c) && !is.null(sel$c2)) {
      rv$selected <- list(
        startRow = sel$r,
        endRow = sel$r2,
        startCol = sel$c,
        endCol = sel$c2
      )
    } else {
      rv$selected <- NULL
    }
  }
  
  # Fonction de rendu du tableau Excel avec indicateurs visuels :
  # - "🔷" indique la zone auto-détectée,
  # - "🏷️" indique les headers automatiques,
  # - "✏️" indique les cellules sources sélectionnées (headers manuels).
  renderExcelTable <- function() {
    req(rv$df, rv$manual_trigger)
    df <- rv$df
    df_char <- as.data.frame(lapply(df, as.character), stringsAsFactors = FALSE)
    
    if(nrow(df_char) == 0 || ncol(df_char) == 0) {
      showNotification("📭 Aucune donnée trouvée.", type="warning")
      return(NULL)
    }
    
    # Affichage de la zone auto-détectée (préfixe bleu "🔷")
    if(!is.null(rv$current_rows) && !is.null(rv$current_cols)) {
      for(r in rv$current_rows) {
        for(c in rv$current_cols) {
          if(r <= nrow(df_char) && c <= ncol(df_char)) {
            val <- as.character(df_char[r, c])
            if(is.na(val)) { val <- "" }
            if(nzchar(val) && !startsWith(val, "🔷")) {
              df_char[r, c] <- paste0("🔷 ", val)
            } else if(!nzchar(val)) {
              df_char[r, c] <- "🔷"
            }
          }
        }
      }
    }
    
    # Affichage des en-têtes automatiques (préfixe "🏷️")
    if(!is.null(rv$header_highlights) && length(rv$header_highlights) > 0) {
      for(cell in rv$header_highlights) {
        if(!is.list(cell) || is.null(cell$row) || is.null(cell$col)) next
        r <- cell$row; c <- cell$col
        if(r <= nrow(df_char) && c <= ncol(df_char)) {
          val <- as.character(df_char[r, c])
          if(is.na(val)) { val <- "" }
          if(nzchar(val) && !startsWith(val, "🏷️")) {
            df_char[r, c] <- paste0("🏷️ ", val)
          }
        }
      }
    }
    
    # Affichage des headers manuels (cellules sélectionnées) : préfixe "✏️ "
    if(length(rv$edited_cells) > 0) {
      for(cell in rv$edited_cells) {
        r_edit <- cell$row; c_edit <- cell$col
        if(r_edit <= nrow(df_char) && c_edit <= ncol(df_char)) {
          val <- as.character(df_char[r_edit, c_edit])
          if(!startsWith(val, "✏️ ")) {
            df_char[r_edit, c_edit] <- paste0("✏️ ", val)
          }
        }
      }
    }
    
    rhandsontable::rhandsontable(df_char, width = "1200px", height = "600px", selectCallback = TRUE) %>% 
      rhandsontable::hot_table(selection = list(mode = "range"))
  }
  
  ### 2. Observers et gestion des événements
  
  # Chargement du fichier Excel et détection des feuilles
  observeEvent(input$upload_file, {
    req(input$upload_file)
    rv$all_sheets <- readxl::excel_sheets(input$upload_file$datapath)
  })
  
  output$sheet_selector <- renderUI({
    req(rv$all_sheets)
    selectInput("selected_sheet", "Feuilles disponibles", choices = rv$all_sheets)
  })
  
  observeEvent(input$load_sheet, {
    req(input$upload_file, input$selected_sheet)
    rv$df <- readxl::read_excel(input$upload_file$datapath,
                                sheet = input$selected_sheet, col_names = FALSE)
  })
  
  # Mise à jour de la zone sélectionnée depuis rhandsontable
  observe({
    updateSelectedZone(input$excel_table_select$select)
    print("Sélection capturée :")
    print(rv$selected)
  })
  
  output$excel_table <- renderRHandsontable({
    tryCatch({
      renderExcelTable()
    }, error = function(e) {
      showNotification(paste("Erreur dans le rendu du tableau :", e$message), type = "error")
      return(NULL)
    })
  })
  
  output$selection_info <- renderText({
    req(rv$selected)
    sel <- rv$selected
    paste("Sélection : Lignes", sel$startRow, "à", sel$endRow,
          "| Colonnes", num_to_excel_col(sel$startCol), "à", num_to_excel_col(sel$endCol))
  })
  
  output$zone_stats <- renderText({
    req(rv$current_rows, rv$current_cols)
    paste("Cellules analysées :", length(rv$current_rows) * length(rv$current_cols))
  })
  
  output$tags_view <- renderReactable({
    req(rv$tags)
    if(length(rv$tags) == 0) {
      showNotification("📭 Aucun tag trouvé.", type="warning")
      return(NULL)
    }
    tag_list <- lapply(rv$tags, function(tag) {
      if(is.list(tag) && all(c("row", "col", "labels") %in% names(tag))) {
        data.frame(
          Row = tag$row,
          Col = tag$col,
          Cell = paste0(num_to_excel_col(tag$col), tag$row),
          Labels = paste(unlist(tag$labels), collapse = "; "),
          SourceCells = if(!is.null(tag$header_cells)) paste(tag$header_cells, collapse = "; ") else "",
          stringsAsFactors = FALSE
        )
      } else NULL
    })
    tag_list <- Filter(Negate(is.null), tag_list)
    if(length(tag_list) > 0) {
      reactable::reactable(do.call(rbind, tag_list),
                           searchable = TRUE, pagination = TRUE, highlight = TRUE)
    } else {
      reactable::reactable(data.frame(Dummy = character(0)))
    }
  })
  
  # Passage en mode édition (bouton "toggle_edit_labels" dans l'UI)
  observeEvent(input$toggle_edit_labels, {
    rv$edit_labels_mode <- !rv$edit_labels_mode
    msg <- if(rv$edit_labels_mode) {
      "Mode Édition Labels activé. Sélectionnez une seule cellule pour ajouter ou retirer un header manuel."
    } else {
      "Mode Édition Labels désactivé."
    }
    showNotification(msg, type = "message")
  })
  
  # Observer pour le bouton "toggle_label" : permet de sélectionner/déselectionner des headers manuels
  observeEvent(input$toggle_label, {
    req(rv$selected, rv$df)
    # On exige que la sélection corresponde à une seule cellule pour définir un header manuel
    if(rv$selected$startRow == rv$selected$endRow && rv$selected$startCol == rv$selected$endCol) {
      current_cell <- list(row = rv$selected$startRow, col = rv$selected$startCol)
      already <- sapply(rv$edited_cells, function(cell) {
        cell$row == current_cell$row && cell$col == current_cell$col
      })
      if(any(already)) {
        rv$edited_cells <- rv$edited_cells[!already]
        showNotification("Header manuel retiré.", type = "message")
      } else {
        rv$edited_cells <- c(rv$edited_cells, list(current_cell))
        showNotification("Header manuel ajouté.", type = "message")
      }
    } else {
      showNotification("Veuillez sélectionner une seule cellule pour définir un header manuel.", type = "error")
    }
    rv$manual_trigger <- Sys.time()
    print("Headers manuels sélectionnés (rv$edited_cells) :")
    print(rv$edited_cells)
  })
  
  # Fusion définitive des headers lors du clic sur "Taguer cette zone"
  observeEvent(input$tag_zone, {
    req(rv$current_rows, rv$current_cols, rv$df)
    df <- rv$df
    header_row_index <- min(rv$current_rows) - 1
    header_col_index <- min(rv$current_cols) - 1
    
    # Calcul de l'union des headers manuels
    manual_label_union <- character(0)
    manual_address_union <- character(0)
    if(length(rv$edited_cells) > 0) {
      for(cell in rv$edited_cells) {
        lbl <- as.character(df[cell$row, cell$col])
        if(!is.na(lbl) && nzchar(lbl)) {
          manual_label_union <- union(manual_label_union, lbl)
          manual_address_union <- union(manual_address_union, paste0(num_to_excel_col(cell$col), cell$row))
        }
      }
      print(paste("Union des headers manuels :", paste(manual_label_union, collapse = "; ")))
    }
    
    # Pour chaque cellule de la zone auto-détectée, fusionner les auto-labels et les headers manuels
    for(r in rv$current_rows) {
      for(c in rv$current_cols) {
        auto_labels <- c()
        header_cells <- list()
        # Auto-label issu de la ligne d'en-tête
        if(min(rv$current_rows) > 1 && header_row_index <= nrow(df)) {
          lbl <- as.character(df[header_row_index, c])
          if(!is.na(lbl) && nzchar(lbl)) {
            auto_labels <- c(auto_labels, lbl)
            header_cells <- c(header_cells, paste0(num_to_excel_col(c), header_row_index))
          }
        }
        # Auto-label issu de la colonne d'en-tête
        if(min(rv$current_cols) > 1 && header_col_index <= ncol(df)) {
          lbl <- as.character(df[r, header_col_index])
          if(!is.na(lbl) && nzchar(lbl)) {
            auto_labels <- c(auto_labels, lbl)
            header_cells <- c(header_cells, paste0(num_to_excel_col(header_col_index), r))
          }
        }
        # Fusion des labels : union des auto-labels et des headers manuels
        final_labels <- union(auto_labels, manual_label_union)
        # Fusion des adresses sources : union des adresses auto et manuelles
        final_source_cells <- union(header_cells, manual_address_union)
        add_or_update_tags(r = r, c = c,
                           new_labels = final_labels,
                           new_header_cells = final_source_cells,
                           new_type = "automatique",
                           new_emoji = "🔷")
      }
    }
    
    showNotification(
      paste(length(rv$current_rows) * length(rv$current_cols), "tags fusionnés et enregistrés ✅"),
      type = "message"
    )
    
    # Nettoyage visuel : on retire uniquement l'emoji (affichage) pour les tags de la zone tagguée,
    # mais on conserve les tags dans rv$tags
    rv$tags <- lapply(rv$tags, function(tag) {
      tag$emoji <- ""
      return(tag)
    })
    # Réinitialiser les variables de zone pour un nouveau tagging, sans toucher aux tags déjà enregistrés
    rv$edited_cells <- list()
    rv$selected <- NULL
    rv$current_rows <- NULL
    rv$current_cols <- NULL
    rv$header_highlights <- list()
    rv$manual_trigger <- Sys.time()
  })
  
  # Détection automatique de la zone (gestion des en-têtes)
  observeEvent(input$auto_detect_zone, {
    req(rv$df, rv$selected)
    sel <- rv$selected
    if(is.null(sel$startRow) || is.null(sel$endRow) ||
       is.null(sel$startCol) || is.null(sel$endCol)) {
      showNotification("⚠️ La sélection est invalide.", type = "error")
      return()
    }
    rows <- seq(sel$startRow, sel$endRow)
    cols <- seq(sel$startCol, sel$endCol)
    rv$current_rows <- rows
    rv$current_cols <- cols
    if(length(rows) == 0 || length(cols) == 0 ||
       any(is.na(rows)) || any(is.na(cols)) || any(rows < 1) || any(cols < 1)) {
      showNotification("⚠️ Sélection invalide.", type = "error")
      return()
    }
    # Définir les en-têtes : la ligne et la colonne immédiatement avant la zone
    header_row <- max(1, min(rows) - 1)
    header_col <- max(1, min(cols) - 1)
    header_cells <- list()
    if(header_row < min(rows)) {
      header_cells <- c(header_cells, lapply(cols, function(cc) list(row = header_row, col = cc)))
    }
    if(header_col < min(cols)) {
      header_cells <- c(header_cells, lapply(rows, function(rr) list(row = rr, col = header_col)))
    }
    rv$header_highlights <- header_cells
    # Pour chaque cellule de la zone, si aucun tag n'existe, on l'ajoute sans écraser les tags existants
    for(r in rows) {
      for(c in cols) {
        tag_idx <- which(vapply(rv$tags, function(tag) {
          if(is.list(tag) && !is.null(tag$row) && !is.null(tag$col)) {
            return(tag$row == r && tag$col == c)
          } else {
            return(FALSE)
          }
        }, FUN.VALUE = logical(1)))
        if(length(tag_idx) == 0) {
          add_or_update_tags(r = r, c = c,
                             new_labels = character(0),
                             new_header_cells = NULL,
                             new_type = "temp",
                             new_emoji = "🔷")
        }
      }
    }
    showNotification(
      paste(length(rows)*length(cols), "tags auto-détectés et enregistrés (temporairement) ✅"),
      type = "message"
    )
  })
  
  # Ajout d'un tag manuel ponctuel via modal
  observeEvent(input$manual_tag, {
    req(rv$selected, rv$df)
    sel <- rv$selected
    rows <- sel$startRow:sel$endRow
    cols <- sel$startCol:sel$endCol
    showModal(modalDialog(
      title = "Ajouter un tag manuel",
      textInput("manual_tag_text", "Tag manuel (texte libre) :", value = ""),
      footer = tagList(
        modalButton("Annuler"),
        actionButton("confirm_manual_tag", "Confirmer")
      )
    ))
    
    observeEvent(input$confirm_manual_tag, {
      manual_tag_text <- input$manual_tag_text
      if(nzchar(manual_tag_text)) {
        for(r in rows) {
          for(c in cols) {
            add_or_update_tags(r = r, c = c,
                               new_labels = list(manual_tag_text),
                               new_type = "manuel",
                               new_emoji = "📝")
          }
        }
        showNotification("✅ Tag manuel ajouté avec succès !", type = "message")
        removeModal()
      } else {
        showNotification("❌ Le texte du tag manuel est vide.", type = "error")
      }
    }, once = TRUE)
  })
  
  # Génération du JSON via le module de téléchargement
  jsonData <- downloadModuleServer("download1",
                                   df = reactive({ rv$df }),
                                   tags_data = reactive({ rv$tags }),
                                   selected_sheet = reactive({ input$selected_sheet }))
}