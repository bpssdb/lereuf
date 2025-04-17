setupBudgetExtraction <- function(input, output, session,
                                  rv, chat_history,
                                  dernier_fichier_contenu,
                                  donnees_extraites) {
  
  ns <- session$ns
  
  observeEvent(input$extract_budget_under_bot_clicked, {
    idx <- input$extract_budget_under_bot_clicked
    user_msg <- get_user_message_for_extraction(idx, chat_history, dernier_fichier_contenu)
    if (is.null(user_msg)) {
      showNotification("Contenu introuvable.", type = "error")
      return()
    }
    
    showNotification("Extraction en cours…", type = "message")
    
    future({
      get_budget_data(user_msg, axes = isolate(rv$axes))
    }) %...>% (function(budget_data) {
      if (is.null(budget_data) || nrow(budget_data) == 0) {
        showNotification("Aucune donnée détectée.", type = "warning")
        return()
      }
      
      # 1) on stocke sans mapping CelluleCible
      budget_data$CelluleCible <- NA_character_
      donnees_extraites(budget_data)
      
      # 2) on ouvre le modal de validation
      showModal(modalDialog(
        title    = "🗺️ Validez l’extraction budgétaire",
        # tableau éditable (seule la colonne CelluleCible restera modifiable ensuite)
        DT::DTOutput(ns("budget_table_mapping")),
        footer   = tagList(
          actionButton(ns("do_mapping_cells"), "Compléter les cellules cibles"),
          modalButton("Fermer")
        ),
        size     = "l",
        easyClose = TRUE
      ))
      
      # 3) on rend le tableau
      output$budget_table_mapping <- DT::renderDT({
        DT::datatable(
          donnees_extraites(),
          options  = list(scrollX = TRUE),
          editable = list(target = "cell", disable = list(columns = 1:6))
        )
      })
      
      # 4) édition inline de CelluleCible
      observeEvent(input$budget_table_mapping_cell_edit, {
        info <- input$budget_table_mapping_cell_edit
        df <- donnees_extraites()
        df[info$row, info$col] <- DT::coerceValue(info$value, df[info$row, info$col])
        donnees_extraites(df)
      })
      
      # 5) bouton “Compléter les cellules cibles”
      observeEvent(input$do_mapping_cells, {
        entries <- donnees_extraites()
        tags    <- isolate(rv$imported_json$tags)
        if (is.null(tags) || length(tags) == 0) {
          showNotification("⚠️ Pas de tags JSON disponibles pour le mapping.", type = "error")
          return()
        }
        mapping <- map_budget_entries(entries, tags)
        if (is.null(mapping) || nrow(mapping) == 0) {
          showNotification("❌ Mapping automatique échoué.", type = "error")
          return()
        }
        mapping_df <- as.data.frame(mapping) |>
          dplyr::select(Axe, Description, cellule) |>
          dplyr::rename(CelluleCible = cellule)
        
        merged <- dplyr::left_join(entries, mapping_df, by = c("Axe", "Description"))
        donnees_extraites(merged)
        
        # on met à jour le DT sans refermer le modal
        proxy <- DT::dataTableProxy(ns("budget_table_mapping"))
        DT::replaceData(proxy, merged, resetPaging = FALSE)
        
        showNotification("✅ Cellules cibles complétées automatiquement.", type = "message")
      })
      
    }) %...!% (function(err) {
      showNotification(paste("Erreur :", err$message), type = "error")
    })
  })
}
