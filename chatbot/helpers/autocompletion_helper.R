library(stringi)
library(stringr)
#Quelques fonctions auxiliaires sont dans llm_utils.

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
      get_budget_data(user_msg)
    }) %...>% (function(budget_data) {
      if (is.null(budget_data) || nrow(budget_data) == 0) {
        showNotification("Aucune donnée détectée.", type = "warning")
        return()
      }
      
      # 1) on ajoute SourcePhrase pour chaque ligne, via l'analyse du message source
      budget_data <- attach_source_phrases(budget_data, user_msg)
      
      # 2) on complète les autres champs
      budget_data$Tags <- NA_character_
      budget_data$Id <- NA_character_
      budget_data$CelluleCible <- NA_character_
      
      donnees_extraites(budget_data)
      
      # 2) on ouvre le modal de validation
      showModal(modalDialog(
        title = "🗺️ Validez l’extraction budgétaire",
        size = "l", easyClose = TRUE,
        footer = tagList(
          actionButton(ns("do_mapping_cells"), "Compléter les cellules cibles"),
          actionButton(ns("write_to_excel"), "Écrire dans Excel"),
          modalButton("Fermer")
        ),
        div(style = "overflow-x: auto; max-width:100%",
            DT::DTOutput(ns("budget_table_mapping"))
        )
      ))
      
      
      
      # 3) on rend le tableau
      output$budget_table_mapping <- DT::renderDT({
        DT::datatable(
          donnees_extraites(),
          options = list(
            scrollX = TRUE
          ),
          editable = list(target = "cell")
        )
        
      })
      
      # 4) édition inline du tableau
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
        
        # on cast en data.frame et on sélectionne les 5 colonnes dont on a besoin :
        mapping_df <- as.data.frame(mapping) %>%
          dplyr::select(
            Axe,
            Description,
            cellule,
            tag_id,
            tags_utilisés
          ) %>%
          dplyr::rename(
            CelluleCible  = cellule,
            TagID         = tag_id,
            TagsUtilises  = tags_utilisés
          )
        
        # on fusionne en conservant toutes les colonnes d'origine plus les 3 nouvelles
        merged <- dplyr::left_join(
          entries %>% dplyr::select(-CelluleCible, -Id, -Tags),
          mapping_df,
          by = c("Axe", "Description")
        )
        
        donnees_extraites(merged)
        
        # on met à jour le DT sans refermer le modal
        proxy <- DT::dataTableProxy(ns("budget_table_mapping"))
        DT::replaceData(proxy, merged, resetPaging = FALSE)
        
        showNotification("✅ Cellules cibles (et leurs tags) complétées automatiquement.", type = "message")
      })
    })
  })
  # ————————————————————————————————
  # 6) bouton “Écrire dans Excel”
  observeEvent(input$write_to_excel, {
    req(rv$fichier_excel, donnees_extraites(), rv$imported_json$tags)
    print(rv$fichier_excel)
    entries <- donnees_extraites()
    tags    <- rv$imported_json$tags
    
    # on (re)construit le mapping pour avoir tag_id, sheet_name et cell_address
    mapping <- map_budget_entries(entries, tags)
    if (is.null(mapping) || nrow(mapping) == 0) {
      showNotification("❌ Pas de mapping disponible pour écrire dans Excel.", type = "error")
      return()
    }
    
    wb <- rv$fichier_excel  # Assure-toi que ceci est bien défini avant la boucle
    
    for (i in seq_len(nrow(mapping))) {
      this_tag <- tags[[ mapping$tag_id[i] ]]
      sheet    <- this_tag$sheet_name
      addr     <- this_tag$cell_address
      
      coords <- parse_address(addr)
      value  <- entries$Montant[i]
      
      print(sprintf("Écriture dans %s!R%dC%d : %s", sheet, coords$row, coords$col, value))
      
      wb <-  openxlsx2::wb_add_data(
        wb,
        sheet    = sheet,
        x        = value,
        start_col = coords$col,
        start_row = coords$row
      )
    }
    
    rv$fichier_excel <- wb  # IMPORTANT : bien remettre à jour le reactiveValue
    rv$excel_updated <- Sys.time()

    showNotification("✅ Les montants ont été écrits dans votre Excel en mémoire.", type = "message")
  })
  
}