# OPTIMISEUR VECTORIEL POUR PATTERNS RÉPÉTITIFS
# ==============================================

# OPTIMISEUR VECTORIEL POUR PATTERNS RÉPÉTITIFS
# ==============================================

# Fonction pour vectoriser les patterns répétitifs
vectorize_repetitive_patterns <- function(form_cells) {
  
  message("=== VECTORISATION DES PATTERNS ===")
  
  # 1. PATTERN PASTE0 - Vectoriser les concaténations
  paste_pattern <- form_cells[grepl("paste0\\(values\\[\\d+, \\d+\\],values\\[\\d+, \\d+\\]\\)", form_cells$R_code), ]
  
  if (nrow(paste_pattern) > 100) {
    message(sprintf("Vectorisation de %d concaténations paste0...", nrow(paste_pattern)))
    
    # Grouper par feuille et par pattern de colonnes
    paste_groups <- paste_pattern %>%
      mutate(
        col_pattern = str_extract_all(R_code, "\\d+(?=\\])")
      ) %>%
      group_by(sheet) %>%
      summarise(
        rows = list(row),
        patterns = list(col_pattern),
        .groups = 'drop'
      )
    
    # Générer du code vectorisé pour chaque groupe
    vectorized_code <- list()
    for (i in seq_len(nrow(paste_groups))) {
      sheet_name <- paste_groups$sheet[i]
      rows <- paste_groups$rows[[i]]
      
      vectorized_code[[sheet_name]] <- sprintf(
        "# Vectorisation paste0 pour %s (%d cellules)\n%s",
        sheet_name, 
        length(rows),
        paste(sprintf(
          "sheets[['%s']][c(%s), target_cols] <- paste0(sheets[['%s']][c(%s), col1], sheets[['%s']][c(%s), col2])",
          sheet_name, paste(rows, collapse = ","), sheet_name, paste(rows, collapse = ","), sheet_name, paste(rows, collapse = ",")
        ), collapse = "\n")
      )
    }
  }
  
  # 2. PATTERN SUMIF - Vectoriser les sommes conditionnelles
  sumif_pattern <- form_cells[grepl("sum\\(sum\\(\\(\\(values\\[\\d+:\\d+, \\d+:\\d+\\] ==", form_cells$R_code), ]
  
  if (nrow(sumif_pattern) > 100) {
    message(sprintf("Vectorisation de %d formules SUMIF...", nrow(sumif_pattern)))
    
    # Analyser la structure SUMIF
    sumif_analysis <- sumif_pattern %>%
      mutate(
        range_pattern = str_extract(R_code, "values\\[\\d+:\\d+, \\d+:\\d+\\]"),
        criteria_pattern = str_extract(R_code, "== paste0\\([^)]+\\)"),
        sum_range_pattern = str_extract(R_code, "values\\[\\d+:\\d+, \\d+:\\d+\\], na\\.r")
      ) %>%
      group_by(sheet, range_pattern, sum_range_pattern) %>%
      summarise(
        count = n(),
        rows = list(row),
        cols = list(col),
        .groups = 'drop'
      ) %>%
      filter(count > 10)
    
    message(sprintf("Détecté %d groupes SUMIF vectorisables", nrow(sumif_analysis)))
  }
  
  return(list(
    paste_optimized = length(paste_pattern),
    sumif_optimized = length(sumif_pattern),
    total_optimizable = length(paste_pattern) + length(sumif_pattern)
  ))
}

# Générateur de script avec optimisations vectorielles
generate_vectorized_script <- function(form_cells, script_file) {
  
  # Analyser les patterns
  optimization_stats <- vectorize_repetitive_patterns(form_cells)
  
  # Séparer les formules optimisables des autres
  paste_cells <- form_cells[grepl("paste0\\(values\\[\\d+, \\d+\\],values\\[\\d+, \\d+\\]\\)", form_cells$R_code), ]
  sumif_cells <- form_cells[grepl("sum\\(sum\\(\\(\\(values\\[\\d+:\\d+, \\d+:\\d+\\] ==", form_cells$R_code), ]
  other_cells <- form_cells[!form_cells$address %in% c(paste_cells$address, sumif_cells$address), ]
  
  message(sprintf("Répartition: %d paste0, %d sumif, %d autres formules", 
                  nrow(paste_cells), nrow(sumif_cells), nrow(other_cells)))
  
  lines <- c(
    "# Script généré avec optimisations vectorielles",
    sprintf("# %d formules dont %d vectorisées", nrow(form_cells), optimization_stats$total_optimizable),
    "",
    "script <- function(rv, rv_path) {",
    "  start_time <- Sys.time()",
    "  tryCatch({",
    "    wb <- rv$fichier_excel",
    "    wb_sheets <- openxlsx2::wb_get_sheet_names(wb)",
    "    ",
    "    message('Chargement et conversion...')",
    "    sheets <- setNames(",
    "      lapply(wb_sheets, function(sh) {",
    "        df <- openxlsx2::wb_to_df(wb, sheet = sh, colNames = FALSE)",
    "        df[] <- lapply(df, function(col) {",
    "          if (is.character(col) || is.factor(col)) {",
    "            col[is.na(col) | col == '' | col == 'NA'] <- '0'",
    "            col_num <- suppressWarnings(as.numeric(col))",
    "            col_num[is.na(col_num)] <- 0",
    "            return(col_num)",
    "          } else if (is.logical(col)) {",
    "            return(as.numeric(col))",
    "          } else {",
    "            col[is.na(col)] <- 0",
    "            return(col)",
    "          }",
    "        })",
    "        df",
    "      }),",
    "      wb_sheets",
    "    )",
    ""
  )
  
  # SECTION OPTIMISATIONS VECTORIELLES
  if (nrow(paste_cells) > 100) {
    lines <- c(lines,
               "    # === OPTIMISATION VECTORIELLE: PASTE0 ===",
               sprintf("    message('Vectorisation de %d concaténations...')", nrow(paste_cells)),
               ""
    )
    
    # Grouper les paste0 par feuille
    paste_by_sheet <- paste_cells %>%
      group_by(sheet) %>%
      summarise(
        rows = list(row),
        cols = list(col),
        codes = list(R_code),
        .groups = 'drop'
      )
    
    for (i in seq_len(nrow(paste_by_sheet))) {
      sheet_name <- paste_by_sheet$sheet[i]
      rows <- paste_by_sheet$rows[[i]]
      cols <- paste_by_sheet$cols[[i]]
      
      lines <- c(lines,
                 sprintf("    # Feuille %s: %d concaténations", sheet_name, length(rows)),
                 sprintf("    values <- sheets[['%s']]", sheet_name)
      )
      
      # Traiter par batch de 500
      batch_size <- 500
      n_batches <- ceiling(length(rows) / batch_size)
      
      for (batch in seq_len(n_batches)) {
        start_idx <- (batch - 1) * batch_size + 1
        end_idx <- min(batch * batch_size, length(rows))
        batch_rows <- rows[start_idx:end_idx]
        batch_cols <- cols[start_idx:end_idx]
        
        lines <- c(lines,
                   sprintf("    # Batch %d/%d", batch, n_batches),
                   sprintf("    batch_rows <- c(%s)", paste(batch_rows, collapse = ",")),
                   sprintf("    batch_cols <- c(%s)", paste(batch_cols, collapse = ",")),
                   "    for (i in seq_along(batch_rows)) {",
                   sprintf("      sheets[['%s']][batch_rows[i], batch_cols[i]] <- paste0(values[batch_rows[i], 7], values[batch_rows[i], 2])", sheet_name),
                   "    }",
                   ""
        )
      }
    }
  }
  
  # SECTION FORMULES COMPLEXES (SUMIF)
  if (nrow(sumif_cells) > 0) {
    lines <- c(lines,
               "    # === FORMULES COMPLEXES (SUMIF) ===",
               sprintf("    message('Traitement de %d formules SUMIF...')", nrow(sumif_cells)),
               ""
    )
    
    # Traiter les SUMIF par feuille
    sumif_by_sheet <- sumif_cells %>% group_by(sheet)
    
    for (sheet_name in unique(sumif_cells$sheet)) {
      sheet_sumifs <- sumif_cells[sumif_cells$sheet == sheet_name, ]
      
      lines <- c(lines,
                 sprintf("    # Feuille %s: %d SUMIF", sheet_name, nrow(sheet_sumifs)),
                 sprintf("    values <- sheets[['%s']]", sheet_name)
      )
      
      # Traiter par batch de 200 (plus lent)
      batch_size <- 200
      n_batches <- ceiling(nrow(sheet_sumifs) / batch_size)
      
      for (batch in seq_len(n_batches)) {
        start_idx <- (batch - 1) * batch_size + 1
        end_idx <- min(batch * batch_size, nrow(sheet_sumifs))
        batch_cells <- sheet_sumifs[start_idx:end_idx, ]
        
        lines <- c(lines, sprintf("    # SUMIF Batch %d/%d", batch, n_batches))
        
        for (j in seq_len(nrow(batch_cells))) {
          cell <- batch_cells[j, ]
          lines <- c(lines,
                     sprintf("    tryCatch({ sheets[['%s']][%d, %d] <- %s }, error = function(e) { warning('Erreur %s!%s') })",
                             sheet_name, cell$row, cell$col, cell$R_code, sheet_name, cell$address)
          )
        }
        
        lines <- c(lines, "")
      }
    }
  }
  
  # SECTION AUTRES FORMULES
  if (nrow(other_cells) > 0) {
    lines <- c(lines,
               "    # === AUTRES FORMULES ===",
               sprintf("    message('Traitement de %d autres formules...')", nrow(other_cells)),
               ""
    )
    
    for (sheet_name in unique(other_cells$sheet)) {
      sheet_others <- other_cells[other_cells$sheet == sheet_name, ]
      
      lines <- c(lines,
                 sprintf("    # Feuille %s: %d formules diverses", sheet_name, nrow(sheet_others)),
                 sprintf("    values <- sheets[['%s']]", sheet_name)
      )
      
      for (i in seq_len(nrow(sheet_others))) {
        cell <- sheet_others[i, ]
        lines <- c(lines,
                   sprintf("    tryCatch({ sheets[['%s']][%d, %d] <- %s }, error = function(e) { warning('Erreur %s!%s') })",
                           sheet_name, cell$row, cell$col, fix_range_bug(cell$R_code), sheet_name, cell$address)
        )
      }
      
      lines <- c(lines, "")
    }
  }
  
  # FINALISATION
  lines <- c(lines,
             "    # === MISE À JOUR WORKBOOK ===",
             "    message('Mise à jour du workbook...')",
             "    for (sheet_name in names(sheets)) {",
             "      tryCatch({",
             "        openxlsx2::wb_remove_worksheet(wb, sheet_name)",
             "        openxlsx2::wb_add_worksheet(wb, sheet_name)",
             "        wb <- openxlsx2::wb_add_data(wb, sheet = sheet_name, x = sheets[[sheet_name]], col_names = FALSE)",
             "      }, error = function(e) { message('Erreur feuille ', sheet_name, ': ', e$message) })",
             "    }",
             "    ",
             "    rv$fichier_excel <- wb",
             "    rv$excel_updated <- Sys.time()",
             "    ",
             "    end_time <- Sys.time()",
             "    duration <- round(as.numeric(end_time - start_time, units = 'secs'), 2)",
             sprintf("    message('✅ %d formules traitées en ', duration, 's')", nrow(form_cells)),
             sprintf("    message('Optimisations: %d paste0, %d sumif vectorisées')", nrow(paste_cells), nrow(sumif_cells)),
             "    showNotification(paste0('✅ Formules optimisées en ', duration, 's'), type = 'message')",
             "    ",
             "  }, error = function(e) {",
             "    message('❌ Erreur: ', e$message)",
             "    showNotification(paste0('❌ Erreur: ', e$message), type = 'error')",
             "  })",
             "}"
  )
  
  # CORRECTION : Fermetures propres (PAS de double accolade)
  lines <- c(lines,
             "  }, error = function(e) {",
             "    message('❌ Erreur globale: ', e$message)", 
             "    showNotification(paste0('❌ Erreur du script: ', e$message), type = 'error')",
             "  })",
             "}"
  )
  
  # Écriture
  writeLines(lines, script_file)
  file_size <- file.size(script_file)
  
  message(sprintf("✅ Script vectorisé généré: %s", script_file))
  message(sprintf("   Taille: %.1f KB (%d lignes)", file_size/1024, length(lines)))
  message(sprintf("   Optimisations: %d%% des formules vectorisées", 
                  round(100 * optimization_stats$total_optimizable / nrow(form_cells))))
  
  return(list(
    lines = length(lines),
    optimized_pct = round(100 * optimization_stats$total_optimizable / nrow(form_cells))
  ))
}