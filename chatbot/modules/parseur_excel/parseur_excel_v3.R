# parse_excel_formulas_with_assignments.R
# Version optimisée pour gros fichiers

library(tidyxl)
library(openxlsx)
library(dplyr)
library(stringr)

# Détermination du répertoire du script pour sourcer les dépendances
script_dir <- dirname(normalizePath(sys.frame(1)$ofile, mustWork = FALSE))
source(file.path(script_dir, "modules/parseur_excel/excel_formula_to_R.R"))
source(file.path(script_dir, "modules/parseur_excel/utils.R"))
source(file.path(script_dir, "modules/parseur_excel/utils_gros_fichier.R"))

# Configuration adaptative selon la taille du fichier -------------------------
analyze_excel_complexity <- function(path) {
  file_size_mb <- file.size(path) / 1e6
  wb_sheets <- getSheetNames(path)
  n_sheets <- length(wb_sheets)
  
  # Échantillonnage rapide
  tryCatch({
    sample_cells <- xlsx_cells(path, sheets = wb_sheets[1]) %>%
      filter(!is.na(formula)) %>%
      head(100)
    
    avg_complexity <- mean(nchar(sample_cells$formula), na.rm = TRUE)
    estimated_formulas <- nrow(sample_cells) * n_sheets * 10
  }, error = function(e) {
    avg_complexity <- 50
    estimated_formulas <- n_sheets * 1000
  })
  
  list(
    size_mb = file_size_mb,
    n_sheets = n_sheets,
    estimated_formulas = estimated_formulas,
    complexity = avg_complexity,
    config = list(
      sheet_batch_size = if(file_size_mb > 100) 3 else if(file_size_mb > 50) 5 else 10,
      formula_chunk_size = if(avg_complexity > 100) 500 else if(avg_complexity > 50) 1000 else 2000,
      use_parallel = file_size_mb > 20 || n_sheets > 10,
      use_cache = avg_complexity > 30,
      streaming_threshold = 10000
    )
  )
}

# Cache global pour les références converties ---------------------------------
ref_cache_env <- new.env(hash = TRUE, parent = emptyenv())

convert_ref_cached <- function(ref) {
  if (exists(ref, envir = ref_cache_env)) {
    return(get(ref, envir = ref_cache_env))
  }
  result <- convert_ref(ref)
  assign(ref, result, envir = ref_cache_env)
  result
}

clear_ref_cache <- function() {
  rm(list = ls(envir = ref_cache_env), envir = ref_cache_env)
  gc()
}

# Chargement par batch des feuilles -------------------------------------------
load_sheets_in_batches <- function(path, batch_size = 5) {
  wb_sheets <- getSheetNames(path)
  n_batches <- ceiling(length(wb_sheets) / batch_size)
  
  message(sprintf("[load_sheets] Chargement de %d feuilles en %d batchs...", 
                  length(wb_sheets), n_batches))
  
  sheets <- list()
  
  for (batch in seq_len(n_batches)) {
    start_idx <- (batch - 1) * batch_size + 1
    end_idx <- min(batch * batch_size, length(wb_sheets))
    batch_sheets <- wb_sheets[start_idx:end_idx]
    
    message(sprintf("[load_sheets] Batch %d/%d...", batch, n_batches))
    
    batch_data <- setNames(
      lapply(batch_sheets, function(sh) {
        message(sprintf(" - Lecture de '%s'", sh))
        df <- read.xlsx(path, sheet = sh, colNames = FALSE)
        
        # Conversion intelligente pour économiser la mémoire
        df[] <- lapply(seq_along(df), function(col_idx) {
          col <- df[[col_idx]]
          
          # Si la colonne contient des formules, elle sera numérique
          # Analyser rapidement le contenu
          sample_values <- head(col[!is.na(col) & col != ''], 20)
          
          if (length(sample_values) == 0) return(col)
          
          # Test rapide de numericité
          numeric_test <- suppressWarnings(as.numeric(sample_values))
          if (sum(!is.na(numeric_test)) / length(sample_values) > 0.8) {
            # Colonne principalement numérique
            result <- suppressWarnings(as.numeric(col))
            result[is.na(result)] <- 0
            return(result)
          } else {
            # Colonne texte - préserver le contenu
            col[is.na(col)] <- ''
            return(col)
          }
        })
        df
      }), 
      batch_sheets
    )
    
    sheets <- c(sheets, batch_data)
    
    # Forcer le garbage collection après chaque batch
    if (batch %% 2 == 0) gc()
  }
  
  sheets
}

# Conversion parallèle des formules (si disponible) --------------------------
convert_formulas_batch <- function(formulas, nom_cellules, use_parallel = FALSE) {
  n_formulas <- length(formulas)
  
  if (!use_parallel || n_formulas < 1000) {
    # Conversion séquentielle avec cache
    return(vapply(formulas, function(f) {
      # Remplacer convert_ref par convert_ref_cached dans la formule
      convert_formula(f, nom_cellules)
    }, character(1)))
  }
  
  # Tentative de parallélisation
  tryCatch({
    library(parallel)
    n_cores <- min(4, detectCores() - 1)
    
    if (n_cores > 1) {
      message(sprintf("[convert_formulas] Utilisation de %d cœurs...", n_cores))
      
      cl <- makeCluster(n_cores)
      
      # Exporter les fonctions nécessaires
      clusterExport(cl, c("convert_formula", "split_at_top", "split_top", 
                          "clean_parentheses", "col2num", "convert_ref", 
                          "convert_criteria", "fun_map", "convert_ref_cached",
                          "ref_cache_env"), envir = environment())
      
      # Diviser le travail
      chunks <- split(formulas, cut(seq_along(formulas), n_cores, labels = FALSE))
      
      results <- parLapply(cl, chunks, function(chunk) {
        vapply(chunk, function(f) convert_formula(f, nom_cellules), character(1))
      })
      
      stopCluster(cl)
      
      return(unlist(results))
    }
  }, error = function(e) {
    message("[convert_formulas] Parallélisation impossible, retour au séquentiel")
  })
  
  # Fallback séquentiel
  vapply(formulas, function(f) convert_formula(f, nom_cellules), character(1))
}

# Fonction pour corriger le bug des plages
fix_range_bug <- function(r_code) {
  gsub("(\\d+):(\\1)", "\\1", r_code)
}

# Génération streaming du script pour gros fichiers ---------------------------
generate_script_streaming <- function(form_cells, script_file, chunk_size = 1000) {
  con <- file(script_file, "w")
  
  # En-tête du script
  header <- c(
    "# Script généré par parse_excel_formulas() - Version optimisée",
    "# Génération par streaming pour gros fichiers",
    "script <- function(rv, rv_path) {",
    "  start_time <- Sys.time()",
    "  tryCatch({",
    "    path <- rv_path()",
    "    wb <- rv$fichier_excel",
    "    if (is.null(wb) || !inherits(wb, \"wbWorkbook\")) {",
    "      showNotification(\"❌ Le workbook Excel n'est pas valide.\", type = \"error\")",
    "      return()",
    "    }",
    "    wb_sheets <- openxlsx2::wb_get_sheet_names(wb)",
    "    ",
    "    # Chargement et conversion INTELLIGENTE des données",
    "    message('Chargement des feuilles...')",
    "    sheets <- setNames(",
    "      lapply(wb_sheets, function(sh) {",
    "        df <- openxlsx2::wb_to_df(wb, sheet = sh, colNames = FALSE)",
    "        # Conversion intelligente : préserver le type original si possible",
    "        df[] <- lapply(seq_along(df), function(col_idx) {",
    "          col <- df[[col_idx]]",
    "          # Analyser le contenu de la colonne pour déterminer le type",
    "          non_na_values <- col[!is.na(col) & col != '']",
    "          ",
    "          # Si la colonne est vide, retourner tel quel",
    "          if (length(non_na_values) == 0) return(col)",
    "          ",
    "          # Tester si c'est numérique",
    "          numeric_test <- suppressWarnings(as.numeric(non_na_values))",
    "          numeric_ratio <- sum(!is.na(numeric_test)) / length(non_na_values)",
    "          ",
    "          # Si > 80% des valeurs sont numériques, convertir en numérique",
    "          if (numeric_ratio > 0.8) {",
    "            result <- suppressWarnings(as.numeric(col))",
    "            # Remplacer NA par 0 seulement pour les colonnes numériques",
    "            result[is.na(result)] <- 0",
    "            return(result)",
    "          } else {",
    "            # Sinon, garder comme caractère mais nettoyer les NA",
    "            col[is.na(col)] <- ''",
    "            return(col)",
    "          }",
    "        })",
    "        df",
    "      }),",
    "      wb_sheets",
    "    )",
    "",
    "    # Définition des helpers",
    "    vlookup_r <- function(lookup, df, col_index) {",
    "      idx <- match(lookup, df[[1]])",
    "      ifelse(is.na(idx), NA, df[[col_index]][idx])",
    "    }",
    "",
    "    # Application des formules",
    "    message('Application des formules...')",
    "    total_formulas <- " %+% nrow(form_cells),
    "    formulas_done <- 0",
    ""
  )
  
  writeLines(header, con)
  
  # Traiter les formules par chunks
  n_chunks <- ceiling(nrow(form_cells) / chunk_size)
  
  for (chunk_idx in seq_len(n_chunks)) {
    start_idx <- (chunk_idx - 1) * chunk_size + 1
    end_idx <- min(chunk_idx * chunk_size, nrow(form_cells))
    chunk_cells <- form_cells[start_idx:end_idx, ]
    
    # Message de progression
    chunk_lines <- c(
      sprintf("    # Chunk %d/%d (%d formules)", chunk_idx, n_chunks, nrow(chunk_cells)),
      sprintf("    if (formulas_done %% 1000 == 0) message(sprintf('Progression: %%d/%%d (%%d%%%%)', formulas_done, total_formulas, round(100*formulas_done/total_formulas)))")
    )
    
    # Grouper par feuille pour optimiser
    for (sheet in unique(chunk_cells$sheet)) {
      sheet_cells <- chunk_cells[chunk_cells$sheet == sheet, ]
      
      chunk_lines <- c(chunk_lines,
                       sprintf("    # Feuille '%s' - %d formules", sheet, nrow(sheet_cells)),
                       sprintf("    values <- sheets[[\"%s\"]]", sheet)
      )
      
      for (i in seq_len(nrow(sheet_cells))) {
        cell <- sheet_cells[i, ]
        code_fixed <- fix_range_bug(cell$R_code)
        
        chunk_lines <- c(chunk_lines,
                         sprintf("    # %s -> %s", cell$address, cell$formula),
                         sprintf("    tryCatch({ sheets[[\"%s\"]][%d, %d] <- %s }, error = function(e) { warning('Erreur %s!%s: ', e$message) })",
                                 sheet, cell$row, cell$col, code_fixed, sheet, cell$address),
                         "    formulas_done <- formulas_done + 1"
        )
      }
      
      chunk_lines <- c(chunk_lines, "")
    }
    
    writeLines(chunk_lines, con)
    
    if (chunk_idx %% 5 == 0) {
      message(sprintf("[generate_script] Écriture: %d%%", round(100 * chunk_idx / n_chunks)))
    }
  }
  
  # Footer du script
  footer <- c(
    "    # Mise à jour du workbook",
    "    message('Mise à jour du workbook...')",
    "    for (sheet_name in names(sheets)) {",
    "      tryCatch({",
    "        sheet_data <- sheets[[sheet_name]]",
    "        for (row in seq_len(nrow(sheet_data))) {",
    "          for (col in seq_len(ncol(sheet_data))) {",
    "            wb <- openxlsx2::wb_add_data(wb, sheet = sheet_name, ",
    "                                          x = sheet_data[row, col], ",
    "                                          start_col = col, start_row = row)",
    "          }",
    "        }",
    "      }, error = function(e) {",
    "        warning(sprintf('Erreur mise à jour %s: %s', sheet_name, e$message))",
    "      })",
    "    }",
    "",
    "    rv$fichier_excel <- wb",
    "    rv$excel_updated <- Sys.time()",
    "    ",
    "    end_time <- Sys.time()",
    "    duration <- round(as.numeric(end_time - start_time, units = 'secs'), 2)",
    "    message(sprintf('✅ %d formules traitées en %.2f secondes', total_formulas, duration))",
    "    showNotification(sprintf('✅ Formules appliquées en %.1fs !', duration), type = 'message')",
    "",
    "  }, error = function(e) {",
    "    showNotification(paste0('❌ Erreur: ', e$message), type = 'error')",
    "    message('Erreur: ', e$message)",
    "  })",
    "}"
  )
  
  writeLines(footer, con)
  close(con)
  
  file_size_kb <- file.size(script_file) / 1024
  message(sprintf("[generate_script] Script généré: %s (%.1f KB)", script_file, file_size_kb))
}

# Fonction principale optimisée -----------------------------------------------
parse_excel_formulas <- function(path, emit_script = FALSE) {
  message("\n[parse_excel_formulas] Analyse du fichier...")
  
  # Analyse de complexité
  complexity <- analyze_excel_complexity(path)
  config <- complexity$config
  
  message(sprintf("[parse_excel_formulas] Fichier: %.1f MB, %d feuilles", 
                  complexity$size_mb, complexity$n_sheets))
  message(sprintf("[parse_excel_formulas] Configuration: batch_size=%d, use_cache=%s, use_parallel=%s",
                  config$sheet_batch_size, config$use_cache, config$use_parallel))
  
  # Activation du cache si recommandé
  if (config$use_cache) {
    message("[parse_excel_formulas] Cache des références activé")
    # Remplacer temporairement convert_ref
    convert_ref_original <- convert_ref
    assignInNamespace("convert_ref", convert_ref_cached, ns = environment())
  }
  
  # 1. Chargement optimisé des feuilles
  sheets <- load_sheets_in_batches(path, config$sheet_batch_size)
  
  # 2. Extraction des variables globales
  message("[parse_excel_formulas] Extraction des variables globales")
  globals <- get_excel_globals(path)
  nom_cellules <- build_named_cell_map(globals)
  
  # 3. Extraction des formules
  message("[parse_excel_formulas] Extraction des cellules contenant des formules...")
  cells_all <- xlsx_cells(path)
  form_cells <- cells_all %>% 
    filter(!is.na(formula)) %>%
    mutate(
      formula = gsub("[\r\n]+", " ", formula),
      formula = gsub("\\s+", " ", formula),
      formula = trimws(formula)
    )
  
  message(sprintf("[parse_excel_formulas] %d formules extraites.", nrow(form_cells)))
  
  # 4. Filtrage des erreurs Excel
  initial_count <- nrow(form_cells)
  form_cells <- form_cells %>%
    filter(
      !grepl("\\bLEFT\\b", formula, ignore.case = TRUE),
      !grepl("#REF!", formula, fixed = TRUE),
      !grepl("#N/A", formula, fixed = TRUE),
      !grepl("#VALUE!", formula, fixed = TRUE),
      !grepl("#DIV/0!", formula, fixed = TRUE),
      !grepl("#NAME\\?", formula),
      !grepl("#NULL!", formula, fixed = TRUE),
      !grepl("#NUM!", formula, fixed = TRUE)
    )
  
  if (initial_count > nrow(form_cells)) {
    message(sprintf("[parse_excel_formulas] %d formules filtrées (erreurs Excel)", 
                    initial_count - nrow(form_cells)))
  }
  
  # 5. Extraction des dépendances
  message("[parse_excel_formulas] Extraction des dépendances...")
  form_cells <- form_cells %>%
    mutate(deps = purrr::pmap(list(formula, sheet), extract_deps))
  
  # 6. Conversion des formules (potentiellement en parallèle)
  message("[parse_excel_formulas] Conversion des formules Excel en code R...")
  
  if (nrow(form_cells) > config$formula_chunk_size * 2) {
    # Traiter par chunks pour les gros volumes
    n_chunks <- ceiling(nrow(form_cells) / config$formula_chunk_size)
    r_codes <- character(nrow(form_cells))
    
    for (chunk in seq_len(n_chunks)) {
      start_idx <- (chunk - 1) * config$formula_chunk_size + 1
      end_idx <- min(chunk * config$formula_chunk_size, nrow(form_cells))
      
      if (chunk %% 5 == 1) {
        message(sprintf("[parse_excel_formulas] Conversion: %d%%", 
                        round(100 * chunk / n_chunks)))
      }
      
      chunk_formulas <- form_cells$formula[start_idx:end_idx]
      r_codes[start_idx:end_idx] <- convert_formulas_batch(
        chunk_formulas, 
        nom_cellules, 
        config$use_parallel
      )
    }
    
    form_cells$R_code <- r_codes
  } else {
    # Conversion normale pour petits volumes
    form_cells$R_code <- convert_formulas_batch(
      form_cells$formula, 
      nom_cellules, 
      config$use_parallel
    )
  }
  
  # Ajout des colonnes row/col
  form_cells <- form_cells %>%
    mutate(
      row = as.integer(sub("^[A-Z]+", "", address)),
      col = vapply(gsub("[0-9]+", "", address), function(x) as.integer(col2num(x)), integer(1))
    ) %>%
    select(sheet, address, row, col, formula, R_code)
  
  message("[parse_excel_formulas] Conversion terminée.")
  
  # 7. Évaluation ou génération de script
  if (!emit_script) {
    message("[parse_excel_formulas] Evaluation des cellules dans le bon ordre...")
    sheets <- evaluate_cells(sheets, form_cells)
  } else {
    message("[parse_excel_formulas] Génération du script...")
    
    script_file <- paste0(tools::file_path_sans_ext(basename(path)), "_converted_formulas.R")
    
    # Utiliser la génération streaming pour gros fichiers
    if (nrow(form_cells) > config$streaming_threshold) {
      message("[parse_excel_formulas] Mode streaming activé (gros fichier)")
      generate_script_streaming(form_cells, script_file, config$formula_chunk_size)
    } else {
      # Utiliser la version vectorisée existante
      generate_vectorized_script(form_cells, script_file)
    }
  }
  
  # 8. Nettoyage
  if (config$use_cache) {
    clear_ref_cache()
    # Restaurer la fonction originale si modifiée
    if (exists("convert_ref_original")) {
      assignInNamespace("convert_ref", convert_ref_original, ns = environment())
    }
  }
  
  gc()
  
  message("[parse_excel_formulas] Traitement terminé.\n")
  
  # Retourner les sheets pour compatibilité
  if (!emit_script) {
    return(sheets)
  }
}