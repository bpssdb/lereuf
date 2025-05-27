# parseur_excel/parseur_excel_v3.R
# Parseur Excel optimisé pour gros fichiers - Version BudgiBot

library(tidyxl)
library(openxlsx2)
library(dplyr)
library(stringr)
library(future)
library(future.apply)
library(digest)

source("~/work/lereuf/chatbot/modules/parseur_excel/fix_parseur_v3.R")


# ===== CONFIGURATION =====
get_parser_config <- function() {
  list(
    chunk_size = getOption("budgibot.chunk_size", 1000),
    max_memory_mb = getOption("budgibot.max_memory", 512),
    parallel_workers = getOption("budgibot.workers", min(4, parallel::detectCores() - 1)),
    cache_enabled = getOption("budgibot.cache", TRUE),
    progress_enabled = getOption("budgibot.progress", TRUE),
    temp_dir = file.path(tempdir(), "budgibot_cache")
  )
}

# ===== GESTION MÉMOIRE =====
monitor_memory <- function(config) {
  gc_info <- gc(verbose = FALSE)
  used_mb <- sum(gc_info[, 2]) * 8 / 1024
  
  if (used_mb > config$max_memory_mb) {
    message(sprintf("⚠️ Mémoire élevée: %.1f MB, nettoyage...", used_mb))
    gc(full = TRUE)
    return(TRUE)
  }
  return(FALSE)
}

# ===== CACHE INTELLIGENT =====
generate_file_signature <- function(file_path) {
  if (!file.exists(file_path)) return(NULL)
  
  file_info <- file.info(file_path)
  signature_data <- list(
    path = normalizePath(file_path),
    size = file_info$size,
    mtime = as.numeric(file_info$mtime)
  )
  
  return(digest::digest(signature_data, algo = "md5"))
}

get_cache_file <- function(signature, config) {
  if (!config$cache_enabled) return(NULL)
  
  cache_dir <- config$temp_dir
  if (!dir.exists(cache_dir)) {
    dir.create(cache_dir, recursive = TRUE, showWarnings = FALSE)
  }
  
  return(file.path(cache_dir, paste0("parsed_", signature, ".rds")))
}

load_cached_result <- function(signature, config) {
  cache_file <- get_cache_file(signature, config)
  if (is.null(cache_file) || !file.exists(cache_file)) return(NULL)
  
  tryCatch({
    result <- readRDS(cache_file)
    message("📦 Données chargées depuis le cache")
    return(result)
  }, error = function(e) {
    message("⚠️ Erreur lecture cache, reparser nécessaire")
    return(NULL)
  })
}

save_cached_result <- function(result, signature, config) {
  if (!config$cache_enabled) return()
  
  cache_file <- get_cache_file(signature, config)
  if (is.null(cache_file)) return()
  
  tryCatch({
    saveRDS(result, cache_file, compress = TRUE)
    message("💾 Résultat sauvé en cache")
  }, error = function(e) {
    warning("Impossible de sauver en cache: ", e$message)
  })
}

# ===== FONCTION UTILITAIRE =====
# Fonction pour convertir colonne Excel en numéro (ex: "A"->1, "AB"->28)
excel_col_to_num_v3 <- function(col_str) {
  if (is.na(col_str) || !nzchar(col_str)) return(1)
  chars <- strsplit(col_str, "")[[1]]
  sum(sapply(seq_along(chars), function(i) {
    (match(chars[i], LETTERS)) * 26^(length(chars) - i)
  }))
}

# ===== EXTRACTION OPTIMISÉE =====
extract_formulas_optimized <- function(path, config) {
  message("🔍 Extraction des formules...")
  
  # Lecture avec gestion d'erreur
  all_cells <- tryCatch({
    tidyxl::xlsx_cells(path)
  }, error = function(e) {
    stop("Impossible de lire le fichier Excel: ", e$message)
  })
  
  # Filtrage et nettoyage des formules
  formula_cells <- all_cells %>%
    filter(!is.na(formula)) %>%
    filter(
      # Filtrer les erreurs Excel connues
      !grepl("#REF!|#N/A|#VALUE!|#DIV/0!|#NAME\\?|#NULL!|#NUM!", formula, ignore.case = TRUE),
      # Filtrer les fonctions non supportées (temporairement)
      !grepl("\\bLEFT\\b|\\bRIGHT\\b", formula, ignore.case = TRUE)
    ) %>%
    mutate(
      # Nettoyage des formules
      formula = str_trim(str_replace_all(formula, "[\r\n]+", " ")),
      formula = str_replace_all(formula, "\\s+", " "),
      # Ajout d'identifiants avec fonction locale
      row_coord = as.integer(str_extract(address, "\\d+")),
      col_coord = sapply(str_extract(address, "^[A-Z]+"), excel_col_to_num_v3),
      chunk_id = ceiling(row_number() / config$chunk_size)
    ) %>%
    filter(!is.na(row_coord), !is.na(col_coord))
  
  if (nrow(formula_cells) == 0) {
    message("ℹ️ Aucune formule exploitable trouvée")
    return(data.frame())
  }
  
  message(sprintf("📊 %d formules trouvées, %d chunks à traiter", 
                  nrow(formula_cells), max(formula_cells$chunk_id)))
  
  return(formula_cells)
}

# ===== TRAITEMENT PAR CHUNKS =====
process_single_chunk <- function(chunk_data, noms_cellules, config) {
  chunk_id <- unique(chunk_data$chunk_id)[1]
  
  tryCatch({
    # Vérifier que convert_formula est disponible
    if (!exists("convert_formula")) {
      stop("Fonction convert_formula non trouvée. Vérifiez le chargement des dépendances.")
    }
    
    # Conversion des formules avec la fonction existante
    chunk_results <- chunk_data %>%
      rowwise() %>%
      mutate(
        R_code = tryCatch({
          convert_formula(formula, noms_cellules)
        }, error = function(e) {
          # Log de l'erreur pour debug
          message(sprintf("Erreur conversion formule %s: %s", address, e$message))
          paste0("# Erreur: ", e$message)
        }),
        conversion_success = !is.na(R_code) && 
          !str_detect(R_code, "^# Erreur:") &&
          nzchar(str_trim(R_code)) &&
          R_code != "NA"
      ) %>%
      ungroup()
    
    # Séparation succès/erreurs
    successful <- chunk_results %>% filter(conversion_success)
    errors <- chunk_results %>% filter(!conversion_success)
    
    # Nettoyage mémoire
    monitor_memory(config)
    
    return(list(
      processed = successful,
      errors = errors,
      chunk_id = chunk_id,
      success_count = nrow(successful),
      error_count = nrow(errors)
    ))
    
  }, error = function(e) {
    message(sprintf("❌ Erreur chunk %d: %s", chunk_id, e$message))
    return(list(
      processed = data.frame(),
      errors = chunk_data,
      chunk_id = chunk_id,
      success_count = 0,
      error_count = nrow(chunk_data)
    ))
  })
}

# ===== FONCTION PRINCIPALE =====
parse_excel_formulas_v3 <- function(path, emit_script = TRUE, force_reparse = FALSE) {
  start_time <- Sys.time()
  config <- get_parser_config()
  
  message("🚀 BudgiBot Parseur Excel v3 - Démarrage")
  message(sprintf("📁 Fichier: %s", basename(path)))
  
  # Vérifications
  if (!file.exists(path)) {
    stop("Fichier introuvable: ", path)
  }
  
  # Vérification des dépendances critiques
  if (!exists("convert_formula")) {
    message("⚠️ Fonction convert_formula non trouvée, tentative de chargement...")
    tryCatch({
      script_dir <- dirname(normalizePath(sys.frame(1)$ofile, mustWork = FALSE))
      source(file.path(script_dir, "excel_formula_to_R2.R"))
      source(file.path(script_dir, "utils.R"))
      message("✅ Dépendances rechargées")
    }, error = function(e) {
      # Essayer le chemin alternatif
      tryCatch({
        source("~/work/lereuf/chatbot/modules/parseur_excel/excel_formula_to_R2.R")
        source("~/work/lereuf/chatbot/modules/parseur_excel/utils.R")
        message("✅ Dépendances chargées depuis chemin alternatif")
      }, error = function(e2) {
        stop("Impossible de charger les fonctions de conversion Excel: ", e2$message)
      })
    })
  }
  
  # Gestion du cache
  file_signature <- generate_file_signature(path)
  if (!force_reparse && !is.null(file_signature)) {
    cached <- load_cached_result(file_signature, config)
    if (!is.null(cached)) {
      return(cached)
    }
  }
  
  # Configuration parallèle
  if (config$parallel_workers > 1) {
    future::plan(future::multisession, workers = config$parallel_workers)
    message(sprintf("⚡ Mode parallèle: %d workers", config$parallel_workers))
  } else {
    future::plan(future::sequential)
    message("🔄 Mode séquentiel")
  }
  
  # Chargement des données de base (compatibilité avec l'existant)
  wb_sheets <- tryCatch({
    openxlsx2::wb_load(path) %>% openxlsx2::wb_get_sheet_names()
  }, error = function(e) {
    stop("Impossible de charger le workbook: ", e$message)
  })
  
  sheets <- setNames(
    lapply(wb_sheets, function(sh) {
      message(sprintf("📊 Lecture feuille '%s'", sh))
      tryCatch({
        openxlsx2::wb_load(path) %>% openxlsx2::wb_to_df(sheet = sh, col_names = FALSE)
      }, error = function(e) {
        message(sprintf("⚠️ Erreur lecture feuille %s: %s", sh, e$message))
        data.frame()
      })
    }), wb_sheets
  )
  
  # Extraction des variables globales (compatibilité)
  globals <- tryCatch({
    get_excel_globals(path)
  }, error = function(e) {
    message("⚠️ Impossible d'extraire les variables globales - continuons sans")
    data.frame()
  })
  
  noms_cellules <- if (nrow(globals) > 0) {
    build_named_cell_map(globals)
  } else {
    list()
  }
  
  # Extraction des formules
  formula_cells <- extract_formulas_optimized(path, config)
  
  if (nrow(formula_cells) == 0) {
    result <- list(
      formulas = data.frame(),
      sheets = sheets,
      globals = globals,
      processing_time = as.numeric(difftime(Sys.time(), start_time, units = "secs")),
      statistics = list(total = 0, success = 0, errors = 0, success_rate = 0)
    )
    return(result)
  }
  
  # Division en chunks
  chunks <- split(formula_cells, formula_cells$chunk_id)
  n_chunks <- length(chunks)
  
  message(sprintf("⚙️ Traitement de %d chunks...", n_chunks))
  
  # Barre de progression
  if (config$progress_enabled) {
    pb <- txtProgressBar(min = 0, max = n_chunks, style = 3)
    on.exit(close(pb), add = TRUE)
  }
  
  # Traitement des chunks
  chunk_results <- if (config$parallel_workers > 1 && n_chunks > 1) {
    # Mode parallèle
    future_lapply(seq_along(chunks), function(i) {
      result <- process_single_chunk(chunks[[i]], noms_cellules, config)
      if (config$progress_enabled) setTxtProgressBar(pb, i)
      return(result)
    }, future.seed = TRUE)
  } else {
    # Mode séquentiel
    lapply(seq_along(chunks), function(i) {
      result <- process_single_chunk(chunks[[i]], noms_cellules, config)
      if (config$progress_enabled) setTxtProgressBar(pb, i)
      return(result)
    })
  }
  
  # Agrégation des résultats
  message("📊 Agrégation des résultats...")
  
  all_processed <- do.call(rbind, lapply(chunk_results, function(x) x$processed))
  all_errors <- do.call(rbind, lapply(chunk_results, function(x) x$errors))
  
  total_formulas <- nrow(formula_cells)
  success_count <- nrow(all_processed)
  error_count <- nrow(all_errors)
  success_rate <- round(100 * success_count / total_formulas, 1)
  
  # Nettoyage parallèle
  future::plan(future::sequential)
  
  # Construction du résultat final
  processing_time <- as.numeric(difftime(Sys.time(), start_time, units = "secs"))
  
  result <- list(
    formulas = all_processed,
    errors = if (nrow(all_errors) > 0) all_errors else NULL,
    sheets = sheets,
    globals = globals,
    noms_cellules = noms_cellules,
    processing_time = processing_time,
    statistics = list(
      total = total_formulas,
      success = success_count,
      errors = error_count,
      success_rate = success_rate
    ),
    config_used = config
  )
  
  # Sauvegarde en cache
  if (!is.null(file_signature)) {
    save_cached_result(result, file_signature, config)
  }
  
  # Génération du script si demandé
  if (emit_script && success_count > 0) {
    script_file <- generate_budgibot_script_v3(result, path)
    result$script_file <- script_file
  }
  
  # Messages finaux
  message(sprintf("✅ Parsing terminé en %.2f secondes", processing_time))
  message(sprintf("📈 Succès: %d/%d (%.1f%%)", success_count, total_formulas, success_rate))
  
  if (error_count > 0) {
    message(sprintf("⚠️ Erreurs: %d formules non converties", error_count))
  }
  
  return(result)
}

# ===== GÉNÉRATION SCRIPT BUDGIBOT =====
generate_budgibot_script_v3 <- function(parsed_result, original_path) {
  script_file <- paste0(tools::file_path_sans_ext(basename(original_path)), "_converted_formulas_v3.R")
  
  formulas <- parsed_result$formulas
  if (nrow(formulas) == 0) return(NULL)
  
  message(sprintf("📝 Génération script BudgiBot: %s", script_file))
  
  # En-tête du script
  lines <- c(
    "# ===== SCRIPT BUDGIBOT V3 - FORMULES EXCEL =====",
    sprintf("# Généré le %s", Sys.time()),
    sprintf("# %d formules traitées en %.2f secondes", 
            parsed_result$statistics$total, parsed_result$processing_time),
    sprintf("# Taux de succès: %.1f%%", parsed_result$statistics$success_rate),
    "",
    "script <- function(rv, rv_path = NULL) {",
    "  start_time <- Sys.time()",
    "  ",
    "  tryCatch({",
    "    # === INITIALISATION ===",
    "    wb <- rv$fichier_excel",
    "    if (is.null(wb) || !inherits(wb, 'wbWorkbook')) {",
    "      showNotification('❌ Workbook Excel invalide', type = 'error')",
    "      return()",
    "    }",
    "    ",
    "    message('🔄 Application des formules BudgiBot v3...')",
    "    ",
    "    # === CHARGEMENT DES DONNÉES ===",
    "    wb_sheets <- openxlsx2::wb_get_sheet_names(wb)",
    "    sheets <- setNames(",
    "      lapply(wb_sheets, function(sh) {",
    "        df <- openxlsx2::wb_to_df(wb, sheet = sh, col_names = FALSE)",
    "        # Conversion numérique optimisée",
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
    "        return(df)",
    "      }),",
    "      wb_sheets",
    "    )",
    "    "
  )
  
  # Groupement par feuille pour optimisation
  formulas_by_sheet <- split(formulas, formulas$sheet)
  
  lines <- c(lines, "    # === APPLICATION DES FORMULES PAR FEUILLE ===")
  
  for (sheet_name in names(formulas_by_sheet)) {
    sheet_formulas <- formulas_by_sheet[[sheet_name]]
    n_formulas <- nrow(sheet_formulas)
    
    lines <- c(lines,
               sprintf("    # FEUILLE: %s (%d formules)", sheet_name, n_formulas),
               sprintf("    if ('%s' %%in%% names(sheets)) {", sheet_name),
               sprintf("      values <- sheets[['%s']]", sheet_name),
               sprintf("      message('Traitement feuille %s: %d formules')", sheet_name, n_formulas),
               "      "
    )
    
    # Traitement par batch pour éviter les scripts trop longs
    batch_size <- 50
    n_batches <- ceiling(n_formulas / batch_size)
    
    for (batch in seq_len(n_batches)) {
      start_idx <- (batch - 1) * batch_size + 1
      end_idx <- min(batch * batch_size, n_formulas)
      batch_formulas <- sheet_formulas[start_idx:end_idx, ]
      
      if (n_batches > 1) {
        lines <- c(lines, sprintf("      # Batch %d/%d", batch, n_batches))
      }
      
      for (i in seq_len(nrow(batch_formulas))) {
        formula_row <- batch_formulas[i, ]
        
        # Ligne de commentaire avec la formule originale (tronquée)
        original_formula <- str_trunc(formula_row$formula, 60)
        lines <- c(lines, sprintf("      # %s -> %s", formula_row$address, original_formula))
        
        # Application avec gestion d'erreur
        lines <- c(lines,
                   "      tryCatch({",
                   sprintf("        sheets[['%s']][%d, %d] <- %s", 
                           sheet_name, formula_row$row_coord, formula_row$col_coord, 
                           formula_row$R_code),
                   sprintf("      }, error = function(e) { warning('Erreur %s!%s: ', e$message) })", 
                           sheet_name, formula_row$address)
        )
      }
      
      lines <- c(lines, "      ")
    }
    
    lines <- c(lines, 
               sprintf("      message('✅ Feuille %s terminée')", sheet_name),
               "    }",
               "    ")
  }
  
  # Finalisation du script
  lines <- c(lines,
             "    # === MISE À JOUR DU WORKBOOK ===",
             "    message('📊 Mise à jour du workbook...')",
             "    updated_sheets <- 0",
             "    ",
             "    for (sheet_name in names(sheets)) {",
             "      tryCatch({",
             "        openxlsx2::wb_remove_worksheet(wb, sheet_name)",
             "        openxlsx2::wb_add_worksheet(wb, sheet_name)",
             "        wb <- openxlsx2::wb_add_data(wb, sheet = sheet_name, x = sheets[[sheet_name]], col_names = FALSE)",
             "        updated_sheets <- updated_sheets + 1",
             "      }, error = function(e) {",
             "        message(sprintf('❌ Erreur mise à jour feuille %s: %s', sheet_name, e$message))",
             "      })",
             "    }",
             "    ",
             "    # === FINALISATION ===",
             "    rv$fichier_excel <- wb",
             "    rv$excel_updated <- Sys.time()",
             "    ",
             "    processing_time <- as.numeric(difftime(Sys.time(), start_time, units = 'secs'))",
             sprintf("    total_formulas <- %d", parsed_result$statistics$success),
             "    ",
             "    success_msg <- sprintf('✅ BudgiBot v3: %d formules appliquées en %.2f sec', total_formulas, processing_time)",
             "    message(success_msg)",
             "    showNotification(success_msg, type = 'message')",
             "    ",
             "  }, error = function(e) {",
             "    error_msg <- sprintf('❌ Erreur BudgiBot v3: %s', e$message)",
             "    message(error_msg)",
             "    showNotification(error_msg, type = 'error')",
             "  })",
             "}"
  )
  
  # Écriture du fichier
  writeLines(lines, script_file)
  
  message(sprintf("✅ Script généré: %s (%d lignes)", script_file, length(lines)))
  
  return(script_file)
}

# ===== FONCTION DE COMPATIBILITÉ =====
# Remplace l'ancienne fonction pour maintenir la compatibilité
parse_excel_formulas <- function(path, emit_script = FALSE) {
  message("🔄 Redirection vers parseur v3...")
  return(parse_excel_formulas_v3(path, emit_script))
}

message("✅ Parseur Excel v3 chargé - Optimisé pour gros fichiers")

# ===== CONFIGURATION POUR GLOBAL_CHATBOT.R =====
# Ajouter ces lignes dans global_chatbot.R pour configurer le parseur

# Configuration du parseur Excel optimisé
options(
  budgibot.chunk_size = 800,      # Ajuster selon vos besoins
  budgibot.max_memory = 1024,     # 1GB limite mémoire
  budgibot.workers = min(6, parallel::detectCores() - 1),
  budgibot.cache = TRUE,
  budgibot.progress = TRUE
)

message("📊 Configuration BudgiBot: chunks=", getOption("budgibot.chunk_size"), 
        ", workers=", getOption("budgibot.workers"),
        ", cache=", getOption("budgibot.cache"))