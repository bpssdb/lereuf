# parse_excel_formulas_with_assignments.R

library(tidyxl)
library(openxlsx)
library(dplyr)
library(stringr)

# Détermination du répertoire du script pour sourcer les dépendances
script_dir <- dirname(normalizePath(sys.frame(1)$ofile, mustWork = FALSE))
source(file.path(script_dir, "modules/parseur_excel/excel_formula_to_R.R"))
source(file.path(script_dir, "modules/parseur_excel/utils.R"))
source(file.path(script_dir, "modules/parseur_excel/utils_gros_fichier.R"))


# Ajoute cette fonction de correction après tes sources
fix_range_bug <- function(r_code) {
  # Corrige le bug X:X -> X dans les plages
  gsub("(\\d+):(\\1)", "\\1", r_code)
}



# Fonction principale --------------------------------------------------------
parse_excel_formulas <- function(path, emit_script = FALSE) {
  message("[parse_excel_formulas] Chargement du fichier Excel...")
  wb_sheets <- getSheetNames(path)
  sheets <- setNames(
    lapply(wb_sheets, function(sh) {
      message(sprintf(" - Lecture de la feuille '%s'", sh))
      read.xlsx(path, sheet = sh, colNames = FALSE)
    }), wb_sheets
  )
  message("[parse_excel_formulas] Extraction des variables globales")
  globals <- get_excel_globals(path)
  nom_cellules <- build_named_cell_map(globals)
  print(nom_cellules)
  
  message("[parse_excel_formulas] Extraction des cellules contenant des formules...")
  cells_all <- xlsx_cells(path)
  form_cells <- cells_all %>% filter(!is.na(formula))
  form_cells <- form_cells %>%
    mutate(
      formula = gsub("[\r\n]+", " ", formula),
      formula = gsub("\\s+", " ", formula),
      formula = trimws(formula)
    )
  message(sprintf("[parse_excel_formulas] %d formules extraites.", nrow(form_cells)))
  
  # Après l'extraction des formules, ajoute ça :
  message("[parse_excel_formulas] Filtrage des erreurs Excel...")
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
  
  message(sprintf("[parse_excel_formulas] %d formules restantes après filtrage (supprimées: %d).", 
                  nrow(form_cells), initial_count - nrow(form_cells)))
  
  message("[parse_excel_formulas] Extraction des dépendances...")
  form_cells <- form_cells %>%
    mutate(
      deps = purrr::pmap(list(formula, sheet), extract_deps)
    )
  
  message("[parse_excel_formulas] Filtrage des formules contenant 'LEFT'...")
  initial_count <- nrow(form_cells)
  form_cells <- form_cells %>%
    filter(!grepl("\\bLEFT\\b", formula, ignore.case = TRUE))
  message(sprintf("[parse_excel_formulas] %d formules restantes après filtrage (supprimées: %d).", 
                  nrow(form_cells), initial_count - nrow(form_cells)))
  
  message("[parse_excel_formulas] Conversion des formules Excel en code R...")
  form_cells <- form_cells %>%
    mutate(
      R_code = vapply(formula, function(f) convert_formula(f, nom_cellules), character(1)),
      row    = as.integer(sub("^[A-Z]+", "", address)),
      col    = vapply(gsub("[0-9]+", "", address), function(x) as.integer(col2num(x)), integer(1))
    ) %>%
    select(sheet, address, row, col, formula, R_code)
  message("[parse_excel_formulas] Conversion terminée.")
  
  if (!emit_script) {
    message("[parse_excel_formulas] Evaluation des cellules dans le bon ordre...")
    sheets <- evaluate_cells(sheets, form_cells)
  }
  
  if (emit_script) {
    message("[parse_excel_formulas] Génération du script vectorisé...")
    
    script_file <- paste0(tools::file_path_sans_ext(basename(path)), "_converted_formulas.R")
    
    # Utiliser la version vectorisée
    generate_vectorized_script(form_cells, script_file)
  }
  
  
  #Ancien script : if (emit_script) {
  if (FALSE) {
    message("[parse_excel_formulas] Génération du script R...")
    
    script_file <- paste0(tools::file_path_sans_ext(basename(path)), "_converted_formulas.R")
    
    lines <- c(
      "# Script généré par parse_excel_formulas()",
      "# Charger le classeur",
      "script <- function(rv, rv_path) {",
      "tryCatch({",
      "path <- rv_path()",
      "print(path)",
      "wb <- rv$fichier_excel",
      "if (is.null(wb) || !inherits(wb, \"wbWorkbook\")) {",
      "showNotification(\"❌ Le workbook Excel n'est pas valide.\", type =  \"error\")",
      "output$status <- renderText(\"❌ Workbook Excel manquant ou invalide.\")",
      "return()",
      "}",
      "wb_sheets <- openxlsx2::wb_get_sheet_names(wb)",
      "sheets <- setNames(",
      "lapply(wb_sheets, function(sh) openxlsx2::wb_to_df(wb, sheet = sh, colNames = FALSE)),",
      "wb_sheets)",
      "    # Conversion forcée en numérique pour les calculs",
      "    sheets <- lapply(sheets, function(df) {",
      "      df[] <- lapply(df, function(col) {",
      "        # Essai de conversion numérique",
      "        if (is.character(col) || is.factor(col)) {",
      "          # Remplacer les chaînes vides et NA par 0",
      "          col[is.na(col) | col == '' | col == 'NA'] <- '0'",
      "          # Conversion en numérique",
      "          col_num <- suppressWarnings(as.numeric(col))",
      "          # Si échec, garder 0",
      "          col_num[is.na(col_num)] <- 0",
      "          return(col_num)",
      "        } else if (is.logical(col)) {",
      "          return(as.numeric(col))",
      "        } else {",
      "          # Remplacer les NA par 0 même pour les numériques",
      "          col[is.na(col)] <- 0",
      "          return(col)",
      "        }",
      "      })",
      "      df",
      "    })"
    )
    
    # ⚡ SECTION CALCULS (ça manquait !) ⚡
    for (i in seq_len(nrow(form_cells))) {
      sheet_val <- form_cells$sheet[i]
      row_val <- form_cells$row[i] 
      col_val <- form_cells$col[i]
      code_val <- fix_range_bug(form_cells$R_code[i])
      address_val <- form_cells$address[i]
      formula_val <- form_cells$formula[i]
      
      assign_line <- sprintf(
        "tryCatch({ sheets[[\"%s\"]][%d, %d] <- %s }, error = function(e) { warning(sprintf('Erreur conversion %s!%s : %%s', e$message), call. = FALSE) })",
        sheet_val, row_val, col_val, code_val, sheet_val, address_val
      )
      
      lines <- c(lines,
                 sprintf("# %s!%s -> %s", sheet_val, address_val, formula_val),
                 sprintf("values <- sheets[[\"%s\"]]", sheet_val),
                 assign_line,
                 ""
      )
    }
    
    # ⚡ SECTION MISE À JOUR WORKBOOK ⚡
    lines <- c(lines,
               "# === DEBUG : Vérifier les calculs ===",
               "message('=== VÉRIFICATION CALCULS ===')"
    )
    
    # Ajouter quelques vérifications debug
    unique_sheets <- unique(form_cells$sheet)
    for (sheet_name in unique_sheets) {
      sheet_cells <- form_cells[form_cells$sheet == sheet_name, ]
      for (i in 1:min(3, nrow(sheet_cells))) {  # Max 3 exemples par feuille
        lines <- c(lines,
                   sprintf("message('%s %s = ', sheets[['%s']][%d, %d])", 
                           sheet_cells$sheet[i], sheet_cells$address[i], 
                           sheet_cells$sheet[i], sheet_cells$row[i], sheet_cells$col[i])
        )
      }
    }
    
    lines <- c(lines,
               "",
               "# === MISE À JOUR DU WORKBOOK ===",
               "tryCatch({"
    )
    
    # Pour chaque cellule calculée, ajouter la mise à jour
    for (i in seq_len(nrow(form_cells))) {
      sheet_val <- form_cells$sheet[i]
      row_val <- form_cells$row[i] 
      col_val <- form_cells$col[i]
      
      lines <- c(lines,
                 sprintf("  # Mise à jour %s!%s", sheet_val, form_cells$address[i]),
                 sprintf("  wb <- openxlsx2::wb_add_data(wb, sheet = '%s', x = sheets[['%s']][%d, %d], start_col = %d, start_row = %d)",
                         sheet_val, sheet_val, row_val, col_val, col_val, row_val)
      )
    }
    
    lines <- c(lines,
               "  message('✅ Workbook mis à jour avec succès')",
               "}, error = function(e) {",
               "  message('❌ Erreur mise à jour workbook: ', e$message)",
               "})",
               "",
               "rv$fichier_excel <- wb",
               "rv$excel_updated <- Sys.time()",
               "showNotification('✅ Formules appliquées !', type = 'message')",
               "",
               "}, error = function(e) {",
               "showNotification(paste0(\"❌ Erreur du script:\", e$message), type = \"error\")",
               "})",
               "}"
    )
    
    writeLines(lines, script_file)
    message("Script écrit dans : ", script_file)
  }
}