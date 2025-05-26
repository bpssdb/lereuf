# parse_excel_formulas_with_assignments.R

library(tidyxl)
library(openxlsx)
library(dplyr)
library(stringr)

# Détermination du répertoire du script pour sourcer les dépendances
script_dir <- dirname(normalizePath(sys.frame(1)$ofile, mustWork = FALSE))
source(file.path(script_dir, "modules/parseur_excel/excel_formula_to_R.R"))
source(file.path(script_dir, "modules/parseur_excel/utils.R"))

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
    message("[parse_excel_formulas] Génération du script R...")
    script_file <- paste0(tools::file_path_sans_ext(basename(path)), "_converted_formulas.R")
    lines <- c(
      "# Script généré par parse_excel_formulas()",
      "# Charger le classeur",
      "script <- function(rv, rv_path) {",
      "tryCatch({",
      "path <- rv_path()",  # <- ajout du path ici
      "print(path)",
      "wb <- rv$fichier_excel",
      "if (is.null(wb) || !inherits(wb, \"wbWorkbook\")) {",
      "showNotification(\"❌ Le workbook Excel n'est pas valide.\", type =  \"error\")",
      "output$status <- renderText(\"❌ Workbook Excel manquant ou invalide.\")",
      "return()",
      "}",
      #"wb_sheets <- getSheetNames(wb)",
      #"sheets <- setNames(lapply(wb_sheets, function(sh) read.xlsx(path, sheet = sh, colNames = FALSE)), wb_sheets)",
      "wb_sheets <- openxlsx2::wb_get_sheet_names(wb)",
      "sheets <- setNames(",
      "lapply(wb_sheets, function(sh) openxlsx2::wb_to_df(wb, sheet = sh, colNames = FALSE)),",
      "wb_sheets)",
      "    # coercition automatique des colonnes character en numeric",
      "    sheets <- lapply(sheets, function(df) {",
      "      df[] <- lapply(df, function(col) {",
      "        if (is.character(col)) {",
      "          col_num <- suppressWarnings(type.convert(col, as.is = TRUE))",
      "          if (!identical(is.na(col_num), is.na(col))) col_num else col",
      "        } else {",
      "          col",
      "        }",
      "      })",
      "      df",
      "    })"
    )
    
    for (i in seq_len(nrow(form_cells))) {
      # ✅ Accès direct par nom de colonne
      sheet_val <- form_cells$sheet[i]
      row_val <- form_cells$row[i] 
      col_val <- form_cells$col[i]
      code_val <- form_cells$R_code[i]
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
    
    lines <- c(lines, 
               "rv$fichier_excel <- wb",
               "rv$excel_updated <- Sys.time()",
               "}, error = function(e) {",
               "showNotification(paste0(\"❌ Erreur du script:\", e$message), type = \"error\")",
               "}", 
               ")", 
               "}"
    )
    
    writeLines(lines, script_file)
    message("Script écrit dans : ", script_file)
  }
  
  message("[parse_excel_formulas] Terminé.")
}