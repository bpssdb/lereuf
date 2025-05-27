# fix_parseur_v3_complete.R
# Correctif COMPLET pour le parseur Excel v3

library(stringr)
library(dplyr)

message("🔧 Correctif complet pour parseur Excel v3...")

# ===== FONCTIONS DE BASE =====

excel_col_to_num <- function(col_str) {
  if (is.na(col_str) || !nzchar(col_str)) return(1)
  chars <- strsplit(col_str, "")[[1]]
  sum(sapply(seq_along(chars), function(i) {
    (match(chars[i], LETTERS)) * 26^(length(chars) - i)
  }))
}

col2num <- excel_col_to_num

get_excel_globals <- function(path) {
  tryCatch({
    wb <- openxlsx::loadWorkbook(path)
    noms <- openxlsx::getNamedRegions(wb)
    if (length(noms) == 0) return(data.frame())
    positions <- attr(noms, "position")
    est_cellule_unique <- grepl("^\\$?[A-Z]+\\$?[0-9]+$", positions)
    data.frame(
      nom = noms[est_cellule_unique],
      cellule = positions[est_cellule_unique],
      sheet = attr(noms, "sheet")[est_cellule_unique],
      stringsAsFactors = FALSE
    )
  }, error = function(e) {
    return(data.frame())
  })
}

build_named_cell_map <- function(df) {
  if (nrow(df) == 0) return(list())
  refs <- lapply(seq_len(nrow(df)), function(i) {
    list(
      ref = paste0("values[", str_extract(df$cellule[i], "\\d+"), ", ", 
                   excel_col_to_num(str_extract(df$cellule[i], "^[A-Z]+")), "]"),
      sheet = df$sheet[i]
    )
  })
  names(refs) <- df$nom
  refs
}

convert_criteria <- function(crit, noms_cellules = NULL) {
  crit <- str_trim(crit)
  if (str_starts(crit, "=")) {
    return(convert_criteria(sub("^=", "", crit), noms_cellules))
  }
  if (str_detect(crit, '^[<>]=?|^!=')) {
    op <- str_extract(crit, '^[<>]=?|^!=')
    val <- sub('^[<>]=?|^!=', '', crit)
    return(list(type = "compare", op = op, value = val))
  }
  if (str_detect(crit, '^[0-9.]+$')) {
    return(list(type = "number", value = crit))
  }
  if (str_detect(crit, '^[A-Za-z]+[0-9]+$')) {
    return(list(type = "ref", value = convert_formula(crit, noms_cellules)))
  }
  txt <- str_replace_all(crit, '^"|"$', '')
  return(list(type = "text", value = paste0('"', txt, '"')))
}

# ===== FONCTIONS AVANCÉES =====

parse_multiple_ranges <- function(range_str) {
  ranges <- str_split(range_str, ",")[[1]] %>% str_trim()
  converted_ranges <- sapply(ranges, function(r) {
    if (str_detect(r, ":")) {
      parts <- str_split(r, ":")[[1]]
      start_row <- str_extract(parts[1], "\\d+")
      start_col <- excel_col_to_num(str_extract(parts[1], "^[A-Z]+"))
      end_row <- str_extract(parts[2], "\\d+")
      end_col <- excel_col_to_num(str_extract(parts[2], "^[A-Z]+"))
      return(paste0("values[", start_row, ":", end_row, ", ", start_col, ":", end_col, "]"))
    } else {
      row_num <- str_extract(r, "\\d+")
      col_str <- str_extract(r, "^[A-Z]+")
      col_num <- excel_col_to_num(col_str)
      return(paste0("values[", row_num, ", ", col_num, "]"))
    }
  })
  return(paste0("c(", paste(converted_ranges, collapse = ", "), ")"))
}

parse_function_args <- function(formula_inside) {
  depth <- 0
  current_arg <- ""
  args <- character(0)
  chars <- strsplit(formula_inside, "")[[1]]
  
  for (char in chars) {
    if (char == "(") {
      depth <- depth + 1
      current_arg <- paste0(current_arg, char)
    } else if (char == ")") {
      depth <- depth - 1
      current_arg <- paste0(current_arg, char)
    } else if (char == "," && depth == 0) {
      args <- c(args, str_trim(current_arg))
      current_arg <- ""
    } else {
      current_arg <- paste0(current_arg, char)
    }
  }
  if (nzchar(current_arg)) {
    args <- c(args, str_trim(current_arg))
  }
  return(args)
}

split_by_operator <- function(formula, operator) {
  depth <- 0
  chars <- strsplit(formula, "")[[1]]
  op_char <- if (operator == "\\+") "+" else 
    if (operator == "\\*") "*" else 
      if (operator == "/") "/" else 
        if (operator == "-") "-" else operator
  
  for (i in seq_along(chars)) {
    char <- chars[i]
    if (char == "(") {
      depth <- depth + 1
    } else if (char == ")") {
      depth <- depth - 1
    } else if (depth == 0 && char == op_char) {
      left <- str_trim(paste(chars[1:(i-1)], collapse = ""))
      right <- str_trim(paste(chars[(i+1):length(chars)], collapse = ""))
      return(c(left, right))
    }
  }
  return(c(formula))
}

# ===== FONCTION CONVERT_FORMULA COMPLÈTE =====

convert_formula <- function(formula, noms_cellules = list()) {
  if (is.na(formula) || !nzchar(formula)) return(NA_character_)
  
  # Nettoyage
  f <- str_trim(str_replace_all(formula, c("^=" = "", "\\$" = "", "<>" = "!=")))
  
  # Nombres
  if (str_detect(f, "^[0-9.]+$")) {
    return(f)
  }
  
  # Références simples (A1, B5)
  if (str_detect(f, "^[A-Z]+\\d+$")) {
    row_num <- str_extract(f, "\\d+")
    col_str <- str_extract(f, "^[A-Z]+")
    col_num <- excel_col_to_num(col_str)
    return(paste0("values[", row_num, ", ", col_num, "]"))
  }
  
  # Plages simples (A1:B5)
  if (str_detect(f, "^[A-Z]+\\d+:[A-Z]+\\d+$")) {
    parts <- str_split(f, ":")[[1]]
    start_row <- str_extract(parts[1], "\\d+")
    start_col <- excel_col_to_num(str_extract(parts[1], "^[A-Z]+"))
    end_row <- str_extract(parts[2], "\\d+")
    end_col <- excel_col_to_num(str_extract(parts[2], "^[A-Z]+"))
    return(paste0("values[", start_row, ":", end_row, ", ", start_col, ":", end_col, "]"))
  }
  
  # Fonction MAX
  if (str_detect(f, "^MAX\\(")) {
    inside <- str_sub(f, 5, -2)  # Enlever MAX( et )
    args <- parse_function_args(inside)
    converted_args <- sapply(args, function(arg) convert_formula(arg, noms_cellules))
    return(paste0("max(", paste(converted_args, collapse = ", "), ", na.rm = TRUE)"))
  }
  
  # Fonction MIN
  if (str_detect(f, "^MIN\\(")) {
    inside <- str_sub(f, 5, -2)
    args <- parse_function_args(inside)
    converted_args <- sapply(args, function(arg) convert_formula(arg, noms_cellules))
    return(paste0("min(", paste(converted_args, collapse = ", "), ", na.rm = TRUE)"))
  }
  
  # Fonction SUM
  if (str_detect(f, "^SUM\\(")) {
    inside <- str_sub(f, 5, -2)
    
    # Plages multiples (G7:G12,G15:G24)
    if (str_detect(inside, ",") && str_detect(inside, ":")) {
      converted_ranges <- parse_multiple_ranges(inside)
      return(paste0("sum(", converted_ranges, ", na.rm = TRUE)"))
    } else if (str_detect(inside, ":")) {
      # Plage simple
      converted_range <- convert_formula(inside, noms_cellules)
      return(paste0("sum(", converted_range, ", na.rm = TRUE)"))
    } else {
      # Arguments séparés
      args <- parse_function_args(inside)
      converted_args <- sapply(args, function(arg) convert_formula(arg, noms_cellules))
      return(paste0("sum(", paste(converted_args, collapse = ", "), ", na.rm = TRUE)"))
    }
  }
  
  # Division (priorité haute)
  if (str_detect(f, "/") && !str_detect(f, "^[A-Z]+\\(")) {
    parts <- split_by_operator(f, "/")
    if (length(parts) == 2) {
      left <- convert_formula(parts[1], noms_cellules)
      right <- convert_formula(parts[2], noms_cellules)
      return(paste0("(", left, ") / (", right, ")"))
    }
  }
  
  # Multiplication
  if (str_detect(f, "\\*") && !str_detect(f, "^[A-Z]+\\(")) {
    parts <- split_by_operator(f, "\\*")
    if (length(parts) == 2) {
      left <- convert_formula(parts[1], noms_cellules)
      right <- convert_formula(parts[2], noms_cellules)
      return(paste0("(", left, ") * (", right, ")"))
    }
  }
  
  # Addition
  if (str_detect(f, "\\+") && !str_detect(f, "^[A-Z]+\\(")) {
    parts <- split_by_operator(f, "\\+")
    if (length(parts) == 2) {
      left <- convert_formula(parts[1], noms_cellules)
      right <- convert_formula(parts[2], noms_cellules)
      return(paste0("(", left, ") + (", right, ")"))
    }
  }
  
  # Soustraction
  if (str_detect(f, "-") && !str_detect(f, "^[A-Z]+\\(")) {
    parts <- split_by_operator(f, "-")
    if (length(parts) == 2) {
      left <- convert_formula(parts[1], noms_cellules)
      right <- convert_formula(parts[2], noms_cellules)
      return(paste0("(", left, ") - (", right, ")"))
    }
  }
  
  # Fonction IF
  if (str_detect(f, "^IF\\(")) {
    inside <- str_sub(f, 4, -2)
    args <- parse_function_args(inside)
    if (length(args) >= 2) {
      test_expr <- convert_formula(args[1], noms_cellules)
      true_expr <- convert_formula(args[2], noms_cellules)
      false_expr <- if (length(args) >= 3) convert_formula(args[3], noms_cellules) else "NA"
      return(paste0("ifelse(", test_expr, ", ", true_expr, ", ", false_expr, ")"))
    }
  }
  
  # SUMIF
  if (str_detect(f, "^SUMIF\\(")) {
    inside <- str_sub(f, 7, -2)
    args <- parse_function_args(inside)
    if (length(args) >= 2) {
      range <- convert_formula(args[1], noms_cellules)
      crit_raw <- args[2]
      sum_range <- if (length(args) >= 3) convert_formula(args[3], noms_cellules) else range
      
      crit <- convert_criteria(crit_raw, noms_cellules)
      expr_test <- switch(
        crit$type,
        compare = sprintf("(%s %s %s)", range, crit$op, crit$value),
        sprintf("(%s == %s)", range, crit$value)
      )
      return(paste0("sum((", expr_test, ")*", sum_range, ", na.rm=TRUE)"))
    }
  }
  
  # Si rien ne correspond
  return(paste0("# Complex: ", f))
}

# ===== ASSIGNATION =====

assign("excel_col_to_num", excel_col_to_num, envir = .GlobalEnv)
assign("col2num", col2num, envir = .GlobalEnv)
assign("convert_formula", convert_formula, envir = .GlobalEnv)
assign("convert_criteria", convert_criteria, envir = .GlobalEnv)
assign("get_excel_globals", get_excel_globals, envir = .GlobalEnv)
assign("build_named_cell_map", build_named_cell_map, envir = .GlobalEnv)
assign("parse_multiple_ranges", parse_multiple_ranges, envir = .GlobalEnv)
assign("parse_function_args", parse_function_args, envir = .GlobalEnv)
assign("split_by_operator", split_by_operator, envir = .GlobalEnv)

message("✅ Correctif complet appliqué")

# ===== TESTS =====

message("🧪 Test des formules de votre fichier...")

test_formulas <- c(
  "MAX((1+SUM(G7,I7,K7)/SUM(F7,H7,J7)),1)",
  "SUM(G7:G12,G15:G24)", 
  "G7",
  "1",
  "SUM(A1,B2,C3)"
)

for (i in seq_along(test_formulas)) {
  formula <- test_formulas[i]
  tryCatch({
    result <- convert_formula(formula)
    status <- if (str_starts(result, "# Complex:") || str_starts(result, "# TODO:")) "⚠️" else "✅"
    message(sprintf("%s Test %d: %s", status, i, substr(formula, 1, 40)))
    if (status == "✅") {
      message(sprintf("    → %s", substr(result, 1, 60)))
    }
  }, error = function(e) {
    message(sprintf("❌ Test %d: %s → ERREUR: %s", i, substr(formula, 1, 40), e$message))
  })
}

message("🏁 Correctif terminé ! Relancez le parseur v3.")