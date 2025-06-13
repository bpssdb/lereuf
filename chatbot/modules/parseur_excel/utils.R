# Utilitaires ----------------------------------------------------------------
vlookup_r <- function(lookup, df, col_index) {
  idx <- match(lookup, df[[1]])
  ifelse(is.na(idx), NA, df[[col_index]][idx])
}

extract_deps <- function(form, sheet_name) {
  pattern <- "(?:'([^']+)'!)?([A-Z]+[0-9]+)"
  m <- gregexpr(pattern, form, perl = TRUE)
  refs <- regmatches(form, m)[[1]]
  deps <- vapply(refs, function(x) {
    parts <- regmatches(x, regexec(pattern, x, perl = TRUE))[[1]]
    sh <- if (nzchar(parts[2])) parts[2] else sheet_name
    paste0(sh, "!", parts[3])
  }, character(1))
  unique(deps)
}

topo_order <- function(ids, deps_list) {
  remaining <- setNames(deps_list, ids)
  result <- character(0)
  while (length(remaining) > 0) {
    ready <- names(remaining)[vapply(remaining, length, integer(1)) == 0]
    if (length(ready) == 0) stop("Boucle détectée dans les dépendances !")
    result <- c(result, ready)
    remaining[ready] <- NULL
    for (i in seq_along(remaining)) {
      remaining[[i]] <- setdiff(remaining[[i]], ready)
    }
  }
  result
}

evaluate_cells <- function(sheets, form_cells) {
  # construire les identifiants
  ids <- paste0(form_cells$sheet, "!", form_cells$address)
  # topologiquement ordonner
  order_ids <- topo_order(ids, form_cells$deps)
  # mapping id -> informations de position
  pos <- setNames(
    Map(function(sh, addr, row, col) {
      list(sheet = sh, row = row, col = col)
    }, form_cells$sheet, form_cells$address, form_cells$row, form_cells$col),
    ids
  )
  
  for (id in order_ids) {
    info <- pos[[id]]
    code <- form_cells$R_code[ids == id]
    
    # --- nouveau bloc : on crée un env avec 'values' = la feuille en cours ---
    ctx <- new.env(parent = globalenv())
    ctx$sheets    <- sheets             # pour vlookup_r, etc.
    ctx$vlookup_r <- vlookup_r          # votre helper
    ctx$col2num   <- col2num            # votre helper
    ctx$values    <- sheets[[info$sheet]]  # *la* feuille courante
    
    # évaluation dans ce contexte
    val <- eval(parse(text = code), envir = ctx)
    
    # on écrit dans le data.frame de la feuille
    sheets[[info$sheet]][info$row, info$col] <- val
  }
  
  sheets
}

get_excel_globals <- function(path) {
  wb <- loadWorkbook(path)
  noms <- getNamedRegions(wb)
  positions <- attr(noms, "position")
  print(noms)
  print(positions)
  
  # On garde les plages qui pointent vers UNE SEULE cellule (ex : "D77", "B5", etc.)
  est_cellule_unique <- grepl("^\\$?[A-Z]+\\$?[0-9]+$", positions)
  # Association nom → cellule
  assoc_cellules <- data.frame(
    nom = noms[est_cellule_unique],
    cellule = positions[est_cellule_unique],
    sheet = attr(noms, "sheet")[est_cellule_unique],
    stringsAsFactors = FALSE
  )
  return(assoc_cellules)  # Résultat : liste des noms associés à une cellule unique
}

build_named_cell_map <- function(df) {
  refs <- lapply(seq_len(nrow(df)), function(i) {
    list(
      ref   = convert_ref(df$cellule[i]),
      sheet = df$sheet[i]
    )
  })
  names(refs) <- df$nom
  refs
}

