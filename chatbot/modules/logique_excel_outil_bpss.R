insert_data_to_sheet <- function(wb, sheet_name, data, startCol, startRow) {
  if (!inherits(wb, "wbWorkbook")) stop("❌ wb n'est pas un workbook openxlsx2.")
  if (!sheet_name %in% openxlsx2::wb_get_sheet_names(wb)) {
    wb <- openxlsx2::wb_add_worksheet(wb, sheet_name)
    
  }
  wb <- openxlsx2::wb_add_data(wb, sheet = sheet_name, x = data, startCol = startCol, startRow = startRow)
  return(wb)
}

# Etape 1 : Fonction R qui reproduit la logique de ChargementDonnéesDePaye_v10()
load_ppes_data <- function(wb, df_pp_categ, df_entrants, df_sortants, code_programme) {
  if (!inherits(wb, "wbWorkbook")) stop("❌ wb n'est pas un workbook valide")
  
  nb_rows <- 100
  ligne1 <- 7
  ligne2 <- ligne1 + nb_rows + 6   # 113
  ligne3 <- ligne2 + nb_rows + 5   # 218
  
  code_prefix <- substr(code_programme, 1, 3)
  
  df1 <- dplyr::filter(df_pp_categ, substr(as.character(nom_prog), 1, 3) == code_prefix)
  df2 <- dplyr::filter(df_entrants,  substr(as.character(nom_prog), 1, 3) == code_prefix)
  df3 <- dplyr::filter(df_sortants,  substr(as.character(nom_prog), 1, 3) == code_prefix)
  
  # Nettoyage des plages
  wb <- insert_data_to_sheet(wb, "Données PP-E-S", data.frame(), startCol = 3, startRow = ligne1)
  wb <- insert_data_to_sheet(wb, "Données PP-E-S", data.frame(), startCol = 3, startRow = ligne2)
  wb <- insert_data_to_sheet(wb, "Données PP-E-S", data.frame(), startCol = 3, startRow = ligne3)
  
  # Insertion des données
  wb <- insert_data_to_sheet(wb, "Données PP-E-S", df1, startCol = 3, startRow = ligne1)
  wb <- insert_data_to_sheet(wb, "Données PP-E-S", df2, startCol = 3, startRow = ligne2)
  wb <- insert_data_to_sheet(wb, "Données PP-E-S", df3, startCol = 3, startRow = ligne3)
  
  # Traitement "Indicié"
  df_indicie <- dplyr::filter(df1, marqueur_masse_indiciaire == "Indicié")
  
  if (nrow(df_indicie) > 0) {
    wb <- insert_data_to_sheet(wb, "Accueil", data.frame(), startCol = 2, startRow = 43)
    wb <- insert_data_to_sheet(wb, "Accueil", data.frame(), startCol = 3, startRow = 43)
    wb <- insert_data_to_sheet(wb, "Accueil", df_indicie[[2]], startCol = 2, startRow = 43)
    wb <- insert_data_to_sheet(wb, "Accueil", df_indicie[[6]], startCol = 3, startRow = 43)
  }
  print('ppes OK')
  return(wb)
}


# Etape 2 : Fonction R qui reproduit la logique de ChargerINFDPP18_v6()
load_inf_dpp18 <- function(wb, df_dpp18, code_programme) {
  if (!inherits(wb, "wbWorkbook")) stop("❌ wb n'est pas un workbook valide")
  if (is.null(df_dpp18) || nrow(df_dpp18) == 0) return(wb)
  print("dpp 18 entrée")
  # Nettoyage
  wb <- insert_data_to_sheet(wb, "INF DPP 18", data.frame(), startCol = 2, startRow = 1)
  
  # Header
  header <- head(df_dpp18, 5)
  wb <- insert_data_to_sheet(wb, "INF DPP 18", header, startCol = 2, startRow = 1)
  print("DPP 18 Header")
  # Filtrage

  df_filtered <- df_dpp18 %>%
    mutate(col1 = as.character(.[[1]])) %>%
    filter(stringr::str_detect(col1, fixed(substr(code_programme, 1, 3)))) %>%
    select(-col1)
  
  #df_filtered <- dplyr::filter(df_dpp18, 
  #                             stringr::str_detect(as.character(.data[[1]]), fixed(substr(code_programme, 1, 3))))
  wb <- insert_data_to_sheet(wb, "INF DPP 18", df_filtered, startCol = 2, startRow = 6)
  return(wb)
}

# Etape 3 : Fonction R qui reproduit la logique de ChargerINFDPP45_v2_corrige()
load_inf_bud45 <- function(wb, df_bud45, code_programme) {
  if (!inherits(wb, "wbWorkbook")) stop("❌ wb n'est pas un workbook valide")
  if (is.null(df_bud45) || nrow(df_bud45) == 0) return(wb)
  
  wb <- insert_data_to_sheet(wb, sheet_name = "INF BUD 45", data = data.frame(), startCol = 2, startRow = 1)
  
  header <- head(df_bud45, 5)
  wb <- insert_data_to_sheet(wb, "INF BUD 45", header, startCol = 2, startRow = 1)
  
  # Correction du filtrage
  df_filtered <- df_bud45 %>%
    dplyr::mutate(col2 = as.character(.[[2]])) %>%
    dplyr::filter(stringr::str_detect(col2, fixed(substr(code_programme, 1, 3)))) %>%
    select(-col2)
  
  wb <- insert_data_to_sheet(wb, "INF BUD 45", df_filtered, startCol = 2, startRow = 6)
  print("Bud 45 OK")
  return(wb)
}


