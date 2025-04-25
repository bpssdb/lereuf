# Etape 3 : Fonction R qui reproduit la logique de ChargerINFDPP45_v2_corrige()


load_inf_bud45 <- function(wb, df_bud45, code_programme) {
  if (!inherits(wb, "wbWorkbook")) stop("❌ wb n'est pas un workbook valide")
  if (is.null(df_bud45) || nrow(df_bud45) == 0) return(wb)
  
  # Effacer colonnes B à Z (laisser colonne A)
  wb <- openxlsx2::wb_add_data(wb, sheet = "INF BUD 45", x = "", dims = paste0("B1:Z3000"))
  
  # Copier les 5 premières lignes non filtrées
  header <- head(df_bud45, 5)
  wb <- openxlsx2::wb_add_data(wb, "INF BUD 45", header, startCol = 2, startRow = 1)
  
  # Filtrage sur la colonne B (index 2)
  df_filtered <- dplyr::filter(df_bud45, stringr::str_detect(.data[[2]], fixed(code_programme)))
  
  # Insertion à partir de ligne 6
  wb <- openxlsx2::wb_add_data(wb, "INF BUD 45", df_filtered, startCol = 2, startRow = 6)
  
  return(wb)
}