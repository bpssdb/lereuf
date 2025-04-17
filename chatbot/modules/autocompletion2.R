library(shiny)
library(DT)

# =============================
# UI du module de mapping
# =============================
mod_budget_mapping_ui <- function(id) {
  # Création du namespace pour ce module afin d'éviter toute collision des IDs
  ns <- NS(id)
  tagList(
    # Affichage du tableau interactif avec DT
    DT::DTOutput(ns("mapping_table")),
    br(),  # Saut de ligne pour espacer les éléments UI
    # Bouton pour valider le mapping effectué sur le tableau
    actionButton(ns("validate_mapping"), "Valider le mapping")
  )
}


# =============================
# Server du module de mapping
# =============================
# Ce module gère l'affichage et la modification en ligne du data frame contenant
# les données budgétaires extraites. Le paramètre 'mapping_data' doit être une reactiveVal ou reactive() 
# contenant le data frame à éditer.
mod_budget_mapping_server <- function(id, mapping_data, tags_json) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    
    # Tableau DT éditable
    output$mapping_table <- DT::renderDT({
      req(mapping_data())
      DT::datatable(mapping_data(), options = list(scrollX = TRUE), editable = TRUE)
    })
    
    # Édition inline
    observeEvent(input$mapping_table_cell_edit, {
      info <- input$mapping_table_cell_edit
      df   <- mapping_data()
      df[info$row, info$col] <- DT::coerceValue(info$value, df[info$row, info$col])
      mapping_data(df)
    })
    
    # 👉 Au clic sur “Valider le mapping” : appel du helper et mise à jour
    observeEvent(input$validate_mapping, {
      df   <- mapping_data()
      tags <- tags_json()
      req(nrow(df)>0, length(tags)>0)
      
      mapping <- map_budget_entries(df, tags)
      if (is.null(mapping) || nrow(mapping)==0) {
        showNotification("❌ Le mapping a échoué.", type="error")
        return()
      }
      
      # on fusionne la colonne 'cellule' en CelluleCible
      merged <- dplyr::left_join(
        df,
        mapping %>% dplyr::select(Axe, Description, cellule),
        by = c("Axe","Description")
      )
      merged$CelluleCible <- merged$cellule
      merged$cellule <- NULL
      
      mapping_data(merged)
      showNotification("✅ Mapping validé et appliqué.", type="message")
    })
  })
}

