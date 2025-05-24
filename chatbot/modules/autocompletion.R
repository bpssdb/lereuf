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
# autocompletion_helper.R

mod_budget_mapping_server <- function(id, mapping_data, tags_json) {
  moduleServer(id, function(input, output, session) {
    
    # 1️⃣ Affiche le tableau avec cellules éditables
    output$mapping_table <- DT::renderDT({
      req(mapping_data())
      DT::datatable(
        mapping_data(),
        options = list(scrollX = TRUE),
        editable = TRUE
      )
    })
    
    # 2️⃣ Mise à jour inline
    observeEvent(input$mapping_table_cell_edit, {
      info <- input$mapping_table_cell_edit
      df   <- mapping_data()
      df[info$row, info$col + 1] <- DT::coerceValue(info$value, df[info$row, info$col + 1])
      mapping_data(df)
    })
    
    # 3️⃣ Quand l’utilisateur valide, on appelle map_budget_entries
    observeEvent(input$validate_mapping, {
      df    <- mapping_data()
      tags  <- tags_json()
      # si pas de tags, on sort
      if (is.null(tags) || length(tags)==0) {
        showNotification("Pas de tags JSON pour le mapping.", type="error")
        return()
      }
      mapped <- map_budget_entries(df, tags)
      if (is.null(mapped) || nrow(mapped)==0) {
        showNotification("Le mapping a échoué.", type="error")
        return()
      }
      # on renomme et rejoint
      mapped <- mapped %>%
        dplyr::select(Axe, Description, cellule) %>%
        dplyr::rename(CelluleCible = cellule)
      
      new_df <- dplyr::left_join(df, mapped, by = c("Axe","Description"))
      # on met à jour
      mapping_data(new_df)
      
      removeModal()
      showNotification("✅ Mapping validé !", type="message")
    })
  })
}


