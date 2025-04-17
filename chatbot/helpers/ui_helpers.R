# helpers/ui_helpers.R

show_budget_modal <- function() {
  showModal(modalDialog(
    title = "Données budgétaires extraites",
    size = "l",
    easyClose = TRUE,
    footer = tagList(
      modalButton("Fermer"),
      actionButton("save_budget_changes", "Enregistrer")
    ),
    DTOutput("budget_table")
  ))
}