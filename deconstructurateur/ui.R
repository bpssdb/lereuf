# ui.R

# Assurez-vous que shinyjs est bien initialisé
ui <- fluidPage(
  useShinyjs(),
  titlePanel("📊 Déconstructurateur d'Excel"),
  sidebarLayout(
    sidebarPanel(
      fileInput("upload_file", "📂 Charger un fichier Excel (.xlsx)"),
      uiOutput("sheet_selector"),
      actionButton("load_sheet", "📄 Charger la feuille"),
      actionButton("auto_detect_zone", "🔍 Détecter zone", class = "btn btn-info mt-2"),
      actionButton("tag_zone", "➕ Taguer cette zone", class = "btn btn-success mt-2"),
      actionButton("manual_tag", "📝 Tag manuel", class = "btn btn-warning mt-2"),
      actionButton("toggle_label", "Sélectionner/désélectionner le label", class = "btn btn-info mt-2"),
      
      verbatimTextOutput("zone_type"),
      tags$div(
        style = "margin-top: 10px;",
        textOutput("selection_info"),
        textOutput("zone_stats")
      ),
      # Dans ui.R ou dans la partie UI de app.R
      downloadModuleUI("download1"),
      hr(),
      h4("📝 Visualisation des tags"),
      reactableOutput("tags_view", height = "200px")
    ),
    mainPanel(
      div(
        id = "excel_table_container",
        style = "overflow: auto; white-space: nowrap; max-height: 600px; width: 1000px;",
        rhandsontable::rHandsontableOutput("excel_table")
      )
    )
  )
)