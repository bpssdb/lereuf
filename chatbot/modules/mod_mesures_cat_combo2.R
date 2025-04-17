# mod_mesures_cat.R

# ============================
# MODULE UI
# ============================

# Fonction UI du module "mesures_cat"
mod_mesures_cat_ui <- function(id) {
  # Création du namespace pour éviter les conflits de noms d'inputs/output
  ns <- NS(id)
  
  # Retourne un ensemble de composants UI organisés dans un tagList
  tagList(
    fluidRow(
      # Colonne de gauche : permet de charger un fichier Excel
      column(6,
             fileInput(ns("upload_file"), "📂 Charger un fichier Excel (.xlsx)")
      ),
      # Colonne de droite : affichage dynamique du sélecteur de feuille
      column(6,
             uiOutput(ns("sheet_selector"))
      )
    ), 
    
    # Affichage du tableau reactable avec défilement horizontal en cas de dépassement
    div(style = "overflow-x: auto;", reactableOutput(ns("reactable_table"))),
    
    # Bouton pour ouvrir un éditeur en plein écran (pour modifier la feuille)
    actionButton(ns("open_full_editor"), "🖋️ Modifier la feuille", class = "btn btn-secondary mt-2"),
    # Bouton pour télécharger/exporter le tableau modifié en Excel
    downloadButton(ns("download_table"), "💾 Exporter le tableau modifié")
  )
}


# ============================
# MODULE SERVER
# ============================

# Fonction server du module "mesures_cat" 
# Le paramètre on_analysis_summary (optionnel) permet de fournir une fonction de rappel recevant un résumé de l'analyse du fichier Excel.
mod_mesures_cat_server <- function(id, rv, on_analysis_summary = NULL) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns  # Namespace local
    
    # Chargement des packages nécessaires au traitement
    library(readxl)         # Pour lire les fichiers Excel
    library(rhandsontable)  # Pour l'édition en tableau interactif
    library(openxlsx2)      # Pour l'export vers Excel
    library(reactable)       # Pour l'affichage du tableau avec reactable
    
    # Réactifs pour stocker les données importantes
    rv_table <- reactiveVal()    # Stocke le data frame de la feuille Excel sélectionnée
    rv_sheets <- reactiveVal()   # Stocke les noms des feuilles présentes dans le fichier
    rv_path <- reactiveVal()     # Stocke le chemin du fichier Excel chargé
    
    # ---------------------------
    # 1️⃣ Lecture du fichier Excel et extraction des feuilles
    # ---------------------------
    observeEvent(input$upload_file, {
      req(input$upload_file)
      # Vérifier l'extension du fichier (doit être xlsx)
      ext <- tools::file_ext(input$upload_file$name)
      if (tolower(ext) != "xlsx") {
        showNotification("Veuillez charger un fichier Excel (.xlsx)", type = "error")
        return()
      }
      # Récupère le chemin temporaire du fichier uploadé
      path <- input$upload_file$datapath
      rv_path(path)
      # Extraction des noms de feuilles
      sheets <- excel_sheets(path)
      rv_sheets(sheets)
    })
    
    # ---------------------------
    # 2️⃣ Affichage du sélecteur de feuille
    # ---------------------------
    # UI dynamique qui propose un selectInput pour choisir la feuille parmi celles détectées.
    output$sheet_selector <- renderUI({
      req(rv_sheets())
      selectInput(ns("selected_sheet"), "🗂️ Choisir une feuille", choices = rv_sheets())
    })
    
    # ---------------------------
    # 3️⃣ Lecture et traitement de la feuille sélectionnée
    # ---------------------------
    observeEvent(input$selected_sheet, {
      req(rv_path(), input$selected_sheet)
      
      # Lecture de la feuille sélectionnée (en ignorant les 4 premières lignes)
      df <- read_excel(rv_path(), sheet = input$selected_sheet, skip = 4)
      rv_table(as.data.frame(df))
      
      # Appel de la fonction de résumé sur le fichier Excel
      if (!is.null(on_analysis_summary)) {
        summary <- analyser_feuille_excel(df, sheet_name = input$selected_sheet)
        on_analysis_summary(summary)
      }
    })
    
    # ---------------------------
    # 4️⃣ Affichage du tableau via Reactable
    # ---------------------------
    output$reactable_table <- reactable::renderReactable({
      req(rv_table())
      df <- rv_table()
      reactable(
        df,
        searchable = TRUE,
        resizable = TRUE,
        highlight = TRUE,
        bordered = TRUE,
        striped = TRUE,
        pagination = FALSE,  # La pagination est désactivée pour afficher tout le tableau
        defaultColDef = colDef(
          minWidth = 100,
          style = list(
            whiteSpace = "pre-wrap",    # Permet de gérer les retours à la ligne
            overflow = "hidden",
            textOverflow = "ellipsis",
            fontSize = "12px",
            padding = "4px"
          )
        ),
        theme = reactableTheme(
          borderColor = "#ddd",
          stripedColor = "#f6f8fa",
          highlightColor = "#eaeaea"
        ),
        # Fixe une hauteur maximale et active le scroll vertical si nécessaire
        style = list(maxHeight = "70vh", overflowY = "auto")
      )
    })
    
    # ---------------------------
    # 5️⃣ Edition complète en mode plein écran (via rHandsontable)
    # ---------------------------
    observeEvent(input$open_full_editor, {
      # Affiche un modal avec une interface d'édition en plein écran
      showModal(modalDialog(
        title = "🖋️ Édition plein écran",
        easyClose = TRUE,
        footer = tagList(
          modalButton("❌ Fermer"),
          actionButton(ns("save_edits"), "💾 Enregistrer")
        ),
        # Affichage du tableau éditable
        rhandsontable::rHandsontableOutput(ns("hot_table"), height = "70vh"),
        size = "l",
        style = "width: 95vw; max-width: none;"
      ))
    })
    
    # Rendu du tableau éditable via rHandsontable
    output$hot_table <- rhandsontable::renderRHandsontable({
      req(rv_table())
      df <- rv_table()
      nb_cols <- ncol(df)
      col_widths <- rep(120, nb_cols)  # Largeur fixe de 120px par colonne (modifiable)
      
      rhandsontable(df, rowHeaders = NULL, height = "calc(100vh - 150px)", width = "100%") %>%
        hot_cols(colWidths = col_widths) %>%
        hot_table(stretchH = "none")  # Désactivation de l'étirement horizontal
    })
    
    # Sauvegarde des modifications après édition dans le modal et fermeture du modal
    observeEvent(input$save_edits, {
      rv_table(hot_to_r(input$hot_table))
      removeModal()
    })
    
    # ---------------------------
    # 6️⃣ Exportation du tableau modifié vers un fichier Excel
    # ---------------------------
    output$download_table <- downloadHandler(
      filename = function() {
        paste0("mesures_cat_", Sys.Date(), ".xlsx")
      },
      content = function(file) {
        wb <- wb_workbook()
        wb_add_worksheet(wb, "Mesures")
        wb_add_data(wb, sheet = 1, x = rv_table())
        wb_save(wb, file)
      }
    )
    
    # Stocker les données Excel dans un environnement partagé
    observeEvent(rv_table(), {
      req(rv_table(), input$selected_sheet)
      rv$excel_data <- rv_table()
      rv$excel_sheet <- input$selected_sheet
    })
    
  })
}
