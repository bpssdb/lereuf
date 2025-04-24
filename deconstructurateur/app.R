# app.R

# Charger les fichiers globaux, UI et serveur
source("~/work/lereuf/deconstructurateur/global.R")
source("~/work/lereuf/deconstructurateur/ui.R")
source("~/work/lereuf/deconstructurateur/server.R")

# Lancer l'application Shiny
shinyApp(ui, server)
