# app.R

# Charger les fichiers globaux, UI et serveur
source("~/work/budgibot/deconstructurateur/global.R")
source("~/work/budgibot/deconstructurateur/ui.R")
source("~/work/budgibot/deconstructurateur/server.R")

# Lancer l'application Shiny
shinyApp(ui, server)