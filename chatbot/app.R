# app.R

source("~/work/budgibot/chatbot/global_chatbot.R")  # Charger les dépendances et les fonctions partagées
source("~/work/budgibot/chatbot/ui_chatbotv2.R")      # Charger l'interface utilisateur
source("~/work/budgibot/chatbot/server_chatbot.R")  # Charger la logique du serveur

shinyApp(ui = ui, server = server)
