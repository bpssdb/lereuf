# app.R

source("~/work/lereuf/chatbot/global_chatbot.R")  # Charger les dépendances et les fonctions partagées
source("~/work/lereuf/chatbot/ui_chatbotv2.R")      # Charger l'interface utilisateur
source("~/work/lereuf/chatbot/server_chatbot.R")  # Charger la logique du serveur

shinyApp(ui = ui, server = server)