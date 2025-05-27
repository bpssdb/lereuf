# app_v3.R - BudgiBot avec parseur Excel optimisé

# Charger la configuration globale mise à jour
source("~/work/lereuf/chatbot/global_chatbot_v3.R")  # Version mise à jour
source("~/work/lereuf/chatbot/ui_chatbot.R")      # UI inchangée
source("~/work/lereuf/chatbot/server_chatbot2.R")    # Serveur inchangé

# Message de démarrage avec info version
message(strrep("=", 50))
message("🤖 BudgiBot v3 - Parseur Excel Optimisé")
message("📊 Configuration active:")
message(sprintf("   • Chunks: %d", getOption("budgibot.chunk_size")))
message(sprintf("   • Workers: %d", getOption("budgibot.workers")))
message(sprintf("   • Mémoire max: %d MB", getOption("budgibot.max_memory")))
message(sprintf("   • Cache: %s", if(getOption("budgibot.cache")) "Activé" else "Désactivé"))
message(strrep("=", 50))

# Lancement de l'application
shinyApp(ui = ui, server = server)