# global_chatbot.R - Version mise à jour avec parseur v3

library(DT)
library(shiny)
library(shinyjs)
library(httr)
library(pdftools)
library(later)
library(officer)
library(promises)
library(future)
library(mime)
library(reticulate)
library(callr)
library(reactable)
library(bslib)
library(shinycssloaders)

# === NOUVELLE CONFIGURATION PARSEUR V3 ===
# Configuration du parseur Excel optimisé
options(
  budgibot.chunk_size = 800,      # Taille optimale pour la plupart des fichiers
  budgibot.max_memory = 1024,     # 1GB limite mémoire
  budgibot.workers = min(6, parallel::detectCores() - 1),  # Parallélisation intelligente
  budgibot.cache = TRUE,          # Cache activé pour éviter le reprocessing
  budgibot.progress = TRUE        # Barres de progression
)

message("📊 Configuration BudgiBot v3:")
message(sprintf("  - Chunks: %d formules", getOption("budgibot.chunk_size")))
message(sprintf("  - Workers: %d parallèles", getOption("budgibot.workers")))
message(sprintf("  - Mémoire: %d MB max", getOption("budgibot.max_memory")))
message(sprintf("  - Cache: %s", if(getOption("budgibot.cache")) "✅ Activé" else "❌ Désactivé"))

# === SOURCES EXISTANTES ===
source("~/work/lereuf/chatbot/helpers/api_mistral.R")
source("~/work/lereuf/chatbot/helpers/llm_utils.R")
source("~/work/lereuf/chatbot/helpers/ui_helpers.R")
source("~/work/lereuf/chatbot/helpers/server_utils.R")
source("~/work/lereuf/chatbot/helpers/autocompletion_helper.R")
source("~/work/lereuf/chatbot/helpers/map_budget_entries.R")  

# === MODULES AVEC PARSEUR V3 ===
source("~/work/lereuf/chatbot/modules/mod_mesures_cat_combo_v3.R")  # Version mise à jour
source("~/work/lereuf/chatbot/modules/mod_outil_bpss.R")
source("~/work/lereuf/chatbot/modules/mod_json_helper.R")
source("~/work/lereuf/chatbot/modules/autocompletion.R")
source("~/work/lereuf/chatbot/modules/video_transcription_module.R")
# Détermination du répertoire du script pour sourcer les dépendances
source("~/work/lereuf/chatbot/modules/parseur_excel/excel_formula_to_R2.R")
source("~/work/lereuf/chatbot/modules/parseur_excel/utils.R")

# === PARSEUR EXCEL V3 ===
source("~/work/lereuf/chatbot/modules/parseur_excel/parseur_excel_v3.R")  # Nouveau parseur optimisé

# Configuration
plan(multisession)  # Or multicore if on Linux

Sys.setenv(MISTRAL_API_KEY = Sys.getenv("MISTRAL_API_KEY"))
mistral_api_key <- Sys.getenv("MISTRAL_API_KEY")
cat("DEBUG: API Key is <", Sys.getenv("MISTRAL_API_KEY"), ">\n")

# Initialize reticulate virtualenv and module import
reticulate::use_virtualenv("r-reticulate", required = TRUE)
extract_msg <- reticulate::import("extract_msg")

cat("DEBUG: API Key is <", Sys.getenv("MISTRAL_API_KEY"), ">\n")

# === FONCTION DE LECTURE FICHIER (inchangée) ===
read_file_content <- function(file_path, file_name) {
  file_ext <- tools::file_ext(file_name)
  
  if (file_ext == "txt") {
    content <- readLines(file_path, warn = FALSE)
    content <- paste(content, collapse = "\n")
    
  } else if (file_ext == "pdf") {
    content <- pdf_text(file_path)
    content <- paste(content, collapse = "\n")
    
  } else if (file_ext == "docx") {
    doc <- read_docx(file_path)
    content <- docx_summary(doc)$text
    content <- paste(content, collapse = "\n")
    
  } else if (file_ext == "msg") {
    m <- extract_msg$Message(file_path)
    
    subject <- if (!is.null(m$subject) && nzchar(m$subject)) enc2utf8(m$subject) else "(Aucun sujet)"
    sender <- if (!is.null(m$sender) && nzchar(m$sender)) enc2utf8(m$sender) else "(Expéditeur inconnu)"
    
    recipients <- tryCatch({
      recips <- m$recipients
      if (length(recips) == 0 || is.null(recips)) {
        "(Destinataires inconnus)"
      } else {
        recipient_emails <- sapply(recips, function(x) {
          if (!is.null(x$email) && grepl("@", x$email)) {
            enc2utf8(x$email)
          } else {
            NA
          }
        })
        recipient_emails <- recipient_emails[!is.na(recipient_emails)]
        if (length(recipient_emails) > 0) paste(recipient_emails, collapse = "; ") else "(Destinataires inconnus)"
      }
    }, error = function(e) {
      message("Error processing recipients: ", e$message)
      "(Destinataires inconnus)"
    })
    
    date <- if (!is.null(m$date) && nzchar(m$date)) as.character(m$date) else "(Date inconnue)"
    
    body_raw <- m$body
    body <- if (is.null(body_raw) || length(body_raw) == 0 || !nzchar(body_raw)) {
      "(Aucun contenu dans l'email)"
    } else {
      if (!is.character(body_raw)) body_raw <- as.character(body_raw)
      body_utf8 <- iconv(body_raw, from = "", to = "UTF-8", sub = "")
      body_clean <- paste(body_utf8, collapse = "\n")
      if (nchar(trimws(body_clean)) == 0) "(Aucun contenu dans l'email)" else body_clean
    }
    
    content <- paste(
      "Type de message : Mail",
      "Sujet :", subject,
      "\nDe :", sender,
      "\nÀ :", recipients,
      "\nDate :", date,
      "\n\n--- Contenu du message ---\n\n", body
    )
    
    print("Final content assembled!")
    
  } else {
    content <- "(Format de fichier non pris en charge)"
  }
  
  return(content)
}

# === FONCTIONS UTILITAIRES POUR LE PARSEUR V3 ===

# Fonction pour nettoyer le cache si nécessaire
clean_budgibot_cache <- function() {
  cache_dir <- file.path(tempdir(), "budgibot_cache")
  if (dir.exists(cache_dir)) {
    files <- list.files(cache_dir, full.names = TRUE)
    # Supprimer les fichiers de plus de 24h
    old_files <- files[file.mtime(files) < (Sys.time() - 24*3600)]
    if (length(old_files) > 0) {
      file.remove(old_files)
      message(sprintf("🧹 Cache nettoyé: %d fichiers supprimés", length(old_files)))
    }
  }
}

# Fonction pour ajuster la configuration dynamiquement
adjust_parser_config <- function(file_size_mb) {
  if (file_size_mb < 5) {
    # Petit fichier
    options(budgibot.chunk_size = 1500, budgibot.workers = 1)
  } else if (file_size_mb < 20) {
    # Fichier moyen
    options(budgibot.chunk_size = 1000, budgibot.workers = min(4, parallel::detectCores() - 1))
  } else if (file_size_mb < 50) {
    # Gros fichier
    options(budgibot.chunk_size = 500, budgibot.workers = min(6, parallel::detectCores() - 1))
  } else {
    # Très gros fichier
    options(budgibot.chunk_size = 200, budgibot.workers = min(8, parallel::detectCores() - 1))
    options(budgibot.max_memory = 2048)  # Augmenter la limite mémoire
  }
}

# Nettoyage automatique du cache au démarrage
clean_budgibot_cache()

message("✅ BudgiBot global_chatbot v3 chargé avec parseur optimisé")
message("🚀 Prêt pour le traitement de gros fichiers Excel")