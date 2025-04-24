library(shiny)
library(httr)
library(rvest)
library(stringr)
library(reticulate)
library(base64enc)

mod_videoTranscriberServer <- function(id, trigger) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    transcript_text <- reactiveVal("")
    raw_text <- reactiveVal("")
    log_messages <- reactiveVal("[📥] Prêt à traiter une vidéo ou un lien.\n")
    status <- reactiveVal("En attente d'une action utilisateur.")
    detected_urls <- reactiveVal(character(0))
    downloaded_file <- reactiveVal(NULL)
    segment_times <- reactiveVal(data.frame(Segment=integer(), Duration=numeric()))
    
    log_append <- function(msg) {
      msg_full <- paste0(msg, "\n")
      cat(msg_full)  # Affiche aussi dans la console
      isolate(log_messages(paste0(log_messages(), msg_full)))
    }
    
    log_error <- function(msg) {
      log_append(paste0("[❌] ", msg))
    }
    
    observeEvent(trigger(), {
      removeModal()
      log_append("[📥] Module ouvert.")
      status("Saisie URL ou upload de fichier...")
      detected_urls(character(0))
      downloaded_file(NULL)
      
      showModal(modalDialog(
        title = "Transcription vidéo / audio",
        size = "l",
        easyClose = TRUE,
        tagList(
          div(style = "margin-bottom:1em;", strong("Statut : "), textOutput(ns("current_status"))),
          tags$div(
            style="max-height:150px; overflow-y:auto; background:#f9f9f9; padding:0.5em; border:1px solid #ccc; font-family:monospace;",
            verbatimTextOutput(ns("log"))
          ),
          selectInput(
            ns("whisper_model"), 
            "Modèle Whisper", 
            choices = c("tiny", "base", "small", "medium", "large"), 
            selected = "base"
          ),
          textInput(
            ns("page_url"), 
            "URL page/flux (.m3u8/.mp4)", 
            placeholder = "https://..."
          ),
          actionButton(
            ns("detect_btn"), 
            "Détecter les flux de la page", 
            class = "btn btn-outline-primary mb-2"
          ),
          uiOutput(ns("detected_ui")),
          fileInput(
            ns("video_file"), 
            "Ou chargez un fichier local", 
            accept = c("video/*", "audio/*")
          ),
          actionButton(
            ns("download_btn"), 
            "Télécharger l'audio", 
            class = "btn btn-info mb-2"
          ),
          uiOutput(ns("audio_preview")),
          plotOutput(ns("timeline_plot"), height = "200px")
        ),
        footer = tagList(
          actionButton(ns("submit"), "Transcrire", class = "btn btn-success"),
          modalButton("Fermer")
        )
      ))
    }, ignoreInit = TRUE)
    
    output$log <- renderText({
      log_messages()
    })
    
    output$current_status <- renderText({
      status()
    })
    
    observeEvent(input$detect_btn, {
      req(input$page_url)
      url <- input$page_url
      log_append(paste0("[🔍] Chargement de : ", url))
      status("Chargement de la page...")
      
      page_text <- tryCatch({
        content(GET(url), "text", encoding = "UTF-8")
      }, error = function(e) {
        log_error("Impossible de charger la page.")
        status("Erreur de chargement.")
        NULL
      })
      
      if (is.null(page_text)) return()
      
      status("Extraction des liens...")
      page <- tryCatch({
        read_html(page_text)
      }, error = function(e) {
        log_error("Erreur lors de l'analyse HTML.")
        status("Erreur d'analyse HTML.")
        NULL
      })
      
      if (is.null(page)) return()
      
      # Extraction des liens vidéo
      metas <- page %>% html_nodes("meta[property='og:video']") %>% html_attr("content")
      contenturls <- page %>% html_nodes("meta[itemprop='contentURL']") %>% html_attr("content")
      video_src <- page %>% html_nodes("video source, video") %>% html_attr("src")
      scripts <- paste(page %>% html_nodes("script") %>% html_text(), collapse = " ")
      
      pattern <- "https?://[^\"']+\\.(?:m3u8|mp4)(\\?[^\"']*)?"
      script_links <- unlist(str_extract_all(scripts, pattern))
      dynamic_links <- unlist(str_extract_all(page_text, pattern))
      
      candidates <- unique(na.omit(c(metas, contenturls, video_src, script_links, dynamic_links)))
      candidates <- candidates[str_detect(candidates, "\\.(mp4|m3u8)(\\?|$)")]
      
      if (length(candidates) == 0) {
        log_append("[⚠️] Aucun flux .mp4/.m3u8 détecté.")
        status("Aucun flux détecté.")
      } else {
        log_append(paste0("[✅] Flux détectés : ", paste(candidates, collapse = ", ")))
        status("Flux détectés. Choisissez-en un.")
      }
      detected_urls(candidates)
    })
    
    output$detected_ui <- renderUI({
      urls <- detected_urls()
      if (length(urls) == 0) return(NULL)
      
      selectInput(
        ns("video_url_choice"),
        "Flux à transcrire",
        choices = urls
      )
    })
    
    observeEvent(input$download_btn, {
      req(input$video_url_choice)
      url <- input$video_url_choice
      log_append(paste0("[🎬] Téléchargement audio : ", url))
      status("Téléchargement...")
      
      dest <- tempfile(fileext = ".mp3")
      dir.create("www", showWarnings = FALSE)
      
      future({
        cmd <- sprintf(
          "ffmpeg -y -i '%s' -map 0:a:0 -vn -q:a 0 -acodec libmp3lame '%s'",
          url,
          dest
        )
        system(cmd, wait = TRUE)
        dest  # on retourne le chemin vers le fichier téléchargé
      }) %...>% {
        file_path <- .
        
        if (file.exists(file_path) && file.info(file_path)$size > 0) {
          downloaded_file(file_path)
          log_append("[✅] Audio téléchargé.")
          status("Audio prêt.")
          
          # Copie du fichier dans www/ avec un nom unique
          www_name <- paste0("audio_", format(Sys.time(), "%Y%m%d%H%M%S"), ".mp3")
          www_path <- file.path("www", www_name)
          file.copy(file_path, www_path, overwrite = TRUE)
          
          output$audio_preview <- renderUI({
            tags$audio(
              src = www_name,
              type = "audio/mp3",
              controls = TRUE,
              style = "width:100%;"
            )
          })
        } else {
          log_error("Échec du téléchargement.")
          status("Erreur de téléchargement.")
        }
      } %...!% {
        error <- .
        log_error(paste0("Erreur lors du téléchargement : ", error$message))
        status("Erreur de téléchargement.")
      }
    })
    
    observeEvent(input$submit, {
      req(downloaded_file())
      updateActionButton(session, "submit", label = "Transcription...", disabled = TRUE)
      status("Découpage audio...")
      src <- downloaded_file()
      
      # Segmentation audio
      seg_dir <- tempfile()
      dir.create(seg_dir)
      log_append("[⚙️] Segmentation avec ffmpeg...")
      seg_pattern <- file.path(seg_dir, "seg_%03d.mp3")
      cmd <- sprintf("ffmpeg -y -i '%s' -f segment -segment_time 1800 -c copy '%s'", src, seg_pattern)
      log_append(paste0("[💻] Commande : ", cmd))
      
      tryCatch({
        system(cmd, wait = TRUE)
        segs <- list.files(seg_dir, pattern = "seg_[0-9]{3}\\.mp3$", full.names = TRUE)
        
        if (length(segs) == 0) {
          log_error("Aucun segment créé.")
          status("Erreur de segmentation.")
          return()
        }
        
        log_append(paste0("[🔪] ", length(segs), " segments créés : ", paste(basename(segs), collapse = ", ")))
        
        # Chargement du modèle faster-whisper
        status("Transcription via WhisperX en cours...")
        
        # 📁 Création dossier de sortie pour whisperx
        output_dir <- tempfile("whisperx_")
        dir.create(output_dir, recursive = TRUE)
        
        # 📦 Lancement de whisperx en CLI
        cmd <- sprintf(
          "whisperx %s --language French --output_dir %s --model %s --output_format txt --diarize --device cpu",
          shQuote(downloaded_file()), 
          shQuote(output_dir),
          input$whisper_model
        )
        
        log_append(paste0("[🐍] Exécution WhisperX : ", cmd))
        
        tryCatch({
          output <- system(cmd, intern = TRUE)
          log_append(paste(output, collapse = "\n"))
        }, error = function(e) {
          log_error(paste0("Erreur WhisperX : ", e$message))
          status("Erreur de transcription.")
          updateActionButton(session, "submit", label = "Transcrire", disabled = FALSE)
          return()
        })
        
        # 📄 Vérification + attente de génération du fichier texte
        txt_file <- list.files(output_dir, pattern = "\\.txt$", full.names = TRUE)
        
        i <- 0
        while (length(txt_file) == 0 && i < 10) {
          Sys.sleep(0.5)  # Attente active
          i <- i + 1
          txt_file <- list.files(output_dir, pattern = "\\.txt$", full.names = TRUE)
        }
        
        if (length(txt_file) == 1) {
          log_append(paste("📄 Fichier texte détecté après", i * 0.5, "s :", basename(txt_file)))
          content <- readLines(txt_file, encoding = "UTF-8")
          full_transcript <- paste(content, collapse = "\n")
          raw_text(full_transcript)
          transcript_text(full_transcript)
          log_append("[✅] Transcription lue avec succès.")
        } else {
          log_append("[❌] Fichier texte introuvable même après attente.")
          status("Erreur : aucun fichier généré.")
          return()
        }
        
        
        # ✅ Affichage de fin
        removeModal()
        showModal(modalDialog(
          title = "Transcription terminée",
          size = "l",
          easyClose = TRUE,
          tagList(
            verbatimTextOutput(ns("raw_transcript")),
            fluidRow(
              column(6, downloadButton(ns("download_raw"), "Exporter transcription brute")),
              column(6, actionButton(ns("generate_summary"), "Générer un compte rendu", class = "btn-primary"))
            )
          ),
          footer = modalButton("Fermer")
        ))
        updateActionButton(session, "submit", label = "Transcrire", disabled = FALSE)
      })  
    })
    
    
    output$raw_transcript <- renderText({
      txt <- raw_text()
      if (nzchar(txt)) txt else "[⚠️] Aucun contenu généré."
    })
    
    output$download_raw <- downloadHandler(
      filename = function() {
        paste0("transcription_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".txt")
      },
      content = function(file) {
        writeLines(raw_text(), file)
      }
    )
    
    observeEvent(input$generate_summary, {
      req(raw_text())
      log_append("[🧠] Génération du compte rendu via Ollama...")
      status("Génération du compte rendu...")
      
      summary_prompt <- paste0(
        "Tu es un assistant administratif. Voici une transcription brute d'une audition parlementaire. ",
        "Génère un compte rendu structuré destiné à l'administration, avec les sections suivantes :\n",
        "- Liste des participants\n- Résumé par thème\n- Citations clés\n- Conclusion\n\n",
        "Texte source :\n",
        raw_text()
      )
      
      tryCatch({
        summary <- system2("ollama", args = c("run", "mistral"), input = summary_prompt, stdout = TRUE)
        summary_txt <- paste(summary, collapse = "\n")
        transcript_text(summary_txt)
        log_append("[📄] Compte rendu généré.")
        
        showModal(modalDialog(
          title = "Compte rendu généré",
          size = "l",
          easyClose = TRUE,
          tagList(
            verbatimTextOutput(ns("transcript")),
            downloadButton(ns("download_summary"), "Exporter compte rendu")
          ),
          footer = modalButton("Fermer")
        ))
      }, error = function(e) {
        log_error(paste0("Erreur lors de la génération du résumé : ", e$message))
        status("Erreur de génération du résumé.")
      })
    })
    
    output$transcript <- renderText({
      txt <- transcript_text()
      if (nzchar(txt)) txt else "[⚠️] Aucun contenu généré."
    })
    
    output$download_summary <- downloadHandler(
      filename = function() {
        paste0("compte_rendu_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".txt")
      },
      content = function(file) {
        writeLines(transcript_text(), file)
      }
    )
    
    output$download_raw <- downloadHandler(
      filename = function() {
        paste0("transcription_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".txt")
      },
      content = function(file) {
        writeLines(raw_text(), file)
      }
    )
  })
}