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
      tryCatch({
        withProgress(message = "Téléchargement audio", value = 0, {
          cmd <- sprintf("ffmpeg -y -i '%s' -vn -q:a 0 -map a '%s'", url, dest)
          system(cmd, wait = TRUE)
          incProgress(1)
        })
        
        if (file.exists(dest) && file.info(dest)$size > 0) {
          downloaded_file(dest)
          log_append("[✅] Audio téléchargé.")
          status("Audio prêt.")
          
          # Prévisualisation audio
          encoded_audio <- base64enc::dataURI(file = dest, mime = "audio/mp3")
          output$audio_preview <- renderUI({
            tags$audio(
              src = encoded_audio,
              type = "audio/mp3",
              controls = TRUE,
              style = "width:100%;"
            )
          })
        } else {
          log_error("Échec téléchargement.")
          status("Erreur de téléchargement.")
        }
      }, error = function(e) {
        log_error(paste0("Erreur lors du téléchargement : ", e$message))
        status("Erreur de téléchargement.")
      })
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
        fw <- import("faster_whisper")
        model <- tryCatch({
          fw$WhisperModel(input$whisper_model, device = "cpu", compute_type = "float32")
        }, error = function(e) {
          log_error(paste0("Erreur chargement modèle Faster-Whisper : ", e$message))
          NULL
        })
        
        if (is.null(model)) {
          status("Erreur de chargement du modèle.")
          return()
        }
        
        status("Transcription en cours...")
        df <- data.frame(Segment = integer(), Duration = numeric())
        full_transcript <- ""
        
        withProgress(message = "Segments", value = 0, {
          for (i in seq_along(segs)) {
            incProgress(1 / length(segs), detail = paste0(i, "/", length(segs)))
            t0 <- Sys.time()
            log_append(paste0("[🧠] Segment ", i, " / ", length(segs)))
            
            res <- tryCatch({
              segments <- model$transcribe(segs[i])
              log_append(paste0("[🔍] Structure de retour analysée..."))
              
              result_text <- ""
              
              # Méthode universelle pour toutes les versions de faster-whisper
              if (reticulate::py_has_attr(segments, "__iter__")) {
                log_append("[🔍] Itérable Python détecté")
                
                # Conversion en liste R
                segments_list <- reticulate::iterate(segments)
                
                for (segment in segments_list) {
                  if (reticulate::py_has_attr(segment, "text")) {
                    segment_text <- segment$text
                    result_text <- paste0(result_text, segment_text)
                    log_append(paste0("[📝] Texte extrait (", nchar(segment_text), " caractères)"))
                  } else {
                    log_append("[⚠️] Segment sans attribut 'text'")
                  }
                }
              } else {
                log_append("[❌] Format non itérable détecté")
              }
              
              if (nchar(result_text) == 0) {
                log_append("[🔧] Tentative alternative d'extraction...")
                try({
                  # Fallback pour certaines versions
                  segments_r <- reticulate::py_to_r(segments)
                  if (is.list(segments_r)) {
                    for (item in segments_r) {
                      if (is.list(item) && !is.null(item$text)) {
                        result_text <- paste0(result_text, item$text)
                      }
                    }
                  }
                }, silent = TRUE)
              }
              
              if (nchar(result_text) == 0) {
                log_append("[❌] Diagnostic complet:")
                log_append(paste0("- Type retour: ", class(segments)))
                log_append(paste0("- Méthodes disponibles: ", paste(reticulate::py_list_attributes(segments), collapse = ", ")))
                log_append("- Essayez: pip install --upgrade faster-whisper")
              }
              
              result_text
            }, error = function(e) {
              log_append(paste0("[❌] Erreur critique: ", e$message))
              NULL
            })
            
            dt <- as.numeric(difftime(Sys.time(), t0, units = "secs"))
            df <- rbind(df, data.frame(Segment = i, Duration = dt))
            
            if (!is.null(res) && nchar(res) > 0) {
              full_transcript <- paste0(full_transcript, res, "\n")
              log_append(paste0("[✅] Segment ", i, " terminé en ", round(dt, 1), "s."))
            } else {
              log_append(paste0("[⚠️] Aucun texte extrait pour le segment ", i))
            }
          }
        })
        
        segment_times(df)
        raw_text(full_transcript)
        transcript_text(full_transcript)
        log_append(paste0("[⏱️] Temps total de transcription : ", round(sum(df$Duration), 1), "s."))
        
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
        
        
      }, error = function(e) {
        log_error(paste0("Erreur lors de la transcription : ", e$message))
        status("Erreur de transcription.")
      })
      
      updateActionButton(session, "submit", label = "Transcrire", disabled = FALSE)
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
  })
}
