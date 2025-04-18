# modules/video_transcription/R/video_transcription_module.R

library(shiny)
library(httr)
library(rvest)
library(stringr)
library(reticulate)

# Serveur du module Shiny de transcription vidéo/audio
mod_videoTranscriberServer <- function(id, trigger) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    transcript_text <- reactiveVal("")
    log_messages   <- reactiveVal("[📥] Prêt à traiter une vidéo ou un lien.\n")
    status         <- reactiveVal("En attente d'une action utilisateur.")
    detected_urls  <- reactiveVal(character(0))
    downloaded_file<- reactiveVal(NULL)
    segment_times  <- reactiveVal(data.frame(Segment=integer(), Duration=numeric()))
    
    log_append <- function(msg) {
      isolate(log_messages(paste0(log_messages(), msg, "\n")))
    }
    
    observeEvent(trigger(), {
      removeModal()
      log_append("[📥] Module ouvert.")
      status("Saisie URL ou upload de fichier...")
      detected_urls(character(0)); downloaded_file(NULL)
      
      showModal(modalDialog(
        title = "Transcription vidéo / audio",
        size  = "l",
        easyClose = TRUE,
        tagList(
          div(style = "margin-bottom:1em;", strong("Statut : "), status()),
          tags$div(style="max-height:150px; overflow-y:auto; background:#f9f9f9; padding:0.5em; border:1px solid #ccc; font-family:monospace;",
                   verbatimTextOutput(ns("log"))),
          selectInput(ns("whisper_model"), "Modèle Whisper", 
                      choices = c("tiny","base","small","medium","large"), selected = "base"),
          textInput(ns("page_url"), "URL page/flux (.m3u8/.mp4)", placeholder="https://..."),
          actionButton(ns("detect_btn"), "Détecter les flux de la page", class="btn btn-outline-primary mb-2"),
          uiOutput(ns("detected_ui")),
          fileInput(ns("video_file"), "Ou chargez un fichier local", accept=c("video/*","audio/*")),
          actionButton(ns("download_btn"), "Télécharger l'audio", class="btn btn-info mb-2"),
          uiOutput(ns("audio_preview")),
          plotOutput(ns("timeline_plot"), height="200px")
        ),
        footer = tagList(
          actionButton(ns("submit"), "Transcrire", class="btn btn-success"),
          modalButton("Fermer")
        )
      ))
    }, ignoreInit = TRUE)
    
    output$log <- renderText(log_messages())
    
    observeEvent(input$detect_btn, {
      req(input$page_url)
      url <- input$page_url
      log_append(paste0("[🔍] Chargement de : ", url))
      status("Chargement de la page...")
      
      page_text <- tryCatch(content(GET(url), "text", encoding="UTF-8"), error = function(e) NULL)
      if (is.null(page_text)) {
        log_append("[❌] Impossible de charger la page.")
        status("Erreur de chargement.")
        return()
      }
      
      status("Extraction des liens...")
      page <- read_html(page_text)
      metas <- page %>% html_nodes("meta[property='og:video']") %>% html_attr("content")
      contenturls <- page %>% html_nodes("meta[itemprop='contentURL']") %>% html_attr("content")
      video_src <- page %>% html_nodes("video source, video") %>% html_attr("src")
      scripts <- page %>% html_nodes("script") %>% html_text()
      script_links <- unlist(str_extract_all(scripts, 'https?:\/\/[^"\']+\\.(m3u8|mp4)[^"\']*'))
      dynamic_links<- unlist(str_extract_all(page_text, 'https?:\/\/[^"\']+\\.(m3u8|mp4)[^"\']*'))
      
      candidates <- unique(na.omit(c(metas, contenturls, video_src, script_links, dynamic_links)))
      candidates <- candidates[str_detect(candidates, "\\.(mp4|m3u8)(\\?|$)")]
      
      if (length(candidates)==0) {
        log_append("[⚠️] Aucun flux .mp4/.m3u8 détecté.")
        status("Aucun flux détecté.")
      } else {
        log_append(paste0("[✅] Flux détectés : ", paste(candidates, collapse=", ")))
        status("Flux détectés. Choisissez-en un.")
      }
      detected_urls(candidates)
      
      output$detected_ui <- renderUI({
        urls <- detected_urls()
        if (length(urls)==0) return(NULL)
        selectInput(ns("video_url_choice"), "Flux à transcrire", choices=urls)
      })
    })
    
    observeEvent(input$download_btn, {
      req(input$video_url_choice)
      url <- input$video_url_choice
      log_append(paste0("[🎬] Téléchargement audio : ", url))
      status("Téléchargement...")
      dest <- tempfile(fileext=".mp3")
      withProgress(message="Téléchargement audio", value=0, {
        cmd <- sprintf("ffmpeg -y -i '%s' -vn -q:a 0 -map a '%s'", url, dest)
        system(cmd, wait=TRUE)
        incProgress(1)
      })
      if (file.exists(dest) && file.info(dest)$size>0) {
        downloaded_file(dest)
        log_append("[✅] Audio téléchargé.")
        status("Audio prêt.")
        output$audio_preview <- renderUI({
          tags$audio(src=downloaded_file(), type="audio/mp3", controls=TRUE, style="width:100%;")
        })
      } else {
        log_append("[❌] Échec téléchargement.")
        status("Erreur de téléchargement.")
      }
    })
    
    observeEvent(input$submit, {
      updateActionButton(session, "submit", label="Transcription...", disabled=TRUE)
      status("Découpage audio...")
      src <- downloaded_file()
      req(src)
      
      seg_dir <- tempfile()
      dir.create(seg_dir)
      pattern <- file.path(seg_dir, "seg_%03d.mp3")
      cmd <- sprintf("ffmpeg -y -i '%s' -f segment -segment_time 1800 -c copy '%s'", src, pattern)
      system(cmd, wait=TRUE)
      segs <- list.files(seg_dir, pattern="seg_\\d{3}\\.mp3$", full.names=TRUE)
      nseg <- length(segs)
      log_append(paste0("[🔪] ", nseg, " segments créés."))
      
      whisper <- import("whisper")
      model_name <- as.character(input$whisper_model)
      log_append(paste0("[⚙️] Chargement modèle : ", model_name))
      model <- tryCatch(whisper$load_model(model_name), error=function(e) {
        log_append(paste0("[❌] Erreur chargement modèle : ", e$message))
        status("Erreur Whisper")
        return(NULL)
      })
      req(!is.null(model))
      
      status("Transcription en cours...")
      durations <- numeric()
      df <- data.frame(Segment=integer(), Duration=numeric())
      
      withProgress(message="Segments", value=0, {
        for (i in seq_along(segs)) {
          seg <- segs[i]
          t0 <- Sys.time()
          log_append(paste0("[🧠] Segment ", i, "/", nseg))
          incProgress(1/nseg)
          res <- tryCatch({
            r <- model$transcribe(seg)
            text <- r$text
            log_append(paste0("[📤] Extrait: ", substr(text, 1, 50)))
            text
          }, error=function(e) {
            log_append(paste0("[❌] Erreur transcription: ", e$message))
            NULL
          })
          dt <- as.numeric(difftime(Sys.time(), t0, units="secs"))
          durations <- c(durations, dt)
          df <- rbind(df, data.frame(Segment=i, Duration=dt))
          if (!is.null(res)) {
            transcript_text(paste0(isolate(transcript_text()), res, "\n"))
            rem <- round(mean(durations)*(nseg-i),1)
            log_append(paste0("[✅] Terminé en ", round(dt,1), "s, restant ~", rem, "s."))
          }
        }
        segment_times(df)
      })
      
      output$timeline_plot <- renderPlot({
        df <- segment_times()
        barplot(df$Duration, names.arg=paste0("S", df$Segment), main="Durée segments (s)", ylab="s")
      })
      
      removeModal()
      showModal(modalDialog(
        title="Résultat",
        size="l",
        easyClose=TRUE,
        tagList(
          verbatimTextOutput(ns("transcript")),
          downloadButton(ns("download"), "Exporter txt")
        ),
        footer=modalButton("Fermer")
      ))
      
      output$transcript <- renderText({
        txt <- transcript_text()
        if (nzchar(txt)) txt else "[⚠️] Aucun texte transcrit."
      })
      
      output$download <- downloadHandler(
        filename=function() paste0("transcript_", Sys.Date(), ".txt"),
        content=function(f) writeLines(transcript_text(), f, useBytes=TRUE)
      )
      
      updateActionButton(session, "submit", label="Transcrire", disabled=FALSE)
    })
  })
}
