# Thème moderne Flatly
theme <- bs_theme(
  version    = 5,
  bootswatch = "flatly",
  primary    = "#1ABC9C",    # Turquoise pour l'utilisateur
  secondary  = "#4169E1"     # Bleu roi pour le bot
)

ui <- fluidPage(
  theme = theme,
  useShinyjs(),
  
  tags$head(
    tags$title("BudgiBot"),
    # CSS pour chat, bulles, quick replies et modals
    tags$style(HTML(
      ".chat-container {
   display: flex;
   flex-direction: column;
   height: 600px;
   overflow-y: auto;
   padding: 1rem;
   background: #fff;
   border: 1px solid #ddd;
   border-radius: 0.5rem;
}
@keyframes fadeInUp { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
.chat-bubble-container {
   display: flex;
   flex-direction: column;
   margin-bottom: 0.75rem;
   animation: fadeInUp 0.3s forwards;
}
.bot-message-container { align-items: flex-start; }
.user-message-container { align-items: flex-end; }
.chat-sender { font-size: 0.85rem; color: #666; margin-bottom: 0.3rem; }
.chat-message {
   max-width: 80%;
   padding: 1rem 1.25rem;
   border-radius: 1rem;
   font-size: 1.1rem;
   line-height: 1.4;
   word-break: break-word;
   border: 1px solid transparent;
}
.bot-message .chat-message {
   background: var(--bs-secondary) !important;
   color: #FFFFFF;
   border-color: var(--bs-secondary);
}
.user-message .chat-message {
   background: var(--bs-primary) !important;
   color: #FFFFFF;
   border-color: var(--bs-primary);
}
.quick-replies {
   margin-top: 0.5rem;
   display: flex;
   gap: 0.5rem;
   flex-wrap: wrap;
}
.quick-replies button {
   padding: 0.5rem 1rem;
   font-size: 0.9rem;
   border-radius: 1rem;
   border: 1px solid var(--bs-secondary);
   background: var(--bs-secondary) !important;
   color: #FFF;
   transition: background 0.2s;
}
.quick-replies button:hover {
   background: #355C7D !important;
}
.modal-dialog {
   max-width: 95vw !important;
   margin: 1rem auto;
}
.modal-content {
   width: 100%;
   height: 90vh;
}
"    )),
    # JS pour Enter, quick replies, scroll auto et drop
    tags$script(HTML(
      "$(function(){
  // Envoi avec Enter (Enter -> envoyer, Shift+Enter -> saut)
  $('#user_input').on('keydown', function(e){
    if(e.key === 'Enter' && !e.shiftKey){
      e.preventDefault();
      $('#send_btn').click();
    }
  });
  // Quick replies clic -> remplit et envoie
  $(document).on('click', '.quick-replies button', function(){
    var txt = $(this).text();
    $('#user_input').val(txt);
    $('#send_btn').click();
  });
    // Permet de forcer un scroll vers le bas après un update UI classique
    Shiny.addCustomMessageHandler('scrollToBottom', function(message) {
      var chatBox = document.getElementById('chatBox');
      if (chatBox) {
        setTimeout(function() {
          chatBox.scrollTop = chatBox.scrollHeight;
        }, 50);
      }
    });
  // Drop zone interactions
  $(document).on('dragover', '#drop_zone', function(e){ e.preventDefault(); $(this).addClass('hover').text('Relâchez pour déposer'); });
  $(document).on('dragleave', '#drop_zone', function(e){ e.preventDefault(); $(this).removeClass('hover').text('Glissez un fichier ici (PDF, DOCX, TXT, MSG)'); });
  $(document).on('drop', '#drop_zone', function(e){
    e.preventDefault();
    var dz = $(this);
    dz.removeClass('hover').text('Fichier chargé !');
    var files = e.originalEvent.dataTransfer.files;
    if(files.length){ var fi = $('#file_input'), dt = new DataTransfer(); dt.items.add(files[0]); fi[0].files = dt.files; fi.trigger('change'); }
  });
});"
    ))
  ),
  
  fluidRow(
    # Colonne Chat
    column(width = 6, style = "padding:20px; display:flex; flex-direction:column;",
           div(style = "display:flex; align-items:center; margin-bottom:1rem;",
               tags$img(src = "logo.png", height = "40px", alt = "Logo"),
               tags$h3("BudgiBot", style = "margin-left:0.5rem; margin-bottom:0;")
           ),
           div(id = "chatBox", class = "chat-container", uiOutput("chat_history")),
           div(class = "mt-2", uiOutput("typing_indicator")),
           textAreaInput("user_input", NULL,
                         placeholder = "Écrivez un message… (Shift+Enter pour saut)",
                         rows = 2, width = "100%"
           ),
           div(class = "d-flex mt-2",
               actionButton("send_btn", label = NULL, icon = icon("paper-plane"), class = "btn btn-primary me-2"),
               downloadButton("download_chat", "Télécharger", class = "btn btn-secondary")
           ),
           div(id = "drop_zone", class = "mt-3 p-3 text-center border rounded text-secondary",
               "Glissez un fichier ici (PDF, DOCX, TXT, MSG)"
           ),
           fileInput("file_input", NULL,
                     buttonLabel = "Joindre…",
                     accept = c('.pdf','.docx','.txt','.msg'),
                     width = "100%"
           ),
           div(class = "file-list-container mt-2", uiOutput("file_list"))
    ),
    # Colonne Modules
    column(width = 6, style = "padding:20px;",
           tags$h3("Outils Budgétaires"),
           div(class = "mb-3 d-flex",
               actionButton("toggle_bpss_ui", "BPSS Excel", class = "btn btn-outline-primary flex-fill me-2"),
               actionButton("show_video_modal", "Vidéo transcription", class = "btn btn-outline-primary flex-fill")
           ),
           tags$div(class = "accordion", id = "modulesAccordion",
                    # Mesures Catégorie ouvertes par défaut
                    tags$div(class = "accordion-item",
                             tags$h2(class = "accordion-header", id = "headingCat"),
                             tags$button(class = "accordion-button show", type = "button",
                                         `data-bs-toggle` = "collapse", `data-bs-target` = "#collapseCat",
                                         `aria-expanded` = "true", `aria-controls` = "collapseCat",
                                         "Mesures catégorielles"
                             ),
                             tags$div(id = "collapseCat", class = "accordion-collapse collapse show",
                                      `aria-labelledby` = "headingCat", `data-bs-parent` = "#modulesAccordion",
                                      tags$div(class = "accordion-body", mod_mesures_cat_ui("cat1"))
                             )
                    ),
                    tags$div(class = "accordion-item",
                             tags$h2(class = "accordion-header", id = "headingJson"),
                             tags$button(class = "accordion-button show", type = "button",
                                         `data-bs-toggle` = "collapse", `data-bs-target` = "#collapseJson",
                                         `aria-expanded` = "true", `aria-controls` = "collapseJson",
                                         "JSON Helper"
                             ),
                             tags$div(id = "collapseJson", class = "accordion-collapse collapse show",
                                      `aria-labelledby` = "headingJson", `data-bs-parent` = "#modulesAccordion",
                                      tags$div(class = "accordion-body", mod_json_helper_ui("json_helper"))
                             )
                    )
           )
    )
  )
)

# Note: placez logo.png dans www/ pour qu'il s'affiche correctement.