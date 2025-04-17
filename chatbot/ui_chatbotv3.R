ui <- fluidPage(
  useShinyjs(),
  actionButton("toggle_bpss_ui", "", style = "display: none;"),
  
  # 📦 HEAD: CSS + LIBS + JS
  tags$head(
    tags$title("BudgiBot"),
    tags$link(rel = "stylesheet", href = "https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&family=Merriweather:wght@700&display=swap"),
    tags$link(rel = "stylesheet", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css"),
    tags$style(HTML("
    body {
      font-family: 'Mariane', sans-serif;
      background-color: #F8F9FA;
      color: #333333;
    }
    
    .navbar {
      background-color: #0055A4 !important;
      border-bottom: 2px solid #FFD700;
    }
    
    .navbar-brand {
      color: #FFFFFF !important;
      font-size: 24px;
      font-weight: bold;
    }
    
    .chat-container {
      height: 500px;
      overflow-y: auto;
      border: 1px solid #DDDDDD;
      padding: 20px;
      background-color: #FFFFFF;
      border-radius: 10px;
      box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
      margin-bottom: 20px;
    }
    
    .chat-bubble-container {
      display: flex;
      flex-direction: column;
      margin-bottom: 15px;
      opacity: 0;
      transform: translateY(20px);
      animation: fadeInUp 0.5s forwards;
    }
    
    @keyframes fadeInUp {
      to { opacity: 1; transform: translateY(0); }
    }
    
    .bot-message-container {
      align-items: flex-start;
    }
    
    .user-message-container {
      align-items: flex-end;
    }
    
    .chat-sender {
      font-weight: bold;
      font-size: 12px;
      color: #666666;
      margin-bottom: 5px;
    }
    
    .chat-message {
      padding: 12px 16px;
      border-radius: 20px;
      font-size: 14px;
      max-width: 70%;
      word-wrap: break-word;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    .user-message {
      background-color: #0055A4;
      color: #FFFFFF;
    }
    
    .bot-message {
      background-color: #F8F9FA;
      color: #333333;
      border: 1px solid #DDDDDD;
    }
    
    .file-message {
      background-color: #FFD700;
      color: #000000;
      padding: 12px 16px;
      border-radius: 20px;
    }
    
    .quick-replies {
      margin-top: 10px;
    }
    
    .quick-replies button {
      margin: 5px;
      border-radius: 20px;
      background-color: #0055A4;
      color: #FFFFFF;
      border: none;
      padding: 8px 16px;
      font-size: 14px;
      cursor: pointer;
      transition: background-color 0.3s ease;
    }
    
    .quick-replies button:hover {
      background-color: #004080;
    }
    
    #drop_zone {
      border: 2px dashed #0055A4;
      padding: 20px;
      text-align: center;
      color: #0055A4;
      border-radius: 10px;
      margin-top: 10px;
      cursor: pointer;
      transition: background-color 0.3s ease;
    }
    
    #drop_zone.hover {
      background-color: rgba(0, 85, 164, 0.1);
    }
    
    .file-list-container {
      margin-top: 20px;
    }
    
    .file-list-container table {
      width: 100%;
      border-collapse: collapse;
    }
    
    .file-list-container th,
    .file-list-container td {
      border: 1px solid #DDDDDD;
      padding: 10px;
      text-align: left;
    }
    
    .file-list-container th {
      background-color: #F8F9FA;
    }
    
    .input-container {
      margin-top: 20px;
    }
    
    .input-container textarea {
      width: 100%;
      border-radius: 10px;
      border: 1px solid #DDDDDD;
      padding: 12px;
      font-size: 14px;
      resize: none;
    }
    
    .btn-container {
      display: flex;
      justify-content: space-between;
      margin-top: 10px;
    }
    
    .btn-container button {
      width: 48%;
      border-radius: 10px;
      padding: 12px;
      font-size: 14px;
      transition: background-color 0.3s ease;
    }
    
    .btn-primary {
      background-color: #0055A4;
      border: none;
      color: #FFFFFF;
    }
    
    .btn-primary:hover {
      background-color: #004080;
    }
    
    .btn-secondary {
      background-color: #FFD700;
      border: none;
      color: #000000;
    }
    
    .btn-secondary:hover {
      background-color: #E6C200;
    }
    
    /* Typing indicator */
    .typing-indicator-wrapper {
      position: relative;
      margin-top: -10px;
      margin-bottom: 10px;
      height: 40px;
      text-align: left;
      pointer-events: none;
    }
    
    .typing-indicator {
      display: inline-block;
      padding: 12px 16px;
      border-radius: 20px;
      background-color: #F8F9FA;
      color: #333333;
      font-size: 14px;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    .typing-indicator span {
      display: inline-block;
      width: 8px;
      height: 8px;
      margin: 0 2px;
      background: #666666;
      border-radius: 50%;
      animation: blink 1.4s infinite both;
    }
    
    .typing-indicator span:nth-child(1) { animation-delay: 0s; }
    .typing-indicator span:nth-child(2) { animation-delay: 0.2s; }
    .typing-indicator span:nth-child(3) { animation-delay: 0.4s; }
    
    @keyframes blink {
      0%, 80%, 100% { opacity: 0; }
      40% { opacity: 1; }
    }
    
    /* Fullscreen modal */
    .modal-content {
      height: 90vh;
      overflow-y: auto;
    }
    
    .modal-body {
      padding: 0 20px;
    }

    .modal-content {
      height: 90vh;
      overflow-y: auto;
      padding: 20px;
    }
    
    .modal-dialog {
      width: 98vw !important;
      max-width: none !important;
      margin: 0 auto;
    }
    
    #hot_table {
      width: 100% !important;
    }

  "))
  ),
    
   
    # JS scripts
    tags$script(HTML("
      // Auto-scroll chatbox
      Shiny.addCustomMessageHandler('scrollToBottom', function(message) {
        var chatBox = document.getElementById('chatBox');
        if (chatBox) {
          setTimeout(function() {
            chatBox.scrollTop = chatBox.scrollHeight;
          }, 100);
        }
      });

      // Real-time input listener
      Shiny.addCustomMessageHandler('initFastInput', function(message) {
        const inputField = document.getElementById('user_input');
        if (inputField) {
          inputField.addEventListener('input', function() {
            Shiny.setInputValue('user_input', inputField.value, {priority: 'event'});
          });
        }
      });

      // Submit on Enter key, debounce
      let enterPressed = false;
      $(document).on('keydown', function(e) {
        if (e.target.id === 'user_input') {
          if (e.key === 'Enter' && !e.shiftKey && !enterPressed) {
            e.preventDefault();
            enterPressed = true;
            $('#send_btn').prop('disabled', true);
            setTimeout(() => {
              $('#send_btn').click();
              enterPressed = false;
              $('#send_btn').prop('disabled', false);
            }, 150);
          }
        }
      });
    ")),
    tags$script(HTML("
      Shiny.addCustomMessageHandler('jsCode', function(message) {
        eval(message.code);
      });
    ")),
    
    tags$script(HTML("
      $(document).ready(function(){
        var dropZone = $('#drop_zone');

        dropZone.on('dragover', function(e) {
          e.preventDefault(); e.stopPropagation();
          dropZone.addClass('hover');
          dropZone.text('Relâchez pour déposer le fichier...');
        });

        dropZone.on('dragleave', function(e) {
          e.preventDefault(); e.stopPropagation();
          dropZone.removeClass('hover');
          dropZone.text('Déposez votre fichier ici (PDF, DOCX, TXT, MSG)');
        });

        dropZone.on('drop', function(e) {
          e.preventDefault(); e.stopPropagation();
          dropZone.removeClass('hover');
          dropZone.text('Fichier chargé !');

          var files = e.originalEvent.dataTransfer.files;
          if (files.length > 0) {
            var fileInput = $('#file_input');
            var dataTransfer = new DataTransfer();
            dataTransfer.items.add(files[0]);
            fileInput[0].files = dataTransfer.files;
            fileInput.trigger('change');
          }
        });
      });
    ")),
    
    tags$script(HTML("Shiny.addCustomMessageHandler('jsCode', function(message) {
      eval(message.code);
    });")),
  
  tags$script(HTML("
    Shiny.addCustomMessageHandler('appendToChat', function(msg) {
      var chatHistory = $('#chatBox');
      var message = '<div class=\"chat-bubble-container bot-message-container\">' +
                      '<div class=\"chat-sender\">BudgiBot</div>' +
                      '<div class=\"chat-message bot-message\">' + msg.content.replace(/\\n/g, '<br>') + '</div>' +
                    '</div>';
      chatHistory.append(message);
      chatHistory.scrollTop(chatHistory[0].scrollHeight);
    });
  ")),
  
  tags$script(HTML("
  $(document).ready(function(){
    $('#drop_zone_json').on('click', function() {
      $('#json_file').click();
    });
  });
  ")), 
  
  
  # 💬 Chat + 📊 Excel RH
  fluidRow(
    column(
      width = 6,
      style = "padding: 20px;",
      h3("BudgiBot"),
      div(id = "chatBox", class = "chat-container", style = "height: 600px; overflow-y: auto;",
          uiOutput("chat_history")
      ),
      div(class = "typing-indicator-wrapper", uiOutput("typing_indicator")),
      textAreaInput("user_input", label = NULL, placeholder = "Posez votre question ici...", rows = 3),
      div(class = "btn-container",
          actionButton("send_btn", "Envoyer", class = "btn-primary"),
          downloadButton("download_chat", "Télécharger la conversation", class = "btn-secondary")
      ),
      div(id = "drop_zone", "Déposez votre fichier ici (PDF, DOCX, TXT, MSG)"),
      div(class = "file-input-container",
          tags$label("Joindre un fichier", class = "btn btn-primary",
                     tags$input(id = "file_input", type = "file", style = "display: none;", accept = ".txt,.pdf,.docx,.msg"))
      ),
      div(class = "file-list-container", h4("Fichiers envoyés"), uiOutput("file_list")),
      actionButton("show_budget_modal", "Afficher les données budgétaires extraites", class = "btn-secondary"),
    ),
    
    column(
      width = 6,
      style = "padding: 20px;",
      h3("Mesures catégorielles"),
      mod_mesures_cat_ui("cat1"), 
      mod_json_helper_ui("json_helper")
    )
  )
)
