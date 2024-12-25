function enviarEmails() {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1ShuHwTq-w8hbhUxRieXw5BOQpotuh5XZ2leZNa0TYi0/edit?gid=0#gid=0");
  
    var sheet = ss.getSheets()[0];
  
     if (!sheet) {
      Logger.log("A aba 'planilha' não foi encontrada.");
      return; 
    }
  
    var range = sheet.getDataRange().getValues();
    var time = new Date(); 
  
    for (var i = 1; i < range.length; i++) {
      var programeData = new Date(range[i][0]);
      var email = range[i][1];
      var cco = range[i][2];
      var title = range[i][3];
      var bodyOfText = range[i][4];
      var status = range[i][5];
  
        var corpoComLinks = 
        "Lembramos que se o certificado digital não for atualizado, a SEFAZ não permitirá a emissão da NFCe.<br><br>" +
        '<a href="https://www.youtube.com" target="_blank" ' +
        'style="background-color: #007BFF; color: white; padding: 12px 20px; text-align: center; ' +
        'text-decoration: none; display: inline-block; border-radius: 5px; font-size: 16px;">' +
        'Clique aqui para enviar o seu certificado digital</a>';
  
  
      if (programeData.getTime() <= time.getTime() && status !== "Enviado") {
        try {
          GmailApp.sendEmail(email, title, '', {
            htmlBody: corpoComLinks,
            bcc: cco
          });
          
          sheet.getRange(i + 1, 6).setValue("Enviado");
          console.log(`E-mail enviado para ${email} com sucesso!`)
        } catch(e) {
          console.error(`Erro ao enviar e-mail para ${email}: ${e.message}`)
        }
      }
    }
  }
  
  function trigger() {
    ScriptApp.newTrigger('enviarEmails')  // Função que será executada
      .timeBased()
      .everyMinutes(1)  // Executa a cada 1 minuto
      .create();
  }