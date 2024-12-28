function sendEmailsDaysBefore() {
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1ShuHwTq-w8hbhUxRieXw5BOQpotuh5XZ2leZNa0TYi0/edit?gid=0#gid=0");

  var sheet = ss.getSheets()[0];

   if (!sheet) {
    Logger.log("A aba 'planilha' não foi encontrada.");
    return;  
  }

  var range = sheet.getDataRange().getValues();
  var timeNow = new Date();


  for (var i = 1; i < range.length; i++) {
    var programeData = new Date(range[i][0]);
    var email = range[i][1];
    var cco = range[i][2];
    var title = range[i][3];
    var status = range[i][5];

     var bodyWithLinks = 
      "Prezado(a) Cliente,<br><br>" +
      "Informamos que o seu certificado digital modelo A1 expirará nos próximos dias.<br><br>" +
      "Solicitamos que nos envie o novo certificado através do link abaixo.<br><br>" +
      '<a href="https://www.youtube.com" target="_blank" ' +
      'style="background-color: #007BFF; color: white; padding: 12px 20px; text-align: center; ' +
      'text-decoration: none; display: inline-block; border-radius: 5px; font-size: 16px;">' +
      "Envie aqui o seu certificado digital</a><br><br>" +
      "Lembramos que se o certificado digital não for atualizado, a SEFAZ não permitirá a emissão da NFCe.<br><br>" +
      "Caso tenha alguma dificuldade no envio, por favor, entre em contato com a nossa central de atendimento no WhatsApp.<br><br>" +
      "<em>*Este é um e-mail automático, por favor, não responda.</em><br><br>" +
      "Atenciosamente.<br>" + 
      "Equipe Pos Controle";



    if (programeData.getTime() <= timeNow.getTime() && status !== "Enviado" && status === "Pendente") {
      try {
        GmailApp.sendEmail(email, title, '', {
          htmlBody: bodyWithLinks,
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