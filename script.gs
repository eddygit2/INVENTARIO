function enviarHojaPorCorreo() {
  var html = HtmlService.createHtmlOutputFromFile('clave');
  SpreadsheetApp.getUi().showModalDialog(html, '-');
}

function verificarClave(clave) {
  var claveCorrecta = "D1638.";
  if (clave == claveCorrecta) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hojaMultiformato = spreadsheet.getSheetByName("Multiformato");
    var hojaDevolucion = spreadsheet.getSheetByName("Devolución");
    var hojaUsoInterno = spreadsheet.getSheetByName("Uso_interno");
    
    // Crear un archivo PDF de cada hoja
    var urlMultiformato = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=pdf&gid=" + hojaMultiformato.getSheetId() + "&portrait=false";
    var urlDevolucion = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=pdf&gid=" + hojaDevolucion.getSheetId() + "&portrait=false";
    var urlUsoInterno = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=pdf&gid=" + hojaUsoInterno.getSheetId() + "&portrait=false";
    
    var options = {
      "method": "GET",
      "headers": {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken()
      }
    };
    
    var responseMultiformato = UrlFetchApp.fetch(urlMultiformato, options);
    var responseDevolucion = UrlFetchApp.fetch(urlDevolucion, options);
    var responseUsoInterno = UrlFetchApp.fetch(urlUsoInterno, options);
    
    var pdfBlobMultiformato = responseMultiformato.getBlob().setName("Multiformato.pdf");
    var pdfBlobDevolucion = responseDevolucion.getBlob().setName("Devolución.pdf");
    var pdfBlobUsoInterno = responseUsoInterno.getBlob().setName("Uso_Interno.pdf");
    
    // Enviar los archivos PDF por correo electrónico
    var recipient = "gvalencia@tuti.com.ec"; // reemplaza con el correo electrónico del destinatario
    var subject = "220_Multiformato";
    var body = "220 - Multiformato";
    
    MailApp.sendEmail(recipient, subject, body, {
      attachments: [pdfBlobMultiformato, pdfBlobDevolucion, pdfBlobUsoInterno]
    });
    
    // Mostrar un mensaje de confirmación
    SpreadsheetApp.getUi().alert("Las hojas Multiformato, Devolución y Uso Interno han sido enviadas por correo electrónico con éxito.");
  
  } else {
    SpreadsheetApp.getUi().alert("Clave incorrecta. No se han enviado las hojas por correo electrónico.");
  }
}
