/**
 * Envia emails com anexo do Google Drive para endereços armazenados na planilha atual
 */

var EMAIL_ENVIADO = 'EMAIL_ENVIADO'; // Essa constante será gravada na quarta coluna da planilha para indicar que um email já foi enviado para esse endereço

var startRow = 2; // Primeira linha dos dados na planilha
var numRows = 2; // Número de linhas a serem processadas
var startColumn = 1;// Primeira coluna dos dados na planilha
var numColumns = 4;// Número de colunas a serem processadas;

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // Pega as células que serão utilizadas da planilha
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  // Pega os valores de todos as colunas
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];

    var messageName = row[0]; // Primeira coluna
    var emailAddress = row[1]; // Segunda coluna
    var attachmentId = row[2] // Terceira coluna
    var emailSent = row[3]; // Quarta coluna

    var subject = 'Django Girls Brasília 2019 - Certificado'; // Assunto do emails
    var message = 'Olá! Com saudade da gente?\n' + messageName + '\nÉ com muito alegria que lhe enviamos o seu certificado de maravilhosidade.\nTe esperamos no Pós-Django Gilrs nessa quinta dia 29 às 19h na Aceleradora Cotidiano.\nAbraço,\nPyLadies'; // Mensagem do email
    if (emailSent != EMAIL_ENVIADO) { // Confere para duplicatas
      // Envia um email com um arquivo do Google Drive em um anexo de formato PDF
      var file = DriveApp.getFileById(attachmentId); //Id do arquivo
      
      GmailApp.sendEmail(emailAddress, subject, message, {
        attachments: [file.getAs(MimeType.PDF)],
        name: 'PyLadies DF'
      });
      sheet.getRange(startRow + i, 4).setValue(EMAIL_ENVIADO);
      // Certifica que a célula atual foi marcada em caso de interupção do script
      SpreadsheetApp.flush();
    }
  }
}
