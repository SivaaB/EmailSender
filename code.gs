function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('✉️ Send Email ✉️')
      .addItem('Send to first 100', 'sendFormToAll')
      .addToUi();
}

function sendPDFForm()
{
  var row = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  sendEmailWithAttachment(row);
}

function sendEmailWithAttachment(row)
{
  var filename= 'illumm_Sponsorship_Prospectus.pdf';
  
  var file = DriveApp.getFilesByName(filename);
  
  if (!file.hasNext()) 
  {
    console.error("Could not open file "+filename);
    return;
  }
  
  var client = getClientInfo(row);
  
  var template = HtmlService
      .createTemplateFromFile('template');
  template.client = client;
  var message = template.evaluate().getContent();
  
  
  MailApp.sendEmail({
    to: client.email,
    subject: "Invitation to Illumm! ",
    htmlBody: message,
    attachments: [file.next().getAs(MimeType.PDF)]});
  
}

function getClientInfo(row)
{
   var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1');
   
   var values = sheet.getRange(row,1,row,4).getValues();
   var rec = values[0];
  
  var client = 
      {
        name: rec[0],
        email: rec[1],
      };
  client.name = client.name;
  return client;
}

function sendFormToAll()
{
   var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1');
  
   var last_row = sheet.getDataRange().getLastRow();
  
   for(var row=2; row <= last_row; row++)
   {
     sendEmailWithAttachment(row);
     sheet.getRange(row,3).setValue("email sent");
   }
}
