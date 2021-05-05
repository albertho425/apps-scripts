function doGet(request) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate();
}

/* @Include JavaScript and CSS Files */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/* @Process Form */
function processForm(formObject) {

  // type the link of your google sheet here
  var url = "https://docs.google.com/spreadsheets/d/1F6JC0gJuErhwYdgdCz5X3LwxWFEnYK0s8quzA9SuLYA/edit#gid=0";
  var ss = SpreadsheetApp.openByUrl(url);
  var ws = ss.getSheetByName("Sheet1");
  
  ws.appendRow([formObject.name,
                formObject.email,
                formObject.startDate,
                formObject.endDate,
                formObject.startTime,
                formObject.endTime,
                formObject.affiliation,
                formObject.department,
                formObject.eventName,
                formObject.comments
                ]);

                console.log(formObject.name, formObject.email, formObject.startDate, formObject.endDate);

                sendEmails(formObject);

}

/**
 * Sends emails 
 */
function sendEmails(formObject) {
  
  var email = formObject.email;
  var subject = formObject.affiliation + " - " + formObject.startDate;
  var body = "Confirmation message: " + formObject.startDate;
  
  GmailApp.sendEmail(email, subject, body);
  
}


