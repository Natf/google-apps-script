function onFormSubmit(e) {
  const answersFolderId = '1OMBmsLhIpsQw-JyBgY0dYOmYi6VTl0yc';
  var values = e.namedValues;
  Logger.log(values);
  
  var name = values['Name'][0];
  var studentClass = values['Class'][0];
  var chapter = 'NTS Essay';
  var title = values['Title'][0];
  var hrTeacher = values['Home Room Teacher'][0];
  var essayContent = values['Essay'][0];
  Logger.log(essayContent);
  
  var spreadSheet = SpreadsheetApp.openById('1kCOGJGgW-xfGd0HyY1DAkX_aTFR62KSOuVAXxHMJIuM');
  SpreadsheetApp.setActiveSpreadsheet(spreadSheet);
  SpreadsheetApp.setActiveSheet(spreadSheet.getSheetByName('Form Responses 1'));
  
  var templateId = '11owJMHpt7H1g7HPMb_on9nXmDiPvckNrCy_tc17Z6m0';
  
  SpreadsheetApp.setActiveSheet(spreadSheet.getSheetByName('Form Responses 1'));

  var answersFolder = DriveApp.getFolderById(answersFolderId);
  var classFolders = answersFolder.getFolders();
  while (classFolders.hasNext()) {
    var folder = classFolders.next();
    if (folder.getName() == studentClass) {
      var documentId = DriveApp.getFileById(templateId).makeCopy(folder).getId();
      DriveApp.getFileById(documentId).setName(studentClass + ' - Assignment ' + chapter + ' - ' + name);
      var body = DocumentApp.openById(documentId).getBody();
      
      var wordCount = essayContent.trim().split(/\s+/).length;
      
      body.replaceText('<<Chapter>>', chapter);
      body.replaceText('<<Class>>', studentClass)
      body.replaceText('<<Name>>', name);
      body.replaceText('<<Title>>', title);
      body.replaceText('<<Essay>>', essayContent);
      body.replaceText('<<HrTeacher>>', hrTeacher);
      body.replaceText('<<Date>>', Utilities.formatDate(new Date(), "GMT+9", "yyyy/MM/dd"));
      body.replaceText('<<wordcount>>', wordCount);
      
      var slackIncomingWebhookUrl = 'https://hooks.slack.com/services/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';
      var postIcon = ":mailbox_with_mail:";
      var postUser = "Independent Essay Form";
      
      
      var payload = {
        "username": postUser,
        "icon_emoji": postIcon,
        "text": name + " ("+studentClass+") submitted a Number the Stars essay.\nView the Essay at: https://docs.google.com/document/d/" + documentId
      };
      
      var options = {
        'method': 'post',
        'payload': JSON.stringify(payload)
      };
      
      var response = UrlFetchApp.fetch(slackIncomingWebhookUrl, options);
      
      SpreadsheetApp.setActiveSheet(spreadSheet.getSheetByName('Submissions'));
      var durations = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions').getRange('A3:D500');
      var formulas = durations.getFormulas();
      var rows = ["A","B","C","D"];
        
      for (var i = 0; i < 4; i++) {
        for(var col = 0; col < 498; col ++) {
          if(i == 3) {
            formulas[col][i] = "='Form Responses 1'!H" + (col + 2);
          } else {
            formulas[col][i] = "='Form Responses 1'!" + rows[i] + (col + 2);
          }
        }
      }
      durations.setFormulas(formulas);
      SpreadsheetApp.flush();
      return;
    }
  }
}
