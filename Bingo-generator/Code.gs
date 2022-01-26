function generateBingo(){
  var items = parseInt(Sheets.Spreadsheets.Values.get('1Rdi7weXfZg6-qTgDUgY1iEUB3nksDvfdyMA_sVki7E8', 'B2:B2').values[0]);
  var freeSpace = ('TRUE' === Sheets.Spreadsheets.Values.get('1Rdi7weXfZg6-qTgDUgY1iEUB3nksDvfdyMA_sVki7E8', 'D4').values[0][0]);
    //Browser.msgBox("items " + items)
  
  if (items != 9 && items != 16 && items != 25) {
    Browser.msgBox("Please use 9, 16 or 25 words")
    return;
    
  }
  
  var squared = 0;
  
  if (items == 9) {
    squared = 3; 
  } else if (items == 16) {
    squared = 4; 
    freeSpace = false;
  }  else {
    squared = 5; 
  }
  
  var words = removeEmptyValues(Sheets.Spreadsheets.Values.get('1Rdi7weXfZg6-qTgDUgY1iEUB3nksDvfdyMA_sVki7E8', 'A2:'+'A50').values);
  
  if (words.length < items) {
    if (!(freeSpace && (words.length + 1) >= items)) {
      Browser.msgBox("You have entered less than "+items+" words")
      return;
    }
  }
  
  var numRequired = parseInt(Sheets.Spreadsheets.Values.get('1Rdi7weXfZg6-qTgDUgY1iEUB3nksDvfdyMA_sVki7E8', 'D2').values[0]);
  
  if (!(numRequired > 1)) {
    Browser.msgBox("Please enter more than 1 required")
    return;
  }
  
  var topic = Sheets.Spreadsheets.Values.get('1Rdi7weXfZg6-qTgDUgY1iEUB3nksDvfdyMA_sVki7E8', 'C2').values[0]
  var level = Sheets.Spreadsheets.Values.get('1Rdi7weXfZg6-qTgDUgY1iEUB3nksDvfdyMA_sVki7E8', 'B4').values[0]
  var term = Sheets.Spreadsheets.Values.get('1Rdi7weXfZg6-qTgDUgY1iEUB3nksDvfdyMA_sVki7E8', 'C4').values[0]

  
  var FileTemplateFileId = "1doaELh-62oR6o0ssg6PmtCpzZIETiRn4azO4g5bdcFw"
  var doc = DocumentApp.openById(FileTemplateFileId);
  var DocName = doc.getName();

  // Fetch entire table containing data
  //var data = sheet.getDataRange().getValues();
  var SerialLetterID = DriveApp.getFileById(FileTemplateFileId).makeCopy(topic + " - Bingo", DriveApp.getFolderById('1s0I6JMIeX6r2JTrRjtqSGzUPHQARI2O7')).getId();
  var docCopy = DocumentApp.openById(SerialLetterID);
  var body = docCopy.getBody();
  //var words = ['1', '2', '3', '4','5', '6', '7', '8','9', '10', '11', '12','13', '14', '15', '16']

  // Create a two-dimensional array containing the cell contents.
 
  // Build a table from the array.
  
  var table1 = body.findElement(DocumentApp.ElementType.TABLE).getElement()
  
  body.replaceText("#TITLE1", topic)
  body.replaceText("#TERM", term)
  body.replaceText("#LEVEL", level)
  
  //Create copy of the template document and open it
  let newWords = [];
  for (var i = 0; i < (numRequired-1); i++) {
    var tableCopy = table1.copy();
    newWords = JSON.parse(JSON.stringify(words));
    //docCopy.getBody().appendParagraph("")
    newWords = shuffle(newWords);
    if (freeSpace) {
      newWords.splice(Math.ceil(words.length/2), 0, ['Free Space']);
    }
    addBingoTable(newWords, tableCopy.getCell(1, 0), squared)
    docCopy.getBody().appendTable(tableCopy);
  }
  
  addBingoTable(newWords, table1.getCell(1, 0), squared, freeSpace)
  
  docCopy.saveAndClose();
  
  msgBoxWithLink('Click here to view the bingo sheet.','https://docs.google.com/document/d/'+docCopy.getId(),'This will take you to the bingo sheet.');
}

function msgBoxWithLink(msg,link,desc) {
  var link=link;
  var desc=desc;
  var msg=msg;
  var html=Utilities.formatString('<style>input{margin: 5px 0;}</style><a href="%s" target="_blank">Click here</a><br /><input type="button" value="Close" onClick="google.script.host.close();" />',link);
  var userInterface=HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModelessDialog(userInterface, "View the bingo sheet here");
}

function addBingoTable(words, tableCell, elementsPerSubArray, freeSpace) {
  var cells = listToMatrix(words, elementsPerSubArray);
  
  var cellHeight = 280/elementsPerSubArray;
  
  var tableAnswers = tableCell.appendTable(cells);
  
  var style = {};
  var rowStyle = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  style[DocumentApp.Attribute.VERTICAL_ALIGNMENT] =
    DocumentApp.VerticalAlignment.TOP;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  style[DocumentApp.Attribute.FONT_SIZE] = 8;
  style[DocumentApp.Attribute.BOLD] = true;
  
  rowStyle[DocumentApp.Attribute.MINIMUM_HEIGHT] = cellHeight;
  rowStyle[DocumentApp.Attribute.WIDTH] = 60;
  
  // Apply the custom style.
  tableAnswers.setAttributes(style);
  var cellStyle = {}
  var cell = tableAnswers.findElement(DocumentApp.ElementType.TABLE_CELL)
  
  cellStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  cellStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  
  cell.getElement().setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  cell.getElement().getChild(0).asParagraph().setAttributes(cellStyle);
  
  for (var row = 0; row < tableAnswers.getNumRows(); row++) {
    var theRow = tableAnswers.getRow(row);
    theRow.setAttributes(rowStyle)
    for (var i = 1; i <= tableAnswers.getNumChildren(); i ++) {
      cell = theRow.findElement(DocumentApp.ElementType.TABLE_CELL, cell)
      if (cell != null) {
        cell.getElement().setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        cell.getElement().getChild(0).asParagraph().setAttributes(cellStyle);
        if (i/(tableAnswers.getNumChildren()-1) == 0.5 && (row/(tableAnswers.getNumRows()-1) == 0.5) && freeSpace) {
          cell.getElement().setBackgroundColor("#999999");
        }
      }
    }
  }
}

function removeEmptyValues(list) {
  var arrayV = [], i;

    for (i = 0; i < list.length; i++) {
      if (list[i] != undefined && list[i] != "") {
         arrayV[i] = list[i];
      }
    }

    return arrayV;
}

function listToMatrix(list, elementsPerSubArray) {
    var matrix = [], i, k;
  var noWords = elementsPerSubArray*elementsPerSubArray;

    for (i = 0, k = -1; i < noWords; i++) {
        if (i % elementsPerSubArray === 0) {
            k++;
            matrix[k] = [];
        }

        matrix[k].push(list[i][0]);
    }

    return matrix;
}

function shuffle(array) {
  var currentIndex = array.length, temporaryValue, randomIndex;

  // While there remain elements to shuffle...
  while (0 !== currentIndex) {

    // Pick a remaining element...
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex -= 1;

    // And swap it with the current element.
    temporaryValue = array[currentIndex];
    array[currentIndex] = array[randomIndex];
    array[randomIndex] = temporaryValue;
  }

  return array;
}
