function isTimeUp_(start) {
  var now = new Date();
  return now.getTime() - start.getTime() > 300000; // 5 minutes
}

function setTitleOnTable(table, studentData, student) {
  var titleCell = table.getCell(0,0);
  var completionDate = studentData.completionDate;
  titleCell.setText('Number the Stars Quiz - '+student+' - '+completionDate);
}

function getClassName(name, theirclass) {
  const regexOnlyKorean = /[a-z0-9\s\-,.\(\)\/\&]/gmi
  name = theirclass+" "+name.replace(regexOnlyKorean, "");
  return name;
}

function getDataFromQuiz(quizSheet) {
  var data = {};

  var names = quizSheet.getRange(2, 3, quizSheet.getLastRow() - 1);
  var classes = quizSheet.getRange(2, 4, quizSheet.getLastRow() - 1);
  var answers = quizSheet.getRange(2, 6, quizSheet.getLastRow() - 1, 11);
  var scores = quizSheet.getRange(2, 2, quizSheet.getLastRow() - 1, 11);
  var completionDates = quizSheet.getRange(2, 1, quizSheet.getLastRow() - 1, 11);

  for (var nameNo = 0; nameNo < names.getValues().length; nameNo++) {
    var completionDate = completionDates.getValues()[nameNo][0].toString();
    var months = {
      Jan:'1',
      Feb:'2',
      Mar:'3',
      Apr:'4',
      May:'5',
      Jun:'6',
      Jul:'7',
      Aug:'8',
      Sep:'9',
      Oct:'10',
      Nov:'11',
      Dec:'12'
    }
    var stringCompletionDate = completionDate.substring(8,10)+'/'+months[completionDate.substring(4,7)]+'/'+completionDate.substring(11,15);

    var classname = {
      name : names.getValues()[nameNo][0],
      class : classes.getValues()[nameNo][0],
      answers : answers.getValues()[nameNo],
      score: scores.getValues()[nameNo][0],
      completionDate: stringCompletionDate,
      row : nameNo + 2
    };
    data[getClassName(names.getValues()[nameNo][0], classes.getValues()[nameNo][0])] = classname;
  }
  return data;
}

function getUniqueClassNames(data) {
  var uniqueClassNames = [];
  var allClassNames = [];
  for (var quizNo = 0; quizNo < data.length; quizNo++) {
    var keys = Object.keys(data[quizNo]);
    for (var keyNo = 0; keyNo < keys.length; keyNo++) {
      allClassNames.push(keys[keyNo]);
    }
  }

  for (var nameNo = 0; nameNo < allClassNames.length; nameNo++) {
    var name = allClassNames[nameNo];
    if (!uniqueClassNames.includes(name)) {
      uniqueClassNames.push(name);
    }
  }

  return uniqueClassNames;
}

function setStudentScoreAndAnswers(table, classnameData, correctAnswers) {
  for (var questionNo = 0; questionNo < 10; questionNo++) {
    var cell = table.getCell(2 + questionNo,3);
    cell.setText(classnameData.answers[questionNo]);
    if (correctAnswers[questionNo] == classnameData.answers[questionNo]) {
      cell.setBackgroundColor('#b6d7a8');
    } else {
      cell.setBackgroundColor('#ea9999');
    }
  }
  var cellWrittenAnswer = table.getCell(12,1);
  cellWrittenAnswer.setText(classnameData.answers[10]);

  var score = table.getCell(1,0);
  score.setText('Score: ' + classnameData.score+'/10');
  var style = {};
  style[DocumentApp.Attribute.BOLD] = true;
  score.setAttributes(style);
}

function createStudentAnswersDocForStudent(document, allQuizzesData, student) {
  var quizNo = 0;
  var date = new Date();
  var timeStamp = date.getTime();
  document.setName(' Number the Stars Quiz Answers - '+student);
  var body = document.getBody();
  var tables = body.getTables();

  for (quizNo; quizNo < allQuizzesData.length; quizNo++) {
    var classnameData;
    var table = tables[((quizNo*2) + 1)];
    var titleTable = tables[quizNo*2];
    if (allQuizzesData[quizNo][student] != undefined) {
      classnameData = allQuizzesData[quizNo][student];
    } else {
      classnameData = {answers : ['','','','','','','','','','',''], score:'0', completionDate: 'Not Completed'};
    }

    setTitleOnTable(titleTable, classnameData, student);
    setStudentScoreAndAnswers(table, classnameData, allQuizzesData[quizNo]['CORRECT '].answers);
  }
}

function createStudentAnswersPDFForStudent(allQuizzesData, student) {
  const templateId = '19G_aN5VM5u3ofjjNaBS-Q5RCEyTBRPLc6fVBmjYR7Rk';
  const answersFolderId = '1UaaeGwq2gKt-tINjNLMomqy64Te1ms-c';
  var documentId = '1kxzLn_8zFSKugof6-iOsBP5lxoB5qdmCr_KAQ6OXpaQ';
  var document = DocumentApp.openById(documentId);

  createStudentAnswersDocForStudent(document, allQuizzesData, student);
  document.saveAndClose();

  var studentClass = student.match(/[0-9a-z\-]+/gmi)[0];

  var answersFolder = DriveApp.getFolderById(answersFolderId);
  var classFolders = answersFolder.getFolders();
  while (classFolders.hasNext()) {
    var folder = classFolders.next();
    if (folder.getName() == studentClass) {
      document = DocumentApp.openById(documentId);
      var docblob = document.getAs('application/pdf');
      docblob.setName(document.getName() + ".pdf");
      folder.createFile(docblob);
      return;
    }
  }
}

function exportAllQuizzestoPDF() {
  var start = new Date();
  var spreadSheet = SpreadsheetApp.getActive();
  var allQuizSpreadsheets = [
    'Quiz Chapter 1&2',
    'Quiz Chapter 3&4',
    'Quiz Chapter 5&6',
    'Quiz Chapter 7&8',
    'Quiz Chapter 9&10',
    'Quiz Chapter 11&12',
    'Quiz Chapter 13&14',
    'Quiz Chapter 15&16&17'
  ];
  var allQuizzesData = [];

  // get all the names submitted to every quiz as well as a matching indexed list of their classes
  for (var quizNo = 0; quizNo < allQuizSpreadsheets.length; quizNo++) {
    allQuizzesData.push(getDataFromQuiz(spreadSheet.getSheetByName(allQuizSpreadsheets[quizNo])));
  }

  var studentUniqueNames = getUniqueClassNames(allQuizzesData);
  studentUniqueNames.sort();
  Logger.log((studentUniqueNames.length-2) + ' students found\n'+studentUniqueNames.sort());
  
  var studentNo = 0;
  
  // live
  for (;studentNo < studentUniqueNames.length; studentNo++) {
    Logger.log('Generating answer PDF for StudentNo '+studentNo+' - ' + studentUniqueNames[studentNo]);
    createStudentAnswersPDFForStudent(allQuizzesData, studentUniqueNames[studentNo]);
    if (isTimeUp_(start)) {
      Logger.log("Time up continue from studentNo = " + (studentNo + 1));
      break;
    }
  }
}
