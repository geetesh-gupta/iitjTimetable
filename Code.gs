function myFunction() {
  var workbook   = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet2Data = sheet2Info(workbook);  
  var sheet3Data = sheet3Info(workbook, sheet2Data) ;
  var courseList = courses(sheet3Data);
  return timeTable(workbook, courseList);
}

function doGet(){
  return HtmlService
  .createTemplateFromFile('index')
  .evaluate();
}

function sheet2Info(workbook){
  var sheet2 = (workbook.getSheets())[1];
  var lastRow = sheet2.getDataRange().getLastRow();
  var dataRangeTable = sheet2.getRange(2, 1, lastRow, sheet2.getDataRange().getLastColumn()).getValues();
  var courseList = [];
  for (var row = 0 ; row < lastRow - 1; row = row + 1){    
    courseList.push(
      {
        'courseCode':      (dataRangeTable[row][1].toString()).replace(/\s/g, "").toUpperCase(),
        'courseName':      (dataRangeTable[row][2].toString()),
        'courseRoom':      (dataRangeTable[row][3].toString()),
        'courseInstructor':(dataRangeTable[row][4].toString()),
        'courseSlot':      (dataRangeTable[row][5].toString()),
      }
    )
  }
  return courseList;
}

function sheet3Info(workbook, sheet2Data){
  var sheet3 = (workbook.getSheets())[2];
  var lastRow = sheet3.getDataRange().getLastRow();
  var dataRangeTable =  sheet3.getRange(2, 1, lastRow, sheet3.getDataRange().getWidth()).getValues();
  var courseList = [];
  for (var row = 0; row < lastRow - 1; row = row + 1) {
    var courseFound = false;
    var courseCode = (dataRangeTable[row][0].toString()).replace(/\s/g, "").toUpperCase();
    for (var course = 0 ; course < sheet2Data.length; course = course + 1)  
      if (courseCode == sheet2Data[course]['courseCode']){
        courseList.push(
          {
            'program':          dataRangeTable[row][7] ,
            'sem':              dataRangeTable[row][8],
            'department':       dataRangeTable[row][6],
            'courseCode':       courseCode,
            'courseSlot':       sheet2Data[course]['courseSlot'],
            'courseRoom':       sheet2Data[course]['courseRoom'],
            'courseInstructor': sheet2Data[course]['courseInstructor'],
            'courseName':       sheet2Data[course]['courseName'],
          }
        )
        courseFound = true;
        break;
      }
    if(!courseFound)
      courseList.push(
        {
          'program':          dataRangeTable[row][7] ,
          'sem':              dataRangeTable[row][8],
          'department':       dataRangeTable[row][6],
          'courseCode':       courseCode,
          'courseSlot':       (dataRangeTable[row][2] == "-")?(" "):(dataRangeTable[row][2]),
          'courseRoom':       " ",
          'courseInstructor': " ",
          'courseName':       dataRangeTable[row][1],
        }
      )   
  }
  return courseList;
}

function courses(sheet3data){
  var courseList = [];
  courseList.push(
    {
    'department':sheet3data[0]['department'],
    'sem':       sheet3data[0]['sem'],
    'program':   sheet3data[0]['program'],
    'code':      [[sheet3data[0]['courseCode'], sheet3data[0]['courseSlot'], sheet3data[0]['courseRoom'], sheet3data[0]['courseInstructor'], sheet3data[0]['courseName']]],
    }
  )
  for (var row = 1; row < sheet3data.length - 1; row = row + 1){
    var count = false;
    for (var course = 0; course < courseList.length; course = course + 1)
      if (courseList[course]['department'] == sheet3data[row]['department'] && courseList[course]['sem'] == sheet3data[row]['sem'] && courseList[course]['program'] == sheet3data[row]['program']){
        courseList[course]['code'].push([sheet3data[row]['courseCode'], sheet3data[row]['courseSlot'], sheet3data[row]['courseRoom'], sheet3data[row]['courseInstructor'], sheet3data[row]['courseName']]);
        count = true;
        break;
      }
    if (!count)
      courseList.push(
        {
        'department':sheet3data[row]['department'],
        'sem':       sheet3data[row]['sem'],
        'program':   sheet3data[row]['program'],
        'code':      [[sheet3data[row]['courseCode'], sheet3data[row]['courseSlot'], sheet3data[row]['courseRoom'], sheet3data[row]['courseInstructor'], sheet3data[row]['courseName']]]
        }
      )
  }
  return courseList;
}


function timeTable(workbook, courseList){
  var sheet1         = (workbook.getSheets())[0];
  var lastCol        = sheet1.getDataRange().getLastColumn();
  var lastRow        = sheet1.getDataRange().getLastRow();
  var dataRangeTable = sheet1.getRange(2,1, lastRow, lastCol).getValues();
  var timeTable      = [];
  
  for (var course = 0; course < courseList.length; course = course + 1){
    timeTable.push([[courseList[course]['department'], courseList[course]['sem'], courseList[course]['program']]]);
    
    var secondRow = [];
    secondRow.push('');
    for (var col = 1; col < lastCol; col = col + 2)
      secondRow.push(dataRangeTable[0][col]);
    timeTable[course].push(secondRow);
    
    for (var row=1; row < lastRow; row = row + 1){
      var nextRow = [];
      nextRow.push(dataRangeTable[row][0]);
      for (var col = 1; col < lastCol; col = col + 1){
        if (dataRangeTable[row][col] != ''){   
          var yes = false;
          for (var branch = 0; branch < courseList[course]['code'].length; branch = branch + 1)
            if (dataRangeTable[row][col] == courseList[course]['code'][branch][1]){
              nextRow.push([courseList[course]['code'][branch][0],courseList[course]['code'][branch][2],courseList[course]['code'][branch][3],courseList[course]['code'][branch][4]])
              yes = true;
              break;
            } 
            else if ((dataRangeTable[row][col] == '#' + courseList[course]['code'][branch][1]) || (dataRangeTable[row][col] == courseList[course]['code'][branch][1] + '#')){
              if(row==7 || row==8)
                for (var m = 0; m < 5; m = m + 1)
                  if (m == 1)
                    nextRow.push([courseList[course]['code'][branch][0]+' Tut',courseList[course]['code'][branch][2],courseList[course]['code'][branch][3],courseList[course]['code'][branch][4]]);
                  else
                    nextRow.push([' ',' ',' ',' ']);
              else 
                nextRow.push([courseList[course]['code'][branch][0]+' Tut',courseList[course]['code'][branch][2],courseList[course]['code'][branch][3],courseList[course]['code'][branch][4]]);
              yes = true;
              break;
            } 
            else if (dataRangeTable[row][col] == 'Lunch Break'){
              nextRow.push('Lunch Break');
              yes = true;
              break;
            }
         if (!yes)
           nextRow.push([' ',' ',' ',' ']);
        }
      } 
      timeTable[course].push(nextRow);
    }
    timeTable[course].push([' ']);
  } 
  Logger.log(courseList);
  return timeTable;
}


