# appscript-mailmerge-pdf
[Sample Spreadsheet](https://docs.google.com/spreadsheets/d/1MGsFF0FLtdwyxBeQQ2fF-4wsMtMYlY3vFQMLN32PeKI/edit?usp=sharing)

#### This is a simple app script project which convert data from Google Spreadsheet and `mail-merge` with template with Google Slide and generate PDF.

#### 1. This is the blank template (Google Slide)  
![Template](https://github.com/hawkng/appscript-mailmerge-pdf/blob/main/cert-template1.png)

#### 2. PDF output  
![PDF output](https://github.com/hawkng/appscript-mailmerge-pdf/blob/main/cert-output.png)

#### 3. App Script
```
function createAttendanceCert() {

  const ROOT_FOLDER           = '<YOUR ROOT FOLDER ID>';
  const TEMPLATE_SLIDE_ID     = '<TEMPLATE GOOGLE SLIDE FILE ID>';
  const ATTENDANCE_SHEET_ID   = '<ATTENDANCE SPREADSHEET FILE ID>'; 
  const COL_NAME              = 1;
  const COL_EMAIL             = 2;
    
  var userTimeZone            = CalendarApp.getDefaultCalendar().getTimeZone();  
  var attendanceSpreadSheet   = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
  var attendanceDataSheet     = attendanceSpreadSheet.getSheets()[0]; //get the first worksheet
  var attendanceData          = attendanceDataSheet.getRange(6,1, attendanceDataSheet.getLastRow(), attendanceDataSheet.getLastColumn()).getValues(); //(row, column, numRows, numColumns) 
  
  var companyName  = attendanceDataSheet.getRange("C2").getValue();
  var trainingDate = Utilities.formatDate(attendanceDataSheet.getRange("C1").getValue(), userTimeZone,  'dd MMMM yyyy'); //date, timeZone, format;  
  var courseName   = attendanceDataSheet.getRange("C3").getValue();  
  var certFolder   = DriveApp.getFolderById(ROOT_FOLDER).createFolder(companyName);
  
  
  //Process each rows
  for (var rowIdx in attendanceData){
     
     var attendeeName = attendanceData[rowIdx][COL_NAME];     
     var newSlideName = "Certificate of Attendance - " + attendeeName;
     var newSlideFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID).makeCopy(newSlideName, certFolder);     
     var newSlide     = SlidesApp.openById(newSlideFile.getId()); 
     var newTemplate  = newSlide.getSlides()[0];
     
     newTemplate.replaceAllText("{{ATTENDEE_NAME}}",attendeeName);    
     newTemplate.replaceAllText("{{COURSE_NAME}}", courseName);
     newTemplate.replaceAllText("{{TRAINING_DATE}}", trainingDate);
     
     newSlide.saveAndClose();
     
     var blob    = newSlideFile.getBlob();
     blob.setName(newSlideName + ".pdf");       
     var pdfFile = DriveApp.createFile(blob);
     pdfFile.makeCopy(certFolder);
     pdfFile.setTrashed(true);
     newSlideFile.setTrashed(true);     
  }
  
 } 
 ```
