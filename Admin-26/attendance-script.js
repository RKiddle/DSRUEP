function generateAttendanceSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define the start and end dates (Assuming current year)
  var year = new Date().getFullYear();
  var startDate = new Date(year, 5, 3); // Month is 0-indexed (5 = June, 3rd)
  var endDate = new Date(year, 9, 1);   // (9 = October, 1st)
  
  // Generate date headers (skipping weekends)
  var dateHeaders = ['Student Name'];
  var currentDate = new Date(startDate);
  
  while (currentDate <= endDate) {
    var dayOfWeek = currentDate.getDay();
    // Skip Sunday (0) and Saturday (6)
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd-MMM");
      dateHeaders.push(formattedDate);
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  // Calculate total number of days (excluding the 'Student Name' header)
  var numDays = dateHeaders.length - 1; 
  
  // Append the new Grade column header
  dateHeaders.push('Grade (/10)');

  try {
    // Fetch active courses where you are the teacher
    var response = Classroom.Courses.list({
      teacherId: 'me',
      courseStates: ['ACTIVE']
    });
    var courses = response.courses;
    
    if (!courses || courses.length === 0) {
      SpreadsheetApp.getUi().alert('No active courses found in your Google Classroom.');
      return;
    }

    // Loop through each course to create its sheet
    for (var i = 0; i < courses.length; i++) {
      var course = courses[i];
      var safeName = course.name.substring(0, 95); 
      var sheetName = safeName + '-ATT';
      
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        sheet.clear(); // Clears existing sheet to prevent overlapping data
      }
      
      // 1. Set the Date Headers (Row 1)
      sheet.getRange(1, 1, 1, dateHeaders.length).setValues([dateHeaders]);
      sheet.getRange(1, 1, 1, dateHeaders.length).setFontWeight("bold");
      
      // 2. Fetch Students for the course
      var students = [];
      var pageToken = null;
      do {
        var studentResponse = Classroom.Courses.Students.list(course.id, {
          pageToken: pageToken
        });
        if (studentResponse.students) {
          for (var j = 0; j < studentResponse.students.length; j++) {
            students.push([studentResponse.students[j].profile.name.fullName]);
          }
        }
        pageToken = studentResponse.nextPageToken;
      } while (pageToken);
      
      // 3. Populate Students, Checkboxes, and Grades
      if (students.length > 0) {
        // Sort students alphabetically
        students.sort(function(a, b) {
          return a[0].localeCompare(b[0]);
        });
        
        // Write student names to column A
        sheet.getRange(2, 1, students.length, 1).setValues(students);
        
        // Insert checkboxes under the date columns ONLY (leaving the grade column alone)
        sheet.getRange(2, 2, students.length, numDays).insertCheckboxes();
        
        // Apply the grading formula to the last column
        // This formula counts the TRUE checkboxes in the row, divides by total days, multiplies by 10, and rounds to 1 decimal.
        var gradeFormula = "=IFERROR(ROUND((COUNTIF(R[0]C2:R[0]C[-1], TRUE) / " + numDays + ") * 10, 1), 0)";
        sheet.getRange(2, dateHeaders.length, students.length, 1).setFormulaR1C1(gradeFormula);
      }
      
      // Freeze the top row and first column, then auto-resize for neatness
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(1);
      sheet.autoResizeColumns(1, dateHeaders.length);
      
      // Optional: Highlight the Grade column slightly to make it stand out
      sheet.getRange(1, dateHeaders.length, students.length + 1, 1).setBackground("#f3f3f3");
    }
    
    SpreadsheetApp.getUi().alert('Success! Attendance sheets with Grades generated.');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}
