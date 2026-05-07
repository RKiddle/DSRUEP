function listStudentClassCodes() {
  // Fetch active courses where you are the teacher
  const response = Classroom.Courses.list({
    teacherId: 'me',
    courseStates: ['ACTIVE'] 
  });
  
  const courses = response.courses;
  
  if (!courses || courses.length === 0) {
    Logger.log('No active courses found.');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Student Class Codes");
  
  // Create a new sheet for the Class Codes if it doesn't exist, or clear it if it does
  if (!sheet) {
    sheet = ss.insertSheet("Student Class Codes");
  } else {
    sheet.clear();
  }
  
  // Add headers - Now including Course ID
  sheet.appendRow(["Course Name", "Course ID", "Student Class Code (Join Code)", "Section"]);
  
  // Prepare a 2D array of the course data
  const courseData = [];
  for (let i = 0; i < courses.length; i++) {
    
    // Fallback if the join code is disabled
    const joinCode = courses[i].enrollmentCode ? courses[i].enrollmentCode : "Disabled/Unavailable";
    
    courseData.push([
      courses[i].name, 
      courses[i].id,       // Added the Course ID here
      joinCode, 
      courses[i].section || "" 
    ]);
  }
  
  // Write the data to the sheet in one batch operation
  if (courseData.length > 0) {
    sheet.getRange(2, 1, courseData.length, courseData[0].length).setValues(courseData);
  }
  
  // Bold the header row and auto-resize columns for readability
  sheet.getRange("A1:D1").setFontWeight("bold");
  sheet.autoResizeColumns(1, 4);
}
