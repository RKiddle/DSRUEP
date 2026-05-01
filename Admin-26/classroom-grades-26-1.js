function generateGradebook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set your date range here
  const startDate = new Date("2026-04-01T00:00:00Z");
  const endDate = new Date("2026-10-01T23:59:59Z");

  try {
    // 1. Get all active courses
    const courseResponse = Classroom.Courses.list({ courseStates: ['ACTIVE'] });
    const courses = courseResponse.courses;

    if (!courses || courses.length === 0) {
      Logger.log("No active courses found.");
      return;
    }

    // 2. Loop through each course
    for (let i = 0; i < courses.length; i++) {
      let course = courses[i];
      let courseId = course.id;
      
      // Create a new tab for the class (or clear it if it already exists)
      let sheetName = course.name.substring(0, 30); // Sheets limits names to 30 chars
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        sheet.clear();
      }

      // 3. Get Grade Categories for this course
      let categories = {};
      let categoryNames = [];
      if (course.gradebookSettings && course.gradebookSettings.gradeCategories) {
        course.gradebookSettings.gradeCategories.forEach(cat => {
          categories[cat.id] = cat.name;
          // CHANGED: Update header to show it is out of 10
          categoryNames.push(cat.name + " (Avg /10)");
        });
      }

      // 4. Fetch CourseWork (Assignments) and Filter by Date
      let workResponse = Classroom.Courses.CourseWork.list(courseId);
      let allWork = workResponse.courseWork || [];
      
      let validWork = allWork.filter(work => {
        let createdDate = new Date(work.creationTime);
        return createdDate >= startDate && createdDate <= endDate;
      });

      // Reverse to show chronological order (oldest to newest left to right)
      validWork.reverse(); 

      // 5. Fetch Students
      let studentResponse = Classroom.Courses.Students.list(courseId);
      let students = studentResponse.students || [];

      // 6. Fetch ALL Submissions for this course (using the '-' wildcard)
      let submissions = [];
      let pageToken = null;
      do {
        let subResponse = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, '-', {pageToken: pageToken});
        if (subResponse.studentSubmissions) {
          submissions = submissions.concat(subResponse.studentSubmissions);
        }
        pageToken = subResponse.nextPageToken;
      } while (pageToken);

      // Organize submissions: map[studentId][courseWorkId] = assignedGrade
      let gradesMap = {};
      submissions.forEach(sub => {
        if (!gradesMap[sub.userId]) gradesMap[sub.userId] = {};
        // Only record grades that have actually been assigned by the teacher
        if (sub.assignedGrade !== undefined) {
          gradesMap[sub.userId][sub.courseWorkId] = sub.assignedGrade;
        }
      });

      // 7. Build the Header Row
      let headerRow = ["Student Name"];
      headerRow = headerRow.concat(categoryNames); // Add Category Columns
      
      validWork.forEach(work => {
        headerRow.push(work.title + `\n(Max: ${work.maxPoints || 100})`);
      });
      sheet.appendRow(headerRow);
      sheet.getRange(1, 1, 1, headerRow.length).setFontWeight("bold").setWrap(true);

      // 8. Calculate Averages and Build Student Rows
      if (students.length > 0) {
        let rowData = [];
        
        students.forEach(student => {
          let studentRow = [student.profile.name.fullName];
          let studentGrades = gradesMap[student.userId] || {};

          // Calculate Category Averages
          if (course.gradebookSettings && course.gradebookSettings.gradeCategories) {
            course.gradebookSettings.gradeCategories.forEach(cat => {
              let earned = 0;
              let possible = 0;
              
              validWork.forEach(work => {
                if (work.gradeCategoryId === cat.id && studentGrades[work.id] !== undefined) {
                  earned += studentGrades[work.id];
                  possible += (work.maxPoints || 100);
                }
              });
              
              if (possible > 0) {
                // CHANGED: Multiply by 10 instead of 100, and remove the '%' sign
                let average = ((earned / possible) * 10).toFixed(1);
                studentRow.push(average);
              } else {
                studentRow.push("N/A"); // No graded assignments in this category yet
              }
            });
          }

          // Add Individual Assignment Grades
          validWork.forEach(work => {
            if (studentGrades[work.id] !== undefined) {
              studentRow.push(studentGrades[work.id]);
            } else {
              studentRow.push(""); // Blank if not graded yet
            }
          });

          rowData.push(studentRow);
        });

        // Write all student data to the sheet at once (much faster!)
        sheet.getRange(2, 1, rowData.length, headerRow.length).setValues(rowData);
      }

      // Format Sheet nicely
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(1);
      sheet.autoResizeColumns(1, headerRow.length);
    }
    
  } catch (err) {
    Logger.log('Failed with error: ' + err.message);
  }
}
