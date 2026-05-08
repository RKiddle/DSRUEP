function testCreateClassroomAssignment() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  // Get today's date and normalize it to midnight for accurate comparison
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Loop through rows (skipping the header at row 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const postDateRaw = row[0];      // Col A (NEW)
    const title = row[1];            // Col B
    const description = row[2];      // Col C
    const courseId = String(row[3]).trim(); // Col D
    const dueDateString = row[4];    // Col E
    const rawPoints = row[5];        // Col F 
    const topicName = row[6];        // Col G 
    const materialUrl = String(row[7]).trim(); // Col H 
    const status = row[8];           // Col I 

    // Determine if it's time to post based on the Post Date column
    let isReadyToPost = false;
    if (postDateRaw) {
      const postDate = new Date(postDateRaw);
      postDate.setHours(0, 0, 0, 0);
      
      // If the post date is today or earlier, it's ready
      if (postDate.getTime() <= today.getTime()) {
        isReadyToPost = true;
      }
    } else {
      // If you leave the Post Date blank, it assumes you want it posted immediately
      isReadyToPost = true; 
    }

    // Check if the row is meant to be processed
    if (status === "PENDING" && isReadyToPost && courseId && courseId !== "ENTER_YOUR_COURSE_ID_HERE") {
      
      try {
        const assignmentDetails = {
          title: title,
          description: description,
          workType: "ASSIGNMENT",
          state: "PUBLISHED" 
        };

        // Handle Graded vs. Ungraded 
        const pointsStr = String(rawPoints).trim().toLowerCase();
        if (pointsStr === "ungraded" || pointsStr === "0") {
           assignmentDetails.maxPoints = 0; 
        } else {
           assignmentDetails.maxPoints = Number(rawPoints) || 100;
        }

        // Handle Due Date & Time
        if (dueDateString) {
          const dateObj = new Date(dueDateString);
          assignmentDetails.dueDate = {
            year: dateObj.getFullYear(),
            month: dateObj.getMonth() + 1,
            day: dateObj.getDate()
          };
          assignmentDetails.dueTime = { hours: 23, minutes: 59 };
        }

        // Handle Topic 
        if (topicName) {
          const topicsResponse = Classroom.Courses.Topics.list(courseId);
          const topics = topicsResponse.topic || [];
          const foundTopic = topics.find(t => t.name === topicName);

          if (foundTopic) {
            assignmentDetails.topicId = foundTopic.topicId;
          } else {
            const newTopic = Classroom.Courses.Topics.create({ name: topicName }, courseId);
            assignmentDetails.topicId = newTopic.topicId;
          }
        }

        // Handle Material URL & Drive Permissions
        if (materialUrl) {
          const driveIdMatch = materialUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
          
          if (driveIdMatch && driveIdMatch[1]) {
            const fileId = driveIdMatch[1];
            
            try {
              DriveApp.getFileById(fileId);
            } catch (e) {
              throw new Error("Cannot access Drive file. You must own it or have edit access to make student copies.");
            }

            assignmentDetails.materials = [{
              driveFile: {
                driveFile: { id: fileId },
                shareMode: "STUDENT_COPY" 
              }
            }];
          } else {
            assignmentDetails.materials = [{
              link: { url: materialUrl }
            }];
          }
        }

        // Push the fully compiled payload to Google Classroom
        Classroom.Courses.CourseWork.create(assignmentDetails, courseId);
        
        // Update the sheet 
        sheet.getRange(i + 1, 9).setValue("POSTED");
        
      } catch (error) {
        // Write the error directly to the sheet if anything fails
        sheet.getRange(i + 1, 9).setValue("ERROR: " + error.message);
      }
    }
  }
}
