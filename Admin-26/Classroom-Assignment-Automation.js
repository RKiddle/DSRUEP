function testCreateClassroomAssignment() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const today = normalizeDate(new Date());
  const topicCache = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const postDateRaw = row[0];
    const title = String(row[1] || "").trim();
    const description = String(row[2] || "").trim();
    const courseId = String(row[3] || "").trim();
    const dueDateRaw = row[4];
    const rawPoints = row[5];
    const topicName = String(row[6] || "").trim();
    const materialUrl = String(row[7] || "").trim();
    const status = String(row[8] || "").trim().toUpperCase();

    if (
      status !== "PENDING" ||
      !isReadyToPost(postDateRaw, today) ||
      !courseId ||
      courseId === "ENTER_YOUR_COURSE_ID_HERE"
    ) {
      continue;
    }

    try {
      const assignmentDetails = {
        title: title,
        description: description,
        workType: "ASSIGNMENT",
        state: "PUBLISHED"
      };

      assignmentDetails.maxPoints = parsePoints(rawPoints);

      if (dueDateRaw) {
        const dueDate = toValidDate(dueDateRaw, "due date");
        assignmentDetails.dueDate = {
          year: dueDate.getFullYear(),
          month: dueDate.getMonth() + 1,
          day: dueDate.getDate()
        };
        assignmentDetails.dueTime = { hours: 23, minutes: 59 };
      }

      if (topicName) {
        assignmentDetails.topicId = getOrCreateTopicId(courseId, topicName, topicCache);
      }

      if (materialUrl) {
        assignmentDetails.materials = buildMaterials(materialUrl);
      }

      Classroom.Courses.CourseWork.create(assignmentDetails, courseId);
      sheet.getRange(i + 1, 9).setValue("POSTED");
    } catch (error) {
      sheet.getRange(i + 1, 9).setValue("ERROR: " + error.message);
    }
  }
}

function isReadyToPost(postDateRaw, today) {
  if (!postDateRaw) {
    return true;
  }

  const postDate = toValidDate(postDateRaw, "post date");
  return normalizeDate(postDate).getTime() <= today.getTime();
}

function parsePoints(rawPoints) {
  const pointsStr = String(rawPoints == null ? "" : rawPoints).trim().toLowerCase();

  if (!pointsStr || pointsStr === "ungraded" || pointsStr === "0") {
    return 0;
  }

  const points = Number(rawPoints);
  if (Number.isNaN(points) || points < 0) {
    throw new Error("Invalid points value: " + rawPoints);
  }

  return points;
}

function getOrCreateTopicId(courseId, topicName, topicCache) {
  if (!topicCache[courseId]) {
    const topicsResponse = Classroom.Courses.Topics.list(courseId);
    const topics = topicsResponse.topic || [];
    topicCache[courseId] = {};

    topics.forEach(topic => {
      topicCache[courseId][topic.name] = topic.topicId;
    });
  }

  if (topicCache[courseId][topicName]) {
    return topicCache[courseId][topicName];
  }

  const newTopic = Classroom.Courses.Topics.create({ name: topicName }, courseId);
  topicCache[courseId][topicName] = newTopic.topicId;
  return newTopic.topicId;
}

function buildMaterials(materialUrl) {
  const driveIdMatch =
    materialUrl.match(/\/d\/([a-zA-Z0-9-_]+)/) ||
    materialUrl.match(/[?&]id=([a-zA-Z0-9-_]+)/);

  if (driveIdMatch && driveIdMatch[1]) {
    const fileId = driveIdMatch[1];

    try {
      DriveApp.getFileById(fileId);
    } catch (e) {
      throw new Error("Cannot access Drive file. You must own it or have edit access to make student copies.");
    }

    return [{
      driveFile: {
        driveFile: { id: fileId },
        shareMode: "STUDENT_COPY"
      }
    }];
  }

  return [{
    link: { url: materialUrl }
  }];
}

function toValidDate(value, fieldName) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    throw new Error("Invalid " + fieldName + ": " + value);
  }
  return date;
}

function normalizeDate(date) {
  const normalized = new Date(date);
  normalized.setHours(0, 0, 0, 0);
  return normalized;
}
