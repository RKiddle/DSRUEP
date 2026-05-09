/**
 * Reads assignment rows from the active sheet and creates published Google
 * Classroom assignments for rows that are ready to post.
 *
 * Expected sheet columns:
 * 0: Post Date
 * 1: Title
 * 2: Description
 * 3: Course ID
 * 4: Due Date
 * 5: Points
 * 6: Topic Name
 * 7: Material URL
 * 8: Status
 *
 * Only rows with a status of PENDING, a valid course ID, and a post date on or
 * before today are processed. Successful rows are marked POSTED; failures are
 * marked with an ERROR message.
 */
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

/**
 * Determines whether an assignment row is eligible to be posted today.
 *
 * @param {*} postDateRaw Raw spreadsheet value for the post date.
 * @param {Date} today Normalized current date for comparison.
 * @returns {boolean} True when the post date is empty or on/before today.
 */
function isReadyToPost(postDateRaw, today) {
  if (!postDateRaw) {
    return true;
  }

  const postDate = toValidDate(postDateRaw, "post date");
  return normalizeDate(postDate).getTime() <= today.getTime();
}

/**
 * Converts the spreadsheet points value into a valid Classroom maxPoints value.
 *
 * Blank values, "ungraded", and 0 are all treated as ungraded assignments and
 * return 0. Negative or non-numeric values throw an error.
 *
 * @param {*} rawPoints Raw spreadsheet value for assignment points.
 * @returns {number} Parsed non-negative point value.
 * @throws {Error} If the value is negative or not numeric.
 */
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

/**
 * Returns the topic ID for a given course/topic name pair, creating the topic if
 * it does not already exist. Topic data is cached per course for the duration of
 * the script run to reduce repeated API calls.
 *
 * @param {string} courseId Google Classroom course ID.
 * @param {string} topicName Topic name to find or create.
 * @param {Object.<string, Object.<string, string>>} topicCache In-memory cache of
 * topic names to topic IDs keyed by course ID.
 * @returns {string} Existing or newly created topic ID.
 */
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

/**
 * Builds the Classroom materials payload from a URL.
 *
 * If the URL appears to reference a Google Drive file, the file is validated for
 * access and attached as a STUDENT_COPY material. Otherwise, the URL is attached
 * as a standard link.
 *
 * @param {string} materialUrl URL to attach to the assignment.
 * @returns {Array<Object>} Classroom API materials array.
 * @throws {Error} If a referenced Drive file cannot be accessed.
 */
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

/**
 * Converts a value into a valid JavaScript Date.
 *
 * @param {*} value Raw date-like value.
 * @param {string} fieldName Human-readable field name for error messages.
 * @returns {Date} Parsed valid date object.
 * @throws {Error} If the value cannot be parsed as a date.
 */
function toValidDate(value, fieldName) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    throw new Error("Invalid " + fieldName + ": " + value);
  }
  return date;
}

/**
 * Returns a copy of a date with the time set to midnight.
 *
 * @param {Date} date Input date.
 * @returns {Date} Normalized date for date-only comparisons.
 */
function normalizeDate(date) {
  const normalized = new Date(date);
  normalized.setHours(0, 0, 0, 0);
  return normalized;
}
