/**
 * ============================================================================
 * CONFIGURATION & ENTRY POINT
 * ============================================================================
 */

/**
 * 1. THE CONTROL CENTER
 * Put your PDF ID here and click "Run" on this function to start the pipeline.
 */
function runCurriculumGenerator() {
  // Replace with the ID of the M3 Science chapter PDF in your Drive
  const pdfFileId = '1JrIz0zLIJMrxaN5eaUcZgjdvwtnKlXN6'; 
  
  Logger.log("Starting pipeline for PDF ID: " + pdfFileId);
  buildClassroomAssets(pdfFileId);
}


/**
 * ============================================================================
 * ORCHESTRATION & GOOGLE WORKSPACE GENERATORS
 * ============================================================================
 */

/**
 * Main orchestration function. Extracts text, calls Gemini, and builds the files.
 */
function buildClassroomAssets(pdfFileId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is missing from Script Properties.');
  }

  Logger.log("Extracting text and calling Gemini (this may take a minute)...");
  const result = generateScienceLessonsFromPdf(pdfFileId, apiKey);
  
  if (!result.success && result.error) {
    throw new Error("Pipeline failed during AI generation: " + result.error);
  }

  const chapterData = result.success === undefined ? result : result.data; 
  
  // --- NEW: The Safety Net & Debug Log ---
  // Print the first 1000 characters of the JSON so you can see what the AI actually named things
  Logger.log("AI JSON Output Preview: \n" + JSON.stringify(chapterData).substring(0, 1000));
  
  // Try to find the array, even if the AI used a different key name
  let lessonsArray = chapterData.lessons || chapterData.lessonList || chapterData.chapterLessons;
  
  // Sometimes it nests it inside a 'chapter' object
  if (!lessonsArray && chapterData.chapter && chapterData.chapter.lessons) {
    lessonsArray = chapterData.chapter.lessons;
  }

  if (!lessonsArray || !Array.isArray(lessonsArray)) {
    throw new Error("CRITICAL: Could not find the 'lessons' array. Check the JSON Output Preview in the log above to see what keys the AI used.");
  }
  // ----------------------------------------

  const safeTitle = chapterData.chapterTitle ? chapterData.chapterTitle.replace(/[^\w\s-]/g, '').trim() : "New Chapter";
  const chapterFolder = DriveApp.createFolder("M3 Science: " + safeTitle);
  Logger.log("Created master folder: " + chapterFolder.getUrl());
  
  // Loop through the safely found array
  lessonsArray.forEach((lesson, index) => {
    const lessonNum = index + 1;
    Logger.log(`Generating assets for Lesson ${lessonNum}: ${lesson.lessonTitle || 'Untitled Lesson'}`);
    
    createWorksheetDoc(lesson, 'M3A', chapterFolder.getId());
    createWorksheetDoc(lesson, 'M3B', chapterFolder.getId());
    createGoogleForm(lesson, chapterFolder.getId());
  });
  
  Logger.log("=== PIPELINE COMPLETE ===");
  Logger.log("Access your files here: " + chapterFolder.getUrl());
}

/**
 * Creates a formatted Google Doc Worksheet and moves it to the target folder.
 */
function createWorksheetDoc(lessonData, type, folderId) {
  const doc = DocumentApp.create(`Worksheet (${type}) - ${lessonData.lessonTitle}`);
  
  // Move to the designated chapter folder
  const file = DriveApp.getFileById(doc.getId());
  DriveApp.getFolderById(folderId).addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const body = doc.getBody();
  
  // Formatting the Header
  body.appendParagraph(lessonData.lessonTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(type + " Version").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("Name: __________________________ Class: M3/___ Date: ___________");
  
  body.appendHorizontalRule();
  body.appendParagraph("Learning Objectives:").setBold(true);
  lessonData.learningObjectives.forEach(obj => body.appendListItem(obj));
  
  body.appendHorizontalRule();
  
  // Determine which content array to use based on type
  const content = (type === 'M3A') ? lessonData.materials.M3A : lessonData.materials.M3B;
  
  // Add the lesson sentences/fill-in-the-blanks
  if (content && content.length > 0) {
    content.forEach((text, i) => {
      body.appendParagraph(`${i + 1}. ${text}`);
      body.appendParagraph(""); // Adds spacing so you can manually insert diagrams later
    });
  } else {
    body.appendParagraph("[Content generation error for this section]");
  }
}

/**
 * Creates a self-grading Google Form Quiz and moves it to the target folder.
 */
function createGoogleForm(lessonData, folderId) {
  const form = FormApp.create('Quiz: ' + lessonData.lessonTitle);
  form.setIsQuiz(true);
  
  // Move to the designated chapter folder
  const file = DriveApp.getFileById(form.getId());
  DriveApp.getFolderById(folderId).addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // Add Multiple Choice Questions
  if (lessonData.quiz && lessonData.quiz.length > 0) {
    lessonData.quiz.forEach(item => {
      const mcItem = form.addMultipleChoiceItem();
      mcItem.setTitle(item.question)
            .setChoices(item.options.map(opt => {
              // Mark the choice as 'true' if it matches the answer string
              return mcItem.createChoice(opt, opt === item.answer);
            }))
            .setPoints(1)
            .setRequired(true);
    });
  }
}


/**
 * ============================================================================
 * AI & API LOGIC (GEMINI 2.5 PRO)
 * ============================================================================
 */

function generateScienceLessonsFromPdf(pdfFileId, apiKey) {
  const lessonInput = convertChapterPdfToLessonInputJson(pdfFileId);
  return generateScienceLessonsFromLessonInput(lessonInput, apiKey);
}

function generateScienceLessonsFromLessonInput(lessonInput, apiKey) {
  // Using the stable Gemini 2.5 Pro model
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + encodeURIComponent(apiKey);
  const lessonInputJson = JSON.stringify(lessonInput);

  const payload = {
    systemInstruction: {
      parts: [{
        text: 
          'You are an expert Science Curriculum Designer specializing in secondary education. ' +
          'Always output valid JSON. ' +
          'Use the provided chapter JSON as the source of truth. ' +
          'Break the chapter into logical lessons. ' +
          'Generate rigorous, scientifically accurate learning objectives. ' +
          'Create differentiated materials: M3A requires full sentence construction in English. ' +
          'M3B requires fill-in-the-blank English sentences paired with accurate Thai translations for scaffolding. ' +
          'Generate multiple-choice quizzes (with question, options array, and string answer). ' +
          'Do not include markdown fences. Return JSON only.'
      }]
    },
    contents: [{
      parts: [{
        text: 'Process the following chapter JSON. Prioritize generating the M3A, M3B, and quiz arrays. Return only valid JSON.\n\n' + lessonInputJson
      }]
    }],
    generationConfig: {
      temperature: 0.2, 
      responseMimeType: 'application/json',
      // NEW: Force the exact structure so the AI cannot hallucinate key names
      responseSchema: {
        type: "OBJECT",
        properties: {
          chapterTitle: { type: "STRING" },
          lessons: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                lessonTitle: { type: "STRING" },
                learningObjectives: { type: "ARRAY", items: { type: "STRING" } },
                materials: {
                  type: "OBJECT",
                  properties: {
                    M3A: { type: "ARRAY", items: { type: "STRING" } },
                    M3B: { type: "ARRAY", items: { type: "STRING" } }
                  }
                },
                quiz: {
                  type: "ARRAY",
                  items: {
                    type: "OBJECT",
                    properties: {
                      question: { type: "STRING" },
                      options: { type: "ARRAY", items: { type: "STRING" } },
                      answer: { type: "STRING" }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  const body = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error('Gemini request failed with status ' + status + ': ' + body);
  }

  return parseGeminiJsonResponse(body);
}

function parseGeminiJsonResponse(responseText) {
  let response;
  try {
    response = JSON.parse(responseText);
  } catch (error) {
    return { success: false, error: 'Gemini returned non-JSON HTTP body.' };
  }

  const candidateText = extractGeminiTextCandidate(response);
  if (!candidateText) {
    return { success: false, error: 'No text content found in Gemini response.' };
  }

  try {
    return JSON.parse(candidateText);
  } catch (error) {
    return { success: false, error: 'Gemini text was not valid JSON.' };
  }
}

function extractGeminiTextCandidate(response) {
  if (!response || !response.candidates || !response.candidates.length) return '';
  const parts = (((response.candidates[0] || {}).content || {}).parts || []);
  const textParts = parts.map(part => part && part.text ? part.text : '').filter(text => text);
  return textParts.join('\n').trim();
}


/**
 * ============================================================================
 * TEXT EXTRACTION & PROCESSING UTILITIES
 * ============================================================================
 */

function convertChapterPdfToLessonInputJson(pdfFileId) {
  const extracted = extractTextFromPDF(pdfFileId);
  const cleanText = normalizeExtractedText(extracted.text);
  const chapterTitle = extractChapterTitle(cleanText, extracted.fileName);
  
  // We chunk it to help the LLM process logically, even with a massive context window
  const chunks = splitTextIntoChunks(cleanText, 4000);

  return {
    chapter: {
      title: chapterTitle,
      cleanText: cleanText,
      chunks: chunks
    }
  };
}

function extractTextFromPDF(pdfFileId) {
  let tempDocId = null;
  try {
    const pdfFile = DriveApp.getFileById(pdfFileId);
    const fileName = pdfFile.getName();
    const pdfBlob = pdfFile.getBlob();
    
    // Convert via OCR
    const resource = {
      name: 'Temp_OCR_' + fileName,
      mimeType: MimeType.GOOGLE_DOCS
    };

    const tempDocFile = Drive.Files.create(resource, pdfBlob);
    tempDocId = tempDocFile.id;

    const tempDoc = DocumentApp.openById(tempDocId);
    const text = tempDoc.getBody().getText();

    return { fileName: fileName, text: text || '' };
    
  } catch (error) {
    throw new Error('Failed to extract text from PDF: ' + error.message);
  } finally {
    // ALWAYS clean up the temporary doc
    if (tempDocId) {
      DriveApp.getFileById(tempDocId).setTrashed(true);
    }
  }
}

function normalizeExtractedText(text) {
  if (!text) return '';
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/-\n(?=\w)/g, '') // Fix hyphenated words broken across lines
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function extractChapterTitle(text, fallbackFileName) {
  const lines = text.split('\n').map(line => line.trim()).filter(line => line);
  for (let i = 0; i < Math.min(lines.length, 12); i++) {
    if (/^(chapter|unit|lesson)\b/i.test(lines[i])) {
      return lines[i];
    }
  }
  return lines.length > 0 ? lines[0] : fallbackFileName.replace(/\.pdf$/i, '');
}

function splitTextIntoChunks(text, maxChunkLength) {
  if (!text) return [];
  const paragraphs = text.split(/\n\s*\n/).map(p => p.trim()).filter(p => p);
  const chunks = [];
  let currentText = '';

  paragraphs.forEach(paragraph => {
    const candidate = currentText ? currentText + '\n\n' + paragraph : paragraph;
    if (candidate.length > maxChunkLength && currentText) {
      chunks.push({ index: chunks.length + 1, text: currentText });
      currentText = paragraph;
    } else {
      currentText = candidate;
    }
  });

  if (currentText) {
    chunks.push({ index: chunks.length + 1, text: currentText });
  }
  return chunks;
}
