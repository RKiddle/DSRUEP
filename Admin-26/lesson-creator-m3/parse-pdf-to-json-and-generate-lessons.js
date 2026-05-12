/**
 * Reads a textbook chapter PDF from Google Drive, extracts OCR text,
 * cleans it, chunks it, packages it as JSON, and sends it to Gemini
 * for science lesson generation.
 *
 * Requirements:
 * - Apps Script Advanced Drive service enabled (`Drive`)
 * - Access to the source PDF in Google Drive
 */

/**
 * Main orchestration function.
 * Builds the lesson input JSON and sends it to Gemini.
 *
 * @param {string} pdfFileId Google Drive file ID for a PDF.
 * @param {string} apiKey Gemini API key.
 * @return {Object} Parsed JSON response from Gemini when possible, otherwise raw response text.
 */
function generateScienceLessonsFromPdf(pdfFileId, apiKey) {
  if (!apiKey || typeof apiKey !== 'string') {
    throw new Error('A valid Gemini API key is required.');
  }

  const lessonInput = convertChapterPdfToLessonInputJson(pdfFileId);
  return generateScienceLessonsFromLessonInput(lessonInput, apiKey);
}

/**
 * Sends structured lesson input JSON to Gemini.
 *
 * @param {Object} lessonInput Structured chapter JSON.
 * @param {string} apiKey Gemini API key.
 * @return {Object|string}
 */
function generateScienceLessonsFromLessonInput(lessonInput, apiKey) {
  if (!lessonInput || typeof lessonInput !== 'object') {
    throw new Error('A valid lesson input object is required.');
  }

  if (!apiKey || typeof apiKey !== 'string') {
    throw new Error('A valid Gemini API key is required.');
  }

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
          'Generate multiple-choice quizzes formatted strictly for Google Apps Script array ingestion. ' +
          'Do not include markdown fences. Return JSON only.'
      }]
    },
    contents: [{
      parts: [{
        text:
          'Process the following chapter JSON. ' +
          'Prioritize chapter.title, chapter.cleanText, chapter.chunks, and generationHints. ' +
          'Return only valid JSON.\n\n' + lessonInputJson
      }]
    }],
    generationConfig: {
      temperature: 0.3,
      responseMimeType: 'application/json'
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

/**
 * Converts a chapter PDF into structured JSON intended for lesson generation.
 *
 * @param {string} pdfFileId
 * @return {Object}
 */
function convertChapterPdfToLessonInputJson(pdfFileId) {
  const extracted = extractTextFromPDF(pdfFileId);
  const cleanText = normalizeExtractedText(extracted.text);
  const chapterTitle = extractChapterTitle(cleanText, extracted.fileName);
  const chunks = splitTextIntoChunks(cleanText, 4000);

  return {
    task: 'Create science lessons from a textbook chapter',
    audience: {
      programs: ['M3A', 'M3B'],
      subject: 'Science',
      level: 'secondary'
    },
    source: {
      fileId: extracted.fileId,
      fileName: extracted.fileName,
      mimeType: extracted.mimeType
    },
    chapter: {
      title: chapterTitle,
      cleanText: cleanText,
      chunks: chunks
    },
    generationHints: {
      chunkingStrategy: 'paragraph-aware',
      notes: [
        'M3A requires full sentence construction in English.',
        'M3B requires fill-in-the-blank English sentences with Thai translations for scaffolding.',
        'Use the chapter content as the source of truth.',
        'Keep science explanations accurate and grade-appropriate.'
      ]
    },
    stats: {
      rawCharacterCount: extracted.text.length,
      cleanCharacterCount: cleanText.length,
      chunkCount: chunks.length
    }
  };
}

/**
 * Extracts text from a PDF by converting it to a temporary Google Doc.
 *
 * @param {string} pdfFileId
 * @return {{fileId: string, fileName: string, mimeType: string, text: string}}
 */
function extractTextFromPDF(pdfFileId) {
  if (!pdfFileId || typeof pdfFileId !== 'string') {
    throw new Error('A valid PDF file ID string is required.');
  }

  let tempDocId = null;

  try {
    const pdfFile = DriveApp.getFileById(pdfFileId);
    const fileName = pdfFile.getName();
    const mimeType = pdfFile.getMimeType();

    if (mimeType !== MimeType.PDF && mimeType !== 'application/pdf') {
      throw new Error('The provided file is not a PDF. Mime type: ' + mimeType);
    }

    const pdfBlob = pdfFile.getBlob();
    const resource = {
      name: 'Temp_OCR_' + fileName,
      mimeType: MimeType.GOOGLE_DOCS
    };

    const tempDocFile = Drive.Files.create(resource, pdfBlob);
    tempDocId = tempDocFile.id;

    const tempDoc = DocumentApp.openById(tempDocId);
    const text = tempDoc.getBody().getText();

    return {
      fileId: pdfFileId,
      fileName: fileName,
      mimeType: mimeType,
      text: text || ''
    };
  } catch (error) {
    throw new Error('Failed to extract text from PDF "' + pdfFileId + '": ' + error.message);
  } finally {
    if (tempDocId) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
      } catch (cleanupError) {
        Logger.log('Cleanup failed for temp doc ' + tempDocId + ': ' + cleanupError.message);
      }
    }
  }
}

/**
 * Normalizes OCR output for cleaner downstream processing.
 *
 * @param {string} text
 * @return {string}
 */
function normalizeExtractedText(text) {
  if (!text) return '';

  return text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/-\n(?=\w)/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Extracts a likely chapter title from the first lines of text.
 *
 * @param {string} text
 * @param {string} fallbackFileName
 * @return {string}
 */
function extractChapterTitle(text, fallbackFileName) {
  const lines = text.split('\n')
    .map(function(line) { return line.trim(); })
    .filter(function(line) { return line; });

  for (let i = 0; i < Math.min(lines.length, 12); i++) {
    if (/^(chapter|unit|lesson)\b/i.test(lines[i])) {
      return lines[i];
    }
  }

  if (lines.length > 0) {
    return lines[0];
  }

  return fallbackFileName
    .replace(/\.pdf$/i, '')
    .replace(/[_-]+/g, ' ')
    .trim();
}

/**
 * Splits text into LLM-friendly chunks while trying to preserve paragraph boundaries.
 *
 * @param {string} text
 * @param {number} maxChunkLength
 * @return {Array<Object>}
 */
function splitTextIntoChunks(text, maxChunkLength) {
  if (!text) return [];

  const paragraphs = text.split(/\n\s*\n/).map(function(paragraph) {
    return paragraph.trim();
  }).filter(function(paragraph) {
    return paragraph;
  });

  const chunks = [];
  let currentText = '';
  let currentHeading = '';

  paragraphs.forEach(function(paragraph) {
    const inferredHeading = inferHeadingFromParagraph(paragraph);
    if (inferredHeading) {
      currentHeading = inferredHeading;
    }

    const candidate = currentText
      ? currentText + '\n\n' + paragraph
      : paragraph;

    if (candidate.length > maxChunkLength && currentText) {
      chunks.push({
        index: chunks.length + 1,
        heading: currentHeading || 'Chunk ' + (chunks.length + 1),
        text: currentText
      });
      currentText = paragraph;
    } else {
      currentText = candidate;
    }
  });

  if (currentText) {
    chunks.push({
      index: chunks.length + 1,
      heading: currentHeading || 'Chunk ' + (chunks.length + 1),
      text: currentText
    });
  }

  return chunks;
}

/**
 * Attempts to identify a heading from a short standalone paragraph.
 *
 * @param {string} paragraph
 * @return {string}
 */
function inferHeadingFromParagraph(paragraph) {
  const singleLine = paragraph.replace(/\n/g, ' ').trim();

  if (
    singleLine.length > 0 &&
    singleLine.length <= 80 &&
    !/[.!?]$/.test(singleLine)
  ) {
    return singleLine;
  }

  return '';
}

/**
 * Saves the lesson-input JSON as a Drive file.
 *
 * @param {string} pdfFileId
 * @param {string=} outputFolderId
 * @return {string} Created file ID
 */
function saveLessonInputJsonToDrive(pdfFileId, outputFolderId) {
  const exportData = convertChapterPdfToLessonInputJson(pdfFileId);
  const safeTitle = exportData.chapter.title.replace(/[^\w\s-]/g, '').trim();
  const fileName = (safeTitle || 'chapter-export') + '.json';
  const jsonString = JSON.stringify(exportData, null, 2);
  const blob = Utilities.newBlob(jsonString, 'application/json', fileName);

  if (outputFolderId) {
    const folder = DriveApp.getFolderById(outputFolderId);
    return folder.createFile(blob).getId();
  }

  return DriveApp.createFile(blob).getId();
}

/**
 * Attempts to parse the Gemini API response into JSON.
 * Returns parsed JSON when possible, otherwise a wrapper object.
 *
 * @param {string} responseText
 * @return {Object}
 */
function parseGeminiJsonResponse(responseText) {
  let response;
  try {
    response = JSON.parse(responseText);
  } catch (error) {
    return {
      success: false,
      error: 'Gemini returned non-JSON HTTP body.',
      rawResponse: responseText
    };
  }

  const candidateText = extractGeminiTextCandidate(response);
  if (!candidateText) {
    return {
      success: false,
      error: 'No text content found in Gemini response.',
      rawResponse: response
    };
  }

  try {
    return JSON.parse(candidateText);
  } catch (error) {
    return {
      success: false,
      error: 'Gemini text was not valid JSON.',
      rawText: candidateText,
      rawResponse: response
    };
  }
}

/**
 * Pulls the primary text candidate from a Gemini response.
 *
 * @param {Object} response
 * @return {string}
 */
function extractGeminiTextCandidate(response) {
  if (!response || !response.candidates || !response.candidates.length) {
    return '';
  }

  const parts = (((response.candidates[0] || {}).content || {}).parts || []);
  const textParts = parts
    .map(function(part) { return part && part.text ? part.text : ''; })
    .filter(function(text) { return text; });

  return textParts.join('\n').trim();
}

/**
 * Debug helper.
 */
function testGenerateScienceLessonsFromPdf() {
  const pdfFileId = '1JrIz0zLIJMrxaN5eaUcZgjdvwtnKlXN6'; // Put your actual PDF Doc ID here
  
  // Securely fetch the API key from Script Properties
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  // Optional: Add a quick safety check so the script warns you if it can't find the key
  if (!apiKey) {
    Logger.log("ERROR: Could not find GEMINI_API_KEY in Script Properties. Please check Project Settings.");
    return;
  }

  const result = generateScienceLessonsFromPdf(pdfFileId, apiKey);
  Logger.log(JSON.stringify(result, null, 2));
}
