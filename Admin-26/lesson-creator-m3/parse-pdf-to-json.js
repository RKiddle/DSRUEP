function extractTextFromPDF(pdfFileId) {
  // 1. Fetch the PDF file from your Drive using its ID
  const pdfFile = DriveApp.getFileById(pdfFileId);
  const pdfBlob = pdfFile.getBlob();
  
  // 2. Set up the configuration for the temporary Google Doc
  const resource = {
    name: "Temp_OCR_" + pdfFile.getName(),
    mimeType: MimeType.GOOGLE_DOCS // This specific mimeType triggers the OCR conversion
  };
  
  // 3. Create the temporary Doc (Google converts the PDF blob into text here)
  const tempDocFile = Drive.Files.create(resource, pdfBlob);
  
  // 4. Open the newly created Doc and extract the text
  const tempDoc = DocumentApp.openById(tempDocFile.id);
  const chapterText = tempDoc.getBody().getText();
  
  // 5. Clean up: Send the temporary Doc to the trash
  DriveApp.getFileById(tempDocFile.id).setTrashed(true);
  
  // 6. Return the raw text string to be sent to Gemini
  return chapterText;
}

// Example of how to run it:
function testExtraction() {
  // Replace with the actual ID of one of your M3 chapter PDFs
  const myFileId = "1A2B3C4D5E6F7G8H9I0J"; 
  const text = extractTextFromPDF(myFileId);
  Logger.log(text);
}
