// Function to serve the HTML file 
function doGet(e) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Madzimbani Sec. School')  // Optional: Set the title of the web app
      .setSandboxMode(HtmlService.SandboxMode.IFRAME); // IFRAME is optional depending on your needs
  }
  // Function to submit marks data
  function submitMarks(data) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(data.class); // Get the appropriate sheet (e.g., Form 1)
    const subjectCol = getSubjectColumn(sheet, data.subject); // Get the subject column
    const existingMarks = [];
  
    data.marks.forEach(function(markEntry) {
      const admissionNumber = markEntry.admissionNumber;
      const mark = markEntry.mark;
  
      const row = findRowByAdmissionNumber(sheet, admissionNumber); // Find the row for the student
  
      if (row && sheet.getRange(row, subjectCol).getValue()) {
        // If student already has marks, add them to the existing list for potential overwrite
        existingMarks.push({ admissionNumber, mark });
      } else if (row) {
        // If no marks exist, set the new mark
        sheet.getRange(row, subjectCol).setValue(mark);
      }
    });
  
    return existingMarks; // Return the students who already have marks to handle overwriting
  }
  
  // Function to overwrite marks
  function overwriteMarks(data) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(data.class); // Get the appropriate sheet
    const subjectCol = getSubjectColumn(sheet, data.subject); // Get the subject column
  
    data.marks.forEach(function(markEntry) {
      const admissionNumber = markEntry.admissionNumber;
      const mark = markEntry.mark;
  
      const row = findRowByAdmissionNumber(sheet, admissionNumber); // Find the row for the student
  
      if (row) {
        sheet.getRange(row, subjectCol).setValue(mark); // Overwrite the mark
      }
    });
  
    return "Marks overwritten successfully";
  }
  
  // Function to find the row of a student by admission number
  function findRowByAdmissionNumber(sheet, admissionNumber) {
    const data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues(); // Get the admission numbers column
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == admissionNumber) {
        return i + 1; // Return the row number
      }
    }
    return null; // Return null if admission number is not found
  }
  
  // Function to get the subject column
  function getSubjectColumn(sheet, subject) {
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get the header row
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] == subject) {
        return i + 1; // Return the column number
      }
    }
    return null; // Return null if subject is not found
  }
  
  // Function to generate and print report for a student
  function generateReport(form, admissionNumber) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(form); // Get the sheet for the selected form
    const row = findRowByAdmissionNumber(sheet, admissionNumber); // Find the row for the student
  
    if (!row) {
      throw new Error("Student with the given admission number not found");
    }
  
    const reportData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get the student's report data
  
    const pdf = generatePDFReport(reportData); // Generate the PDF report
    const base64PDF = Utilities.base64Encode(pdf.getBytes()); // Convert the PDF to base64
  
    return base64PDF;
  }
  
  // Function to generate a PDF report (simplified version)
  function generatePDFReport(reportData) {
    const doc = DocumentApp.create('Student Report');
    const body = doc.getBody();
  
    body.appendParagraph('Student Report');
    body.appendParagraph('Admission Number: ' + reportData[0]);
  
    for (let i = 1; i < reportData.length; i++) {
      body.appendParagraph('Subject ' + i + ': ' + reportData[i]);
    }
  
    doc.saveAndClose();
  
    const pdfBlob = doc.getAs('application/pdf');
    return pdfBlob;
  }
  
  // Function to create a new exam
  function createExam(examData) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(examData.examName);
    
    // Assuming first row as headers with subject names
    sheet.getRange(1, 1).setValue('Admission Number'); // First column for admission numbers
    examData.subjects.forEach((subject, index) => {
      sheet.getRange(1, index + 2).setValue(subject); // Add subjects as headers in columns
    });
  
    return "Exam created successfully";
  }
  
  // Function to manage existing exams (e.g., delete or rename)
  function manageExam(action, examName, newExamName) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (action === "delete") {
      spreadsheet.deleteSheet(spreadsheet.getSheetByName(examName));
      return "Exam deleted successfully";
    } else if (action === "rename" && newExamName) {
      const sheet = spreadsheet.getSheetByName(examName);
      sheet.setName(newExamName);
      return "Exam renamed successfully";
    }
    return "Invalid action or parameters";
  }
  
  // Function to fetch the available classes for dropdown options
  function getAvailableClasses() {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const classNames = sheets.map(sheet => sheet.getName());
    return classNames;
  }
  
  // Function to fetch the subjects for a particular class
  function getAvailableSubjects(className) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(className);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Assuming subjects start from the second column
    return headerRow.slice(1);
  }
  
  // Handle form submission errors
  function onSubmitFailure(e) {
    Logger.log('Error: ' + e.toString());
    return "Error occurred: " + e.toString();
  }
  
  // Utility function for handling success
  function onSubmitSuccess(result) {
    Logger.log('Success: ' + result);
    return result;
  }
  