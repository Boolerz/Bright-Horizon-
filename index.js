// Function to serve the HTML file 
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Madzimbani Sec. School')  // Optional: Set the title of the web app
    .setSandboxMode(HtmlService.SandboxMode.IFRAME); // IFRAME is optional depending on your needs
}

// Function to get the Google Spreadsheet by class
function getSpreadsheetByClass(className) {
  const folder = DriveApp.getFoldersByName('TRIAL PROJECT MADZIMBANI').next(); // Access the folder
  let fileName = '';

  // Determine the file name based on the class
  if (className === 'Form 1') {
    fileName = 'Form 1';
  } else if (className === 'Form 2') {
    fileName = 'Form 2';
  } else if (className === 'Form 3') {
    fileName = 'Form 3';
  } else if (className === 'Form 4') {
    fileName = 'Form 4';
  }

  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.open(file); // Return the spreadsheet object
  } else {
    throw new Error('Spreadsheet not found for class: ' + className);
  }
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

// Define spreadsheet IDs for each form
const spreadsheetIds = {
    'Form 1': '1brwerE5ir5R5I2v93pfgtOiqWEkc62Hs_I1s_jWdahw', // Replace with actual ID for Form 1
    'Form 2': '1AsHP-PDVuNqXEyZlynAe11M-IJedr9z6Marq7kDiO7M', // Replace with actual ID for Form 2
    'Form 3': 'SPREADSHEET_ID_FORM_3', // Replace with actual ID for Form 3
    'Form 4': 'SPREADSHEET_ID_FORM_4'  // Replace with actual ID for Form 4
};
// Function to get the subject column (starting from column D for Maths)
function getSubjectColumn(subjectName) {
  const subjectColumns = {
    'Maths': 4,    // Column D
    'Eng': 5,      // Column E
    'Kisw': 6,     // Column F
    'Chem': 7,     // Column G
    'Phy': 8,      // Column H
    'Bio': 9,      // Column I
    'Hist': 10,    // Column J
    'Geo': 11,     // Column K
    'C.R.E': 12,   // Column L
    'I.R.E': 13,   // Column M
    'Agr': 14,     // Column N
    'B.std': 15    // Column O
  };
  return subjectColumns[subjectName] || null; // Return null if subject is not found
}
// Function to check if marks already exist for students in the class and subject
function checkMarksExist(data) {
    const sheet = getSheetByClassAndSubject(data.class, data.subject);
    const studentMarks = data.studentMarks;

    for (let i = 0; i < studentMarks.length; i++) {
        const studentRow = findStudentRow(sheet, studentMarks[i].admissionNumber);
        if (studentRow) {
            return true; // Marks already exist for at least one student
        }
    }
    return false; // No existing marks
}
// Function to overwrite marks and send email notification if needed
function overwriteMarks(data) {
    const sheet = getSheetByClassAndSubject(data.class, data.subject);
    const studentMarks = data.studentMarks;
    const subjectColumn = getSubjectColumn(data.subject);

    if (!subjectColumn) {
        throw new Error('Subject column not found. Please check the subject mapping.');
    }

    // Array to track students whose marks will be overwritten
    const studentsToNotify = [];
    let directSubmission = true; // Assume direct submission until an overwrite is detected

    studentMarks.forEach(function(student) {
        const studentRow = findStudentRow(sheet, student.admissionNumber);
        if (studentRow) {
            const existingMarks = sheet.getRange(studentRow, subjectColumn).getValue(); // Get existing marks from the subject column
            const newMarks = student.marks;

            // If existing marks are found, prepare for overwrite notification
            if (existingMarks) {
                studentsToNotify.push({
                    admissionNumber: student.admissionNumber,
                    existingMarks: existingMarks,
                    newMarks: newMarks
                });
                directSubmission = false; // Set to false if overwriting occurs
            }

            // Overwrite with new marks
            sheet.getRange(studentRow, subjectColumn).setValue(newMarks); 
        } else {
            // If student does not exist, add a new row with the marks in the correct column
            const newRow = [student.admissionNumber];
            newRow[subjectColumn - 1] = student.marks; // Place the marks in the correct subject column
            sheet.appendRow(newRow);
        }
    });

    // If there are students to notify, proceed with notifications for overwriting
    if (studentsToNotify.length > 0) {
        studentsToNotify.forEach(function(student) {
            const subject = 'Marks Changed Notification';
            const message = `Marks for a student in ${data.class} with admission number ${student.admissionNumber} have been changed from ${student.existingMarks} to ${student.newMarks} in ${data.subject}.`;
            MailApp.sendEmail('abuufondoh@gmail.com', subject, message);
        });
        return 'Marks have been successfully overwritten, and notifications have been sent.';
    } else if (directSubmission) {
        // If direct submission happened without overwriting existing marks
        return 'Marks submitted successfully without overwriting existing marks.';
    }

    // Fallback return if something unexpected occurs
    return 'Marks have been processed.';
}

// Helper function to find a student row by admission number
function findStudentRow(sheet, admissionNumber) {
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
        if (data[i][0] == admissionNumber) { // Assuming admission number is in the first column
            return i + 1; // Return the row number (1-based index)
        }
    }
    return null; // Student not found
}

// Helper function to get the sheet for the correct class and subject
function getSheetByClassAndSubject(className, subjectName) {
    // Get the spreadsheet ID based on the class (form)
    const spreadsheetId = spreadsheetIds[className];
    if (!spreadsheetId) {
        throw new Error(`Spreadsheet ID for ${className} not found.`);
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('Broad sheet'); // Assuming all marks go to the "Broad sheet"
    
    if (!sheet) {
        throw new Error(`Sheet for ${className} - ${subjectName} not found.`);
    }

    return sheet;
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

function downloadReport() {
    const className = document.getElementById('class').value;
    const term = document.getElementById('term').value;
    const admissionNumber = document.getElementById('admissionNumber').value;

    if (!className || !term || !admissionNumber) {
        alert('Please fill out all the fields.');
        return;
    }

    // This is where the backend report generation logic would go
    alert(`Generating report for:
    Class: ${className}
    Term: ${term}
    Admission Number: ${admissionNumber}`);

    // You can add backend logic here for report generation and download later.
}
