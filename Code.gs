// Function to get data from the sheet with a hard-coded structure
function getSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) { // Check if there are no rows with data
    return [];
  }
  
  // Retrieve data starting from the second row to avoid headers
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 4);
  const data = dataRange.getValues();
  
  // Map the data into an array of objects
  return data.map(row => ({
    firstName: row[0],  // Column A
    lastName: row[1],   // Column B
    email: row[2],      // Column C
    otherInfo: row[3]   // Column D
  }));
}

// Function to validate email addresses
function isValidEmail(email) {
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}

// Function to schedule emails based on the sheet data and user input
function scheduleEmails(scheduledDate, scheduledTime, timezone, emailInterval) {
  const sheetData = getSheetData();  // Get data from the sheet

  // Calculate the initial time to send the first email
  const startTime = new Date(`${scheduledDate} ${scheduledTime}`);

  // Store the time and interval in user properties for the trigger to use
  storeEmailSchedule(startTime, emailInterval);

  // Create a single trigger to run the 'sendEmailsInBatch' function at the specified time
  ScriptApp.newTrigger('sendEmailsInBatch')
    .timeBased()
    .at(startTime)
    .create();
}

// Store the schedule in user properties for use by the triggered function
function storeEmailSchedule(startTime, emailInterval) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('startTime', startTime.toISOString());
  userProperties.setProperty('emailInterval', emailInterval);
}

// Function to send emails in a batch with intervals
function sendEmailsInBatch() {
  const sheetData = getSheetData();  // Get data from the sheet
  const userProperties = PropertiesService.getUserProperties();

  const startTime = new Date(userProperties.getProperty('startTime'));
  const emailInterval = parseInt(userProperties.getProperty('emailInterval'), 10) * 1000;  // Convert to milliseconds

  const emailSubject = userProperties.getProperty('emailSubject');
  const emailContent = userProperties.getProperty('emailContent');

  if (!emailSubject || !emailContent) {
    Logger.log('Email subject or content is missing.');
    return;
  }

  sheetData.forEach((recipient, index) => {
    if (!isValidEmail(recipient.email)) {
      Logger.log(`Invalid email: ${recipient.email}`);
      return;  // Skip invalid email addresses
    }

    const personalizedContent = emailContent
      .replace(/{{FirstName}}/g, recipient.firstName)
      .replace(/{{LastName}}/g, recipient.lastName);

    try {
      // Calculate when this specific email should be sent
      const sendTime = new Date(startTime.getTime() + (index * emailInterval));

      if (new Date() >= sendTime) {
        // If the time has already passed, send the email immediately
        MailApp.sendEmail({
          to: recipient.email,
          subject: emailSubject,
          htmlBody: personalizedContent
        });
        Logger.log(`Email sent to: ${recipient.email}`);
      } else {
        // Delay the email sending until the calculated sendTime
        Utilities.sleep(sendTime - new Date());  // Sleep until it's time to send the email
        MailApp.sendEmail({
          to: recipient.email,
          subject: emailSubject,
          htmlBody: personalizedContent
        });
        Logger.log(`Email sent to: ${recipient.email}`);
      }
    } catch (error) {
      Logger.log(`Failed to send email to: ${recipient.email} with error: ${error.message}`);
    }
  });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mail Merge')
    .addItem('Open Mail Merge', 'openMailMerge')
    .addToUi();
}

function openMailMerge() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Mail Merge');
}

function getTemplateList() {
  var templates = [
    { id: 'template1', name: 'Welcome Template', subject: 'Welcome to Our Service', content: 'Hello {{FirstName}} {{LastName}},<br>Greetings from our team!' },
    { id: 'template2', name: 'Follow-Up Template', subject: 'Follow-Up: How Are You?', content: 'Dear {{FirstName}} {{LastName}},<br>Just checking in to see how you\'re doing.' },
    { id: 'template3', name: 'Event Invitation Template', subject: 'You\'re Invited!', content: 'Hi {{FirstName}},<br>You\'re invited to our exclusive event!' },
    { id: 'template4', name: 'Thank You Note', subject: 'Thank You!', content: 'Dear {{FirstName}},<br>Thank you for your continued support.' },
    { id: 'template5', name: 'Special Offer', subject: 'Special Offer Just for You', content: 'Hi {{FirstName}},<br>We have a special offer just for you!' },
    { id: 'template6', name: 'Survey Request', subject: 'We Value Your Feedback', content: 'Dear {{FirstName}},<br>We value your feedback. Please take our survey.' },
    { id: 'template7', name: 'Appointment Reminder', subject: 'Appointment Reminder', content: 'Hi {{FirstName}},<br>This is a reminder for your upcoming appointment.' },
    { id: 'template8', name: 'Newsletter', subject: 'Monthly Newsletter', content: 'Hello {{FirstName}},<br>Welcome to our monthly newsletter.' },
    { id: 'template9', name: 'Product Update', subject: 'Exciting Product Updates', content: 'Hi {{FirstName}},<br>We have some exciting product updates to share.' },
    { id: 'template10', name: 'Birthday Greeting', subject: 'Happy Birthday!', content: 'Dear {{FirstName}},<br>Happy Birthday! We wish you all the best.' },
    { id: 'template11', name: 'Holiday Wishes', subject: 'Warm Holiday Wishes', content: 'Hi {{FirstName}},<br>Warm wishes for a happy holiday season.' },
    { id: 'template12', name: 'Feedback Request', subject: 'We Would Love Your Feedback', content: 'Dear {{FirstName}},<br>We would love to hear your feedback.' },
    { id: 'template13', name: 'Service Announcement', subject: 'Important Service Announcement', content: 'Hi {{FirstName}},<br>Important service announcement.' },
    { id: 'template14', name: 'Welcome Aboard', subject: 'Welcome Aboard!', content: 'Hello {{FirstName}},<br>Welcome aboard! We\'re excited to have you.' },
    { id: 'template15', name: 'Goodbye Note', subject: 'Goodbye and Best Wishes', content: 'Dear {{FirstName}},<br>We\'re sad to see you go. Best wishes!' }
  ];
  return templates;
}

function getTemplateContent(templateId) {
  var templates = getTemplateList();
  var selectedTemplate = templates.find(template => template.id === templateId);
  return selectedTemplate ? selectedTemplate : { subject: '', content: '' };
}

function getSheetList() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(function(sheet) {
    return { name: sheet.getName() };
  });
}

// Function to store email content, checking if a template is selected
function storeEmailContent(emailSubject, emailContent, templateId) {
  if (templateId) {
    const template = getTemplateContent(templateId);
    if (template.subject && template.content) {
      emailSubject = template.subject;
      emailContent = template.content;
    }
  }
  
  Logger.log('Storing email content: subject=%s, content=%s', emailSubject, emailContent);
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('emailSubject', emailSubject);
  userProperties.setProperty('emailContent', emailContent);
}

// Function to store email content at runtime
function storeRuntimeEmailContent(emailSubject, emailContent, templateId) {
  storeEmailContent(emailSubject, emailContent, templateId);
}

// When the user selects a template, call storeRuntimeEmailContent with the templateId
function scheduleEmails(scheduledDate, scheduledTime, timezone, emailInterval, templateId) {
  const sheetData = getSheetData();  // Get data from the sheet

  // Store the selected template content or runtime content
  storeRuntimeEmailContent(null, null, templateId); // Null for emailSubject and emailContent if using template
  
  // Calculate the initial time to send the first email
  const startTime = new Date(`${scheduledDate} ${scheduledTime}`);

  // Store the time and interval in user properties for the trigger to use
  storeEmailSchedule(startTime, emailInterval);

  // Create a single trigger to run the 'sendEmailsInBatch' function at the specified time
  ScriptApp.newTrigger('sendEmailsInBatch')
    .timeBased()
    .at(startTime)
    .create();
}


function getRuntimeEmailContent() {
  var userProperties = PropertiesService.getUserProperties();
  var emailSubject = userProperties.getProperty('emailSubject');
  var emailContent = userProperties.getProperty('emailContent');
  Logger.log('Retrieved email content: subject=%s, content=%s', emailSubject, emailContent);
  return { emailSubject: emailSubject, emailContent: emailContent };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

