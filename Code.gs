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
    // Add more templates as needed
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

function storeRuntimeEmailContent(emailSubject, emailContent) {
  Logger.log('Storing email content: subject=%s, content=%s', emailSubject, emailContent);
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('emailSubject', emailSubject);
  userProperties.setProperty('emailContent', emailContent);
}

function getRuntimeEmailContent() {
  var userProperties = PropertiesService.getUserProperties();
  var emailSubject = userProperties.getProperty('emailSubject');
  var emailContent = userProperties.getProperty('emailContent');
  Logger.log('Retrieved email content: subject=%s, content=%s', emailSubject, emailContent);
  return { emailSubject: emailSubject, emailContent: emailContent };
}


function getTemplateContent(templateId) {
  var templates = getTemplateList();
  var selectedTemplate = templates.find(template => template.id === templateId);
  return selectedTemplate ? selectedTemplate : { content: '' };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
