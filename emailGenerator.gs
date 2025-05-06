// Gets called automatically 
function onInstall() {
  onOpen();
}

// Adds a custom menu to Google Docs
function onOpen() {
  DocumentApp.getUi()
  .createAddonMenu()
  .addItem("Create Email List", "showSidebar")
  .addToUi();
}

// Displays the sidebar UI
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('createCalendar')
      .setTitle('Generate Emails')
      .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

// Defines a class for each student ambassador
// Input: name (string), email (string without "@usc.edu")
// Creates an object with .name and .email properties
class StudentAmbassador {
  constructor(name, email) {
    this.name = name;
    this.email = email;
  }
}

// Main function that processes the Google Doc and returns a list of matched emails
// Input: None (main)
// Return: Array of strings (e.g., ["prostaff@maillist.usc.edu", "jsmith@usc.edu", ...])
function main() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var text = body.getText();
  var totalObjectList = [];

  // Data has been anonymized to ensure privacy
  let anon_names = ["John Smith", "Emily Johnson", "Michael Davis", "Sarah Brown", "David Wilson"];
  let anon_emails = ["jsmith@usc.edu", "ejohnson@usc.edu", "mdavis@usc.edu", "sbrown@usc.edu", "dwilson@usc.edu"];
  totalObjectList = totalObjectList.concat(group(anon_names, anon_emails));
  let totalEmailList = ["prostaff@maillist.usc.edu"]
  totalEmailList = totalEmailList.concat(extractEmailList(text, totalObjectList));

  return totalEmailList;
}

// Combines name and emails into StudentAmbassador objects
// Input: allNamesList (array of strings), allEmailsList (array of strings)
// Return: array of StudentAmbassador objects
function group(allNamesList, allEmailsList) {
  let group = [];
  for (let i = 0; i < allNamesList.length; i++) {
    group.push(new StudentAmbassador(allNamesList[i], allEmailsList[i]));
  }
  return group;
}

// Matches names in the document to known StudentAmbassador names
// Input: text (string), objectList (array of StudentAmbassador)
// Return: array of email strings
function extractEmailList(text, objectList) {
  const lowerText = text.toLowerCase();
  let emails = [];
  for (let i = 0; i < objectList.length; i++) {
    if (lowerText.includes(objectList[i].name.toLowerCase())) {
      emails.push(objectList[i].email);
    }
  }
  return emails;
}
