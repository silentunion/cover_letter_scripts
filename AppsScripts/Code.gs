function autoFillCoverLetter(e) {
  // Get variables from spreadsheet
  var jobPosition = e.values[1];
  var companyName = e.values[2];
  var companySuffix = e.values[3];
  var companyReason = e.values[4];
  var jobLocation = e.values[5];
  var jobCountry = e.values[7];
  var cLType = e.values[6];

  // Get cover letter template
  var cLTemplate = "";
  if (cLType == "DEV"){
    // Cover Letter Template - Software Engineer
    cLTemplate = DriveApp.getFileById("1eP98mOXS27pI0sJL-qLAX9E6ME7QSMw8fzZLE1frA0Y");
  }
  else if (cLType == "DBA"){
    // Cover Letter Template - Database
    cLTemplate = DriveApp.getFileById("1xxZwi0CmwmLyC1gKN_eqY9Perj9Jmr9sknwJ7EOUNLQ");
  }
  else {
    // Default set to Database, should never get to this point through required dropdown menu
    cLTemplate = DriveApp.getFileById("1xxZwi0CmwmLyC1gKN_eqY9Perj9Jmr9sknwJ7EOUNLQ");
  }

  //Get folder to copy cover letter template to and copy
  var cLResponseFolder = DriveApp.getFolderById("1ojsIeuYNUtb2dCgJLZ1_bzIDMV36flte");
  var companyNameConcat = companyName.replace(/ /g, "_");
  var copy = cLTemplate.makeCopy("Reece_Milliner-CL-" + companyNameConcat, cLResponseFolder);
  
  // Open the newly created cover letter and get the body
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();

  // Replace the parts with the form parts
  body.replaceText("{{JobPosition}}", jobPosition);
  body.replaceText("{{CompanyName}}", companyName);
  body.replaceText("{{CompanySuffix}}", companySuffix);
  body.replaceText("{{CompanyReason}}", companyReason);

  // Location of job determines which sentence is used, option to "Ignore" location should delete sentence entirely
  if (jobCountry == "Canada"){
    body.replaceText("{{JobLocation}}", "Furthermore, I would relish the opportunity to move to " + jobLocation + " and can move almost immediately.");
  }
  else if (jobCountry == "Rest of World"){
    body.replaceText("{{JobLocation}}", "Furthermore, I would relish the opportunity to move to " + jobLocation + " and have nothing holding me back locally from moving.");
  }
  else {
    body.replaceText("{{JobLocation}}", "");
  }

  // Add in today's date
  body.replaceText("{{Date}}", Utilities.formatDate(new Date(), "GMT-8", "MMMM dd, yyyy"));

  doc.saveAndClose();
}