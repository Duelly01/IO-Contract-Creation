// 🔒 SECURITY: Add the exact email addresses allowed to use this tool here
const ALLOWED_USERS = [
  "charuzukevinduelas@gmail.com",
  "charles.duelas@shopee.com",
  "moira.xu@shopee.com",
  "job.roque@shopee.com",
  "jeanette.mancita@shopee.com"
];

/**
 * Helper function to verify if the current user is authorized.
 */
function checkPermission() {
  const userEmail = Session.getActiveUser().getEmail();
  
  if (!userEmail) {
    SpreadsheetApp.getUi().alert("⛔ Security Error: Unable to verify your email identity. Please check your Google account permissions.");
    return false;
  }
  
  if (ALLOWED_USERS.indexOf(userEmail) === -1) {
    SpreadsheetApp.getUi().alert(`⛔ Access Denied: Your email (${userEmail}) is not authorized to run this automation.`);
    return false;
  }
  return true;
}

/**
 * Helper function to clean illegal characters
 */
function cleanFilename(fname, replaceWith) {
  const invchars = /[!"¤%&/()=?`^*>;:@£${[\]}|~\\,.'¨´+-]/g;
  return fname.replace(invchars, replaceWith);
}
