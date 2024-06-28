var BATCH_SIZE = 8;
var SLEEP_DURATION = 1000; // 1 second

function getEmailSearchCriteriaFromSheet() {
  var spreadsheetId = '####EnterSheetId here###'
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('filter');
  var data = sheet.getDataRange().getValues();
  
  var emailCriteria = [];
  
  for (var i = 1; i < data.length; i++) { // Skip header row
    var row = data[i];
    var criteria = {
      sender: row[0],
      months: row[1],
      subject: row[2] ? row[2] : null,
      excludedText: row[3] ? row[3] : null
    };
    emailCriteria.push(criteria);
  }
  
  return emailCriteria;
}

function cleanEmailsBySenderAndMonths() {
  var today = new Date();
  var emailCriteria = getEmailSearchCriteriaFromSheet();

  // Iterate through each email criteria
  emailCriteria.forEach(function(criteria) {
    var monthsToDelete = criteria.months;
    var senderEmail = criteria.sender;
    var subjectFilter = criteria.subject;
    var excludedText = criteria.excludedText;

    // Calculate the date threshold (monthsToDelete months ago from today)
    var olderThanDate = new Date();
    olderThanDate.setMonth(today.getMonth() - monthsToDelete);
    var olderThanDateString = Utilities.formatDate(olderThanDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

    // Construct search query
    var searchQuery = "from:" + senderEmail + " before:" + olderThanDateString;
    if (subjectFilter) {
      searchQuery += " subject:" + subjectFilter;
    }
    if (excludedText) {
      searchQuery += " -" + excludedText;
    }

    // Search for threads matching the criteria
    var threads = GmailApp.search(searchQuery);
      Logger.log("Deleting :" + threads.length +" messages from : " + senderEmail);

    // Process each thread in batches
    var deleteCount = 0;
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();

      messages.forEach(function(message) {
        try {
          message.moveToTrash();
          deleteCount++;

          if (deleteCount % BATCH_SIZE === 0) {
            Logger.log("Deleted :" + deleteCount);
            Utilities.sleep(SLEEP_DURATION); 
          }
        } catch (error) {
          console.error('Error: ' + error.message + '\nStack trace:\n' + error.stack);
        }
      });
    }
    Logger.log("Deleted :" + threads.length +" messages from : " + senderEmail);
  });

}
