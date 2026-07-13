function createFolderOnEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // CONFIGURATION
  var targetColumn = 19; // Column S
  var startRow = 314;     // Starts looking from row 314
  var emailsToShare = [ 
    "moira.xu@shopee.com", 
    "nonilon.macasadia@shopee.com", 
    "billings.ph@shopee.com",
    "mantillasa@sea.com", 
    "ongt@sea.com", 
    "charles.duelas@shopee.com", 
    "sellers.ph@shopee.com"
  ];
  
  // Check if the edit overlaps with our target column and row
  if (range.getColumn() <= targetColumn && (range.getLastColumn() >= targetColumn) && range.getLastRow() >= startRow) {
    
    var values = range.getValues();
    var startRowEdit = range.getRow();
    
    // Loop through each row that was edited
    for (var i = 0; i < values.length; i++) {
      var currentRow = startRowEdit + i;
      
      if (currentRow < startRow) continue;
      
      var rowRange = sheet.getRange(currentRow, targetColumn);
      var folderName = rowRange.getValue();
      
      if (folderName === "") continue;
      
      // Check if a link already exists in Column T
      var statusCell = sheet.getRange(currentRow, targetColumn + 1);
      if (statusCell.getValue() !== "") continue; 
      
      try {
        // Create the folder
        var newFolder = DriveApp.createFolder(folderName);
        var folderId = newFolder.getId();
        
        // Share the folder silently with each person
        emailsToShare.forEach(function(email) {
          try {
            // Using Drive API v3 syntax (.create instead of .insert)
            Drive.Permissions.create(
              {
                'role': 'writer', // Options: 'writer' (editor) or 'reader' (viewer)
                'type': 'user',
                'emailAddress': email // v3 uses 'emailAddress' instead of 'value'
              },
              folderId,
              {
                'sendNotificationEmail': false, // Note: no 's' at the end for v3
                'supportsAllDrives': true       // Crucial for organizational/shared drive environments
              }
            );
          } catch(err) {
            console.log("Failed to share with " + email + ": " + err.toString());
          }
        });
        
        // Write the unique link in Column T
        statusCell.setValue(newFolder.getUrl());
        
      } catch(error) {
        statusCell.setValue("Error: " + error.toString());
      }
    }
  }
}
