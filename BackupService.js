/**
 * ==========================================================
 * SMOKY SWAAD ERP
 * AUTOMATIC DATABASE BACKUP SERVICE
 * ==========================================================
 */

function runDailyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupFolderName = "Smoky Swaad Backups";
  
  // 1. Find or create the backup folder in Google Drive
  const folders = DriveApp.getFoldersByName(backupFolderName);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(backupFolderName);
  }
  
  // 2. Format the current date and time for the filename
  const timeZone = Session.getScriptTimeZone();
  const dateString = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  const backupName = "ERP_Backup_" + dateString;
  
  // 3. Make a complete copy of the live database
  DriveApp.getFileById(ss.getId()).makeCopy(backupName, folder);
  
  // 4. Clean up old backups to save Google Drive storage (Deletes older than 30 days)
  cleanUpOldBackups(folder);
}

/**
 * Helper function to keep Drive storage clean
 */
function cleanUpOldBackups(folder) {
  const maxDays = 15; 
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - maxDays);
  
  const files = folder.getFiles();
  
  while (files.hasNext()) {
    const file = files.next();
    // If the file was created before our 15-day cutoff, delete it
    if (file.getDateCreated() < cutoffDate) {
      file.setTrashed(true);
    }
  }
}





function createBackupTrigger() {
  // Clear any existing triggers for this function to prevent duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDailyBackup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create a new time-driven trigger to run every day at 5:00 AM
  ScriptApp.newTrigger('runDailyBackup')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();
    
  Logger.log("Daily backup trigger has been successfully set up for 5:00 AM!");
}