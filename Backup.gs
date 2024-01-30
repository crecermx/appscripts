function backup() {
  const destinationFolder = DriveApp.getFolderById("1AsjUPCEzATtgAhepiB59WSBd57MOkIsH");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const name = spreadsheet.getName() + "-" + Utilities.formatDate(new Date(), "GMT-6", "yyyy-MM-dd");

  DriveApp.getFileById(spreadsheet.getId()).makeCopy(name, destinationFolder);
}
