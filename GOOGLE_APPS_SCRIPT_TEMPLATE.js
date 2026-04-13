// Google Apps Script Template for Application CMS Integration
// Deploy as Web App with "Execute as me" and "Anyone has access"

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const secret = PropertiesService.getScriptProperties().getProperty('SHARED_SECRET');
    
    // Validate secret
    if (data.secret !== secret) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Invalid secret'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    switch(data.action) {
      case 'test':
        return testConnection();
      case 'upsertCase':
        return upsertCase(data.caseData, data.sheetId);
      case 'appendAudit':
        return appendAudit(data.auditData, data.sheetId);
      case 'uploadFile':
        return uploadFile(data);
      default:
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'Unknown action'
        })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function testConnection() {
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'Connection successful'
  })).setMimeType(ContentService.MimeType.JSON);
}

function upsertCase(caseData, sheetId) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Cases');
    if (!sheet) {
      // Create Cases sheet if it doesn't exist
      const ss = SpreadsheetApp.openById(sheetId);
      sheet = ss.insertSheet('Cases');
      // Add headers
      sheet.appendRow([
        'ID', 'Date/Time', 'Case No', 'Application Name', 'TM No', 
        'Class', 'Status', 'Sub Status', 'IPO Office', 'Delivery Tracking',
        'Notes', 'Created', 'Updated', 'Sync Status'
      ]);
    }
    
    // Find existing row by Case No
    const caseNo = caseData.case_no;
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    let existingRowIndex = -1;
    
    for (let i = 1; i < values.length; i++) { // Skip header row
      if (values[i][2] === caseNo) { // Case No is in column 3 (index 2)
        existingRowIndex = i + 1; // +1 because sheets are 1-indexed
        break;
      }
    }
    
    const rowData = [
      caseData.id,
      caseData.date_time,
      caseData.case_no,
      caseData.app_name,
      caseData.tm_no,
      caseData.class_no,
      caseData.status,
      caseData.sub_status,
      caseData.ipo_office,
      caseData.delivery_tracking_no,
      caseData.notes,
      caseData.created,
      caseData.updated,
      'synced'
    ];
    
    if (existingRowIndex > 0) {
      // Update existing row
      sheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // Append new row
      sheet.appendRow(rowData);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: existingRowIndex > 0 ? 'Case updated' : 'Case created'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function appendAudit(auditData, sheetId) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Audit');
    if (!sheet) {
      // Create Audit sheet if it doesn't exist
      const ss = SpreadsheetApp.openById(sheetId);
      sheet = ss.insertSheet('Audit');
      // Add headers
      sheet.appendRow([
        'Case ID', 'Action', 'Old Values', 'New Values', 'Timestamp', 'Actor'
      ]);
    }
    
    const rowData = [
      auditData.case_id,
      auditData.action,
      auditData.old_values,
      auditData.new_values,
      auditData.timestamp,
      auditData.actor
    ];
    
    sheet.appendRow(rowData);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Audit entry added'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function uploadFile(data) {
  try {
    const folderId = data.driveFolderId;
    const fileName = data.fileName;
    const fileData = Utilities.base64Decode(data.fileData);
    const mimeType = data.fileType;
    
    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(fileData, mimeType, fileName);
    
    const file = folder.createFile(blob);
    
    // Share file with anyone who has link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      fileId: file.getId(),
      fileUrl: file.getUrl()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Setup function - run once to set the shared secret
function setup() {
  const secret = 'YOUR_SECRET_KEY_HERE'; // Change this to a secure random string
  PropertiesService.getScriptProperties().setProperty('SHARED_SECRET', secret);
  Logger.log('Shared secret set: ' + secret);
}
