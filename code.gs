/**
 * GSC Backup Manager - Google Apps Script
 * Unified version with both manual import and automated backup functionality
 */

/**
 * Checks if email permissions are granted
 */
function checkEmailPermissions() {
  try {
    const email = Session.getEffectiveUser().getEmail();
    if (!email) {
      return {
        hasPermission: false,
        message: "Could not determine user email address"
      };
    }
    
    MailApp.getRemainingDailyQuota();
    return {
      hasPermission: true,
      message: "Email permissions are granted",
      email: email
    };
  } catch (error) {
    return {
      hasPermission: false,
      message: "Email permissions not granted: " + error.message
    };
  }
}

/**
 * Creates the menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('GSC Backup Manager')
    .addItem('Manage Backups', 'showSidebar');
  
  try {
    const testSS = SpreadsheetApp.create('Permission_Test_' + Date.now());
    DriveApp.getFileById(testSS.getId()).setTrashed(true);
    ScriptApp.getProjectTriggers();
    getWebsites();
    Session.getEffectiveUser().getEmail();
    menu.addToUi();
  } catch (e) {
    menu.addItem('Authorize All Permissions', 'requestTriggerAuthorization')
        .addToUi();
  }
}

/**
 * Shows authorization dialog
 */
function requestTriggerAuthorization() {
  const html = HtmlService.createHtmlOutput(`
    <p>To enable backup functionality, please authorize the app with all required permissions:</p>
    <button onclick="google.script.run.authorizeAllPermissions()">Authorize All Permissions</button>
  `).setWidth(350).setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Authorization Required');
}

/**
 * Forces authorization for all required scopes
 */
function authorizeAllPermissions() {
  try {
    console.log('Requesting authorization for all required scopes...');
    
    try {
      const testSS = SpreadsheetApp.create('Auth_Test_' + Date.now());
      DriveApp.getFileById(testSS.getId()).setTrashed(true);
      console.log('SpreadsheetApp authorization successful');
    } catch (e) {
      console.log('SpreadsheetApp authorization needed');
    }
    
    try {
      DriveApp.getRootFolder();
      console.log('DriveApp authorization successful');
    } catch (e) {
      console.log('DriveApp authorization needed');
    }
    
    try {
      ScriptApp.getProjectTriggers();
      console.log('ScriptApp authorization successful');
    } catch (e) {
      console.log('ScriptApp authorization needed');
    }
    
    try {
      Session.getEffectiveUser().getEmail();
      console.log('MailApp authorization successful');
    } catch (e) {
      console.log('MailApp authorization needed');
    }
    
    try {
      getWebsites();
      console.log('Search Console API authorization successful');
    } catch (e) {
      console.log('Search Console API authorization needed');
    }
    
    return 'Authorization process completed. Please refresh the page and try again.';
    
  } catch (error) {
    console.error('Authorization error:', error);
    return 'Authorization failed: ' + error.message;
  }
}

/**
 * Shows the sidebar with the backup interface
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('GSC Backup Manager')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Gets list of verified websites from Search Console
 */
function getWebsites() {
  try {
    const url = 'https://www.googleapis.com/webmasters/v3/sites';
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });

    const raw = response.getContentText();
    console.log('Raw websites response:', raw);
    const data = JSON.parse(raw);

    if (data.siteEntry && data.siteEntry.length > 0) {
      return data.siteEntry.map(site => ({
        siteUrl: site.siteUrl,
        permissionLevel: site.permissionLevel
      }));
    }

    return [];
  } catch (error) {
    console.error('Error fetching websites:', error);
    throw new Error('Failed to load websites. Please ensure you have access to Search Console.');
  }
}

/**
 * Gets list of sheets in the current spreadsheet
 */
function getSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheets().map(sheet => sheet.getName());
  } catch (error) {
    console.error('Error fetching sheets:', error);
    throw new Error('Failed to load sheets.');
  }
}

/**
 * Gets current date in YYYY-MM-DD format
 */
function getCurrentDate() {
  const today = new Date();
  return Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Gets date 30 days ago in YYYY-MM-DD format
 */
function getThirtyDaysAgo() {
  const date = new Date();
  date.setDate(date.getDate() - 30);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Imports Search Console data based on user selections
 */
function importGSCData(website, startDate, endDate, dimensions, searchType, rowLimit, clearSheet, sheetName, filters, aggregationType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
    
    return importGSCDataToSheet(sheet, website, startDate, endDate, dimensions, searchType, rowLimit, clearSheet, filters, aggregationType);
    
  } catch (error) {
    console.error('Error in importGSCData:', error);
    throw error;
  }
}

/**
 * Helper function to import GSC data directly to a sheet object
 * This is the core import function used by both manual imports and backups
 */
function importGSCDataToSheet(sheet, website, startDate, endDate, dimensions, searchType, rowLimit, clearSheet, filters, aggregationType) {
  try {
    console.log('Import to sheet parameters:', {
      website, startDate, endDate, dimensions, searchType, rowLimit, clearSheet
    });
    
    // Validate inputs more thoroughly
    if (!website || website.trim() === '') {
      throw new Error('Website URL is required');
    }
    
    if (!startDate || !endDate) {
      throw new Error('Start date and end date are required');
    }
    
    // Validate date format and range
    const start = new Date(startDate);
    const end = new Date(endDate);
    const today = new Date();
    
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      throw new Error('Invalid date format. Please use YYYY-MM-DD format.');
    }
    
    if (start > end) {
      throw new Error('Start date must be before or equal to end date');
    }
    
    if (end > today) {
      throw new Error('End date cannot be in the future');
    }
    
    // Check if date range is too far back (GSC has ~16 month limit)
    const sixteenMonthsAgo = new Date();
    sixteenMonthsAgo.setMonth(sixteenMonthsAgo.getMonth() - 16);
    if (start < sixteenMonthsAgo) {
      throw new Error('Start date is too far back. Google Search Console data is limited to approximately 16 months.');
    }
    
    if (!dimensions || dimensions.length === 0) {
      throw new Error('At least one dimension must be selected');
    }
    
    // Validate dimensions
    const validDimensions = ['query', 'page', 'country', 'device', 'searchAppearance'];
    const invalidDimensions = dimensions.filter(dim => !validDimensions.includes(dim));
    if (invalidDimensions.length > 0) {
      throw new Error(`Invalid dimensions: ${invalidDimensions.join(', ')}`);
    }
    
    // Validate search type
    const validSearchTypes = ['web', 'image', 'video', 'news', 'googleNews', 'discover'];
    if (searchType && !validSearchTypes.includes(searchType)) {
      throw new Error(`Invalid search type: ${searchType}. Valid types are: ${validSearchTypes.join(', ')}`);
    }
    
    // Prepare the request payload
    const requestPayload = {
      startDate: startDate,
      endDate: endDate,
      dimensions: dimensions,
      rowLimit: Math.min(rowLimit || 25000, 25000), // GSC API limit is 25,000
      startRow: 0,
      type: searchType || 'web',
      dimensionFilterGroups: []
    };

    if (aggregationType && aggregationType !== 'auto') {
      requestPayload.aggregationType = aggregationType;
    }
    
    // Add filters if specified
    const filterGroup = {
      filters: []
    };
    
    if (filters && filters.query) {
      filterGroup.filters.push({
        dimension: 'query',
        operator: 'equals',
        expression: filters.query
      });
    }
    
    if (filters && filters.page) {
      filterGroup.filters.push({
        dimension: 'page',
        operator: 'equals',
        expression: filters.page
      });
    }
    
    if (filters && filters.country) {
      filterGroup.filters.push({
        dimension: 'country',
        operator: 'equals',
        expression: filters.country
      });
    }
    
    if (filters && filters.device) {
      filterGroup.filters.push({
        dimension: 'device',
        operator: 'equals',
        expression: filters.device
      });
    }
    
    if (filterGroup.filters.length > 0) {
      requestPayload.dimensionFilterGroups = [filterGroup];
    }
    console.log('Request payload:', JSON.stringify(requestPayload, null, 2));
    
    // Clean and encode the website URL
    let cleanWebsite = website.trim();
    if (!cleanWebsite.startsWith('http://') && !cleanWebsite.startsWith('https://')) {
      cleanWebsite = 'https://' + cleanWebsite;
    }
    
    const encodedWebsite = encodeURIComponent(cleanWebsite);
    const url = `https://www.googleapis.com/webmasters/v3/sites/${encodedWebsite}/searchAnalytics/query`;
    
    console.log('API URL:', url);
    
    // Make the API call
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestPayload)
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('API Response Code:', responseCode);
    console.log('API Response:', responseText);
    
    if (responseCode !== 200) {
      let errorMessage = `API request failed with status ${responseCode}`;
      try {
        const errorData = JSON.parse(responseText);
        if (errorData.error && errorData.error.message) {
          errorMessage += `: ${errorData.error.message}`;
        }
      } catch (parseError) {
        errorMessage += `: ${responseText}`;
      }
      throw new Error(errorMessage);
    }
    
    const data = JSON.parse(responseText);
    
    if (!data.rows || data.rows.length === 0) {
      const noDataMessage = 'No data found for the selected period and filters';
      // Clear sheet if requested, then add no data message
      if (clearSheet) {
        console.log('Clearing sheet content...');
        clearSheetContent(sheet);
      }
      sheet.getRange(1, 1).setValue(noDataMessage);
      return noDataMessage;
    }
    
    // Prepare headers
    const headers = [...dimensions, 'Date and Time', 'Clicks', 'Impressions', 'CTR', 'Position'];
    
    // Prepare data rows
    const rows = data.rows.map(row => {
      const dimensionValues = row.keys || [];
      const currentDateTime = new Date().toLocaleString();
      return [
        ...dimensionValues,
        currentDateTime,
        row.clicks || 0,
        row.impressions || 0,
        row.ctr ? (row.ctr * 100).toFixed(2) + '%' : '0%',
        row.position ? row.position.toFixed(1) : 0
      ];
    });
    
    // Determine where to write the data
    let startRow = 1;
    let startCol = 1;
    
    if (clearSheet) {
      // Clear all existing content first
      console.log('Clearing sheet content...');
      clearSheetContent(sheet);
      startRow = 1; // Start from the top
    } else {
      // Check if we should append or replace
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow === 0) {
        // Empty sheet, start from top
        startRow = 1;
      } else {
        // Check if first row contains headers that match our current headers
        const existingFirstRow = sheet.getRange(1, 1, 1, Math.min(lastCol, headers.length)).getValues()[0];
        const headersMatch = headers.every((header, index) => 
          existingFirstRow[index] && existingFirstRow[index].toString() === header.toString()
        );
        
        if (headersMatch) {
          // Headers match, append new data (skip writing headers again)
          startRow = lastRow + 1;
          const dataOnly = rows;
          const range = sheet.getRange(startRow, startCol, dataOnly.length, headers.length);
          range.setValues(dataOnly);
          
          // Format the new data rows only
          formatDataRows(sheet, headers.length, startRow, dataOnly.length);
          
          return `Successfully imported ${rows.length} rows of data (Search Type: ${searchType || 'web'}) - Data appended to existing data`;
        } else {
          // Headers don't match or sheet has different structure, clear and start fresh
          console.log('Headers do not match existing data, clearing sheet...');
          clearSheetContent(sheet);
          startRow = 1;
        }
      }
    }
    
    // Write headers and data to sheet (for new/cleared sheets)
    const allData = [headers, ...rows];
    const range = sheet.getRange(startRow, startCol, allData.length, headers.length);
    range.setValues(allData);
    
    // Format the sheet
    formatSheet(sheet, headers.length, allData.length);
    
    return `Successfully imported ${rows.length} rows of data (Search Type: ${searchType || 'web'})`;
    
  } catch (error) {
    console.error('Error importing data to sheet:', error);
    
    // More specific error handling
    const errorMessage = error.message || error.toString();
    
    if (errorMessage.includes('403') || errorMessage.includes('Forbidden')) {
      throw new Error('Access denied. Please ensure you have permission to access this website in Search Console and that the URL exactly matches your verified property.');
    } else if (errorMessage.includes('400') || errorMessage.includes('Bad Request')) {
      throw new Error(`Invalid request: ${errorMessage}. Please check your website URL, date range, and selected dimensions.`);
    } else if (errorMessage.includes('401') || errorMessage.includes('Unauthorized')) {
      throw new Error('Authentication failed. Please try refreshing the page and running the import again.');
    } else if (errorMessage.includes('404') || errorMessage.includes('Not Found')) {
      throw new Error('Website not found in Search Console. Please verify the URL exactly matches your verified property.');
    } else {
      throw new Error(`Import failed: ${errorMessage}`);
    }
  }
}

/**
 * Safely clears sheet content without destroying the sheet structure
 */
function clearSheetContent(sheet) {
  try {
    // Get the data range to clear
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow > 0 && lastCol > 0) {
      // Clear content and formatting from the used range
      const range = sheet.getRange(1, 1, lastRow, lastCol);
      range.clearContent();
      range.clearFormat();
      range.clearDataValidations();
      range.clearNote();
    }
    
    // Reset any row/column sizes that might have been set
    sheet.setRowHeights(1, Math.max(1, lastRow), 21); // Default row height
    
  } catch (error) {
    console.error('Error clearing sheet content:', error);
    // If the safe clear fails, fall back to sheet.clear() but handle any errors
    try {
      sheet.clear();
    } catch (clearError) {
      console.error('Error with fallback clear:', clearError);
      throw new Error('Unable to clear sheet content');
    }
  }
}

/**
 * Formats the sheet with headers and data styling
 */
function formatSheet(sheet, numCols, numRows) {
  try {
    // Ensure we have valid dimensions
    if (numRows < 1 || numCols < 1) {
      console.log('Invalid dimensions for formatting:', numRows, numCols);
      return;
    }
    
    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, numCols);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // Auto-resize columns
    for (let i = 1; i <= numCols; i++) {
      try {
        sheet.autoResizeColumn(i);
      } catch (resizeError) {
        console.error(`Error resizing column ${i}:`, resizeError);
      }
    }
    
    // Add borders
    if (numRows > 0 && numCols > 0) {
      const dataRange = sheet.getRange(1, 1, numRows, numCols);
      dataRange.setBorder(true, true, true, true, true, true);
    }
    
    // Format numeric columns if we have enough columns and rows
    if (numCols >= 5 && numRows > 1) {
      formatDataRows(sheet, numCols, 2, numRows - 1);
    }
  } catch (error) {
    console.error('Error formatting sheet:', error);
    // Don't throw error for formatting issues, just log them
  }
}

/**
 * Formats data rows (without headers)
 */
function formatDataRows(sheet, numCols, startRow, numDataRows) {
  try {
    if (numCols >= 5 && numDataRows > 0) {
      // Clicks column (second to last minus 3, accounting for Date and Time column)
      const clicksRange = sheet.getRange(startRow, numCols - 3, numDataRows, 1);
      clicksRange.setNumberFormat('#,##0');
      
      // Impressions column (second to last minus 2)
      const impressionsRange = sheet.getRange(startRow, numCols - 2, numDataRows, 1);
      impressionsRange.setNumberFormat('#,##0');
      
      // Position column (last column)
      const positionRange = sheet.getRange(startRow, numCols, numDataRows, 1);
      positionRange.setNumberFormat('0.0');
      
      // Add borders to the new data
      const dataRange = sheet.getRange(startRow, 1, numDataRows, numCols);
      dataRange.setBorder(true, true, true, true, true, true);
    }
  } catch (formatError) {
    console.error('Error formatting data rows:', formatError);
  }
} 

/**
 * Sets up automated backups
 */
function setupBackup(website, backupType, dimensions, searchType, separateUngrouped, emailNotification) {
  try {
    // First verify we have trigger permissions
    try {
      ScriptApp.getProjectTriggers();
    } catch (authError) {
      throw new Error('Trigger permissions not granted. Please use the "Authorize Backup Functionality" menu option first.');
    }

    // Validate inputs
    if (!website || website.trim() === '') {
      throw new Error('Website URL is required');
    }
    
    if (!dimensions || dimensions.length === 0) {
      throw new Error('At least one dimension must be selected');
    }
    
    // Clean up existing triggers for this website and backup type combination
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runBackup') {
        try {
          const triggerUid = trigger.getUniqueId();
          const storedData = PropertiesService.getScriptProperties().getProperty('backup_' + triggerUid);
          if (storedData) {
            const triggerData = JSON.parse(storedData);
            if (triggerData.website === website && triggerData.backupType === backupType) {
              ScriptApp.deleteTrigger(trigger);
              PropertiesService.getScriptProperties().deleteProperty('backup_' + triggerUid);
              console.log('Removed existing backup for same website and type');
            }
          }
        } catch (e) {
          console.error('Error checking/deleting trigger:', e);
          ScriptApp.deleteTrigger(trigger);
        }
      }
    });
    
    // Create new trigger
    let newTrigger;
    if (backupType === 'daily') {
      newTrigger = ScriptApp.newTrigger('runBackup')
        .timeBased()
        .everyDays(1)
        .atHour(2)
        .create();
    } else {
      newTrigger = ScriptApp.newTrigger('runBackup')
        .timeBased()
        .onMonthDay(3)
        .atHour(2)
        .create();
    }
    
    // Store backup configuration
    const triggerUid = newTrigger.getUniqueId();
    const triggerData = {
      website: website,
      dimensions: dimensions,
      searchType: searchType || 'web',
      separateUngrouped: separateUngrouped || false,
      emailNotification: emailNotification || false,
      backupType: backupType,
      createdAt: new Date().toISOString(),
      status: 'active' // Add status tracking
    };
    
    PropertiesService.getScriptProperties().setProperty('backup_' + triggerUid, JSON.stringify(triggerData));
    
    // Return success message
    const nextRun = backupType === 'daily' ? 'tomorrow at 2 AM' : 'on the 3rd of next month at 2 AM';
    return `Successfully created ${backupType} backup for ${website}. Next backup will run ${nextRun}.`;
    
  } catch (error) {
    console.error('Error setting up backup:', error);
    throw new Error('Failed to setup backup: ' + error.message);
  }
}

/**
 * Runs the actual backup - Modified to fetch latest available data from GSC
 */
function runBackup(e) {
  let triggerData = null;
  
  try {
    console.log('Running backup triggered by:', e);
    
    // Get trigger UID from the event
    const triggerUid = e && e.triggerUid ? e.triggerUid : null;
    
    if (!triggerUid) {
      throw new Error('No trigger UID found in event');
    }
    
    // Get backup configuration from PropertiesService
    const storedData = PropertiesService.getScriptProperties().getProperty('backup_' + triggerUid);
    if (!storedData) {
      throw new Error('No backup configuration found for trigger: ' + triggerUid);
    }
    
    triggerData = JSON.parse(storedData);
    console.log('Retrieved trigger data:', triggerData);
    
    const {
      website,
      dimensions,
      searchType,
      separateUngrouped,
      emailNotification,
      backupType
    } = triggerData;
    
    // Validate required fields
    if (!website) {
      throw new Error('Website URL is missing from backup configuration');
    }
    
    if (!dimensions || dimensions.length === 0) {
      throw new Error('No dimensions specified in backup configuration');
    }
    
    // Calculate date range - GET LATEST AVAILABLE DATA
    const now = new Date();
    let endDate, startDate;
    
    if (backupType === 'daily') {
      // For daily backups, find the most recent available data
      const latestDataDate = findLatestAvailableDate(website, searchType);
      
      if (latestDataDate) {
        startDate = latestDataDate;
        endDate = latestDataDate;
        console.log('Found latest available data date:', latestDataDate);
      } else {
        // Fallback to yesterday if we can't determine latest date
        const yesterday = new Date(now);
        yesterday.setDate(yesterday.getDate() - 1);
        startDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        endDate = startDate;
        console.log('Using fallback date (yesterday):', startDate);
      }
      
    } else {
      // For monthly backups, get the most recent complete month
      const lastCompleteMonth = findLatestCompleteMonth(website, searchType);
      
      if (lastCompleteMonth) {
        startDate = lastCompleteMonth.start;
        endDate = lastCompleteMonth.end;
        console.log('Found latest complete month:', startDate, 'to', endDate);
      } else {
        // Fallback to last month
        const lastMonth = new Date(now);
        lastMonth.setMonth(lastMonth.getMonth() - 1);
        
        const firstDay = new Date(lastMonth.getFullYear(), lastMonth.getMonth(), 1);
        startDate = Utilities.formatDate(firstDay, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        const lastDay = new Date(lastMonth.getFullYear(), lastMonth.getMonth() + 1, 0);
        endDate = Utilities.formatDate(lastDay, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        console.log('Using fallback month:', startDate, 'to', endDate);
      }
    }
    
    console.log('Backup date range:', startDate, 'to', endDate);
    
    // Test API access before proceeding
    console.log('Testing API access before backup...');
    try {
      const testWebsites = getWebsites();
      const hasAccess = testWebsites.some(site => site.siteUrl === website);
      if (!hasAccess) {
        throw new Error(`No access to website: ${website}. Please verify the URL and permissions.`);
      }
      console.log('API access test passed');
    } catch (apiError) {
      console.error('API access test failed:', apiError);
      throw new Error(`API access failed: ${apiError.message}. Please check your Search Console permissions.`);
    }
    
    // Create backup folder if it doesn't exist
    console.log('Creating backup folder...');
    const folderName = 'GSC Backups';
    let folder;
    try {
      const folders = DriveApp.getFoldersByName(folderName);
      
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder(folderName);
      }
      console.log('Backup folder ready');
    } catch (folderError) {
      console.error('Error creating backup folder:', folderError);
      throw new Error('Failed to create backup folder: ' + folderError.message);
    }
    
    // Create spreadsheet for this backup
    console.log('Creating backup spreadsheet...');
    const dateString = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const cleanWebsiteName = website.replace(/[^\w\-\.]/g, '_');
    const dataRangeString = startDate === endDate ? startDate : `${startDate}_to_${endDate}`;
    const ssName = `GSC_Latest_${cleanWebsiteName}_${dataRangeString}_${dateString}`;
    
    let ss;
    try {
      ss = SpreadsheetApp.create(ssName);
      console.log('Backup spreadsheet created:', ssName);
    } catch (ssError) {
      console.error('Error creating spreadsheet:', ssError);
      throw new Error('Failed to create backup spreadsheet: ' + ssError.message);
    }
    
    // Move file to backup folder
    try {
      const file = DriveApp.getFileById(ss.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
      console.log('File moved to backup folder');
    } catch (moveError) {
      console.error('Error moving file to backup folder:', moveError);
      // Don't fail the backup for this, just log the error
    }
    
    // Run the import on the first sheet
    const sheet = ss.getSheets()[0];
    sheet.setName('GSC_Latest_Data');
    
    console.log('Running main backup import...');
    const result = importGSCDataToSheet(
      sheet,
      website,
      startDate,
      endDate,
      dimensions,
      searchType || 'web',
      25000,
      true,
      {}, // No filters for backup
      'auto' // Default aggregation
    );
    
    console.log('Main backup result:', result);
    
    // If separateUngrouped is true, create another sheet with ungrouped data
    if (separateUngrouped) {
      console.log('Creating ungrouped backup...');
      try {
        const ungroupedSheet = ss.insertSheet('GSC_Latest_Ungrouped');
        
        // Run import with all dimensions for ungrouped data
        const ungroupedResult = importGSCDataToSheet(
          ungroupedSheet,
          website,
          startDate,
          endDate,
          ['query', 'page', 'country', 'device'], // All dimensions for ungrouped
          searchType || 'web',
          25000,
          true,
          {}, // No filters
          'auto' // Default aggregation
        );
        
        console.log('Ungrouped backup result:', ungroupedResult);
      } catch (ungroupedError) {
        console.error('Error creating ungrouped backup:', ungroupedError);
        // Don't fail the entire backup for this, just log the error
      }
    }
    
    // Send email notification if enabled
    if (emailNotification) {
      console.log('Preparing to send email notification...');
      try {
        const userEmail = Session.getEffectiveUser().getEmail();
        if (!userEmail) {
          console.error('No email address found for user');
          throw new Error('Could not determine user email address');
        }

        const subject = `GSC Latest Data Backup Completed: ${website}`;
        const body = `The GSC backup for ${website} has completed successfully.\n\n` +
                     `Backup Type: ${backupType}\n` +
                     `Date Range: ${startDate} to ${endDate}\n` +
                     `Data Period: Latest available data from Google Search Console\n` +
                     `Dimensions: ${dimensions.join(', ')}\n` +
                     `Search Type: ${searchType || 'web'}\n` +
                     `Separate Ungrouped: ${separateUngrouped ? 'Yes' : 'No'}\n\n` +
                     `Backup File: ${ss.getUrl()}\n\n` +
                     `Result: ${result}`;
        
        console.log('Sending email to:', userEmail);
        console.log('Email subject:', subject);
        console.log('Email body:', body);
        
        MailApp.sendEmail({
          to: userEmail,
          subject: subject,
          body: body
        });
        console.log('Email notification sent successfully');
      } catch (emailError) {
        console.error('Failed to send email notification:', emailError);
        // Store the error to show in the UI
        triggerData.emailError = emailError.message;
        PropertiesService.getScriptProperties().setProperty('backup_' + triggerUid, JSON.stringify(triggerData));
      }
    }
    
    console.log('Backup completed successfully');
    return `Latest data backup completed successfully for ${website}. Data range: ${startDate} to ${endDate}. File: ${ss.getUrl()}`;
    
  } catch (error) {
    console.error('Error running backup:', error);
    
    // Send error notification if email notifications are enabled and we have trigger data
    if (triggerData && triggerData.emailNotification) {
      try {
        const subject = `GSC Latest Data Backup Failed: ${triggerData.website}`;
        const body = `The GSC backup for ${triggerData.website} failed with error:\n\n${error.message}\n\n` +
                     `Error details: ${error.toString()}\n\n` +
                     `Time: ${new Date().toISOString()}`;
        
        MailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, body);
        console.log('Error notification email sent');
      } catch (emailError) {
        console.error('Failed to send error notification email:', emailError);
      }
    }
    
    throw error;
  }
}

/**
 * Helper function to find the latest available date with data
 */
function findLatestAvailableDate(website, searchType) {
  try {
    const now = new Date();
    
    // Check dates from 2 days ago going back to 7 days ago
    for (let daysBack = 2; daysBack <= 7; daysBack++) {
      const checkDate = new Date(now);
      checkDate.setDate(checkDate.getDate() - daysBack);
      const dateString = Utilities.formatDate(checkDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      console.log('Checking for data on:', dateString);
      
      // Test if this date has data by making a small API call
      if (hasDataForDate(website, dateString, searchType)) {
        console.log('Found data for date:', dateString);
        return dateString;
      }
    }
    
    console.log('No recent data found, using fallback');
    return null; // Will use fallback date
    
  } catch (error) {
    console.error('Error finding latest available date:', error);
    return null; // Use fallback
  }
}

/**
 * Helper function to find the latest complete month with data
 */
function findLatestCompleteMonth(website, searchType) {
  try {
    const now = new Date();
    
    // Check last 3 months
    for (let monthsBack = 1; monthsBack <= 3; monthsBack++) {
      const checkMonth = new Date(now);
      checkMonth.setMonth(checkMonth.getMonth() - monthsBack);
      
      const firstDay = new Date(checkMonth.getFullYear(), checkMonth.getMonth(), 1);
      const lastDay = new Date(checkMonth.getFullYear(), checkMonth.getMonth() + 1, 0);
      
      const startDate = Utilities.formatDate(firstDay, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const endDate = Utilities.formatDate(lastDay, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      console.log('Checking for month data:', startDate, 'to', endDate);
      
      // Test if this month has data
      if (hasDataForDateRange(website, startDate, endDate, searchType)) {
        console.log('Found complete month data:', startDate, 'to', endDate);
        return { start: startDate, end: endDate };
      }
    }
    
    console.log('No recent complete month found, using fallback');
    return null; // Will use fallback dates
    
  } catch (error) {
    console.error('Error finding latest complete month:', error);
    return null; // Use fallback
  }
}

/**
 * Helper function to check if a specific date has data
 */
function hasDataForDate(website, dateString, searchType) {
  try {
    // Make a minimal API call to check for data
    const requestPayload = {
      startDate: dateString,
      endDate: dateString,
      dimensions: ['query'], // Minimal dimension
      rowLimit: 1, // Just need to know if data exists
      startRow: 0,
      type: searchType || 'web'
    };
    
    const cleanWebsite = website.trim();
    const encodedWebsite = encodeURIComponent(cleanWebsite);
    const url = `https://www.googleapis.com/webmasters/v3/sites/${encodedWebsite}/searchAnalytics/query`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestPayload)
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return data.rows && data.rows.length > 0;
    }
    
    return false;
    
  } catch (error) {
    console.error('Error checking data for date:', dateString, error);
    return false;
  }
}

/**
 * Helper function to check if a date range has data
 */
function hasDataForDateRange(website, startDate, endDate, searchType) {
  try {
    // Make a minimal API call to check for data in the range
    const requestPayload = {
      startDate: startDate,
      endDate: endDate,
      dimensions: ['query'], // Minimal dimension
      rowLimit: 1, // Just need to know if data exists
      startRow: 0,
      type: searchType || 'web'
    };
    
    const cleanWebsite = website.trim();
    const encodedWebsite = encodeURIComponent(cleanWebsite);
    const url = `https://www.googleapis.com/webmasters/v3/sites/${encodedWebsite}/searchAnalytics/query`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestPayload)
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return data.rows && data.rows.length > 0;
    }
    
    return false;
    
  } catch (error) {
    console.error('Error checking data for date range:', startDate, 'to', endDate, error);
    return false;
  }
} 

/**
 * Gets a list of all backup configurations with their status
 */
function getBackupList() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const backupList = [];
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runBackup') {
        const triggerUid = trigger.getUniqueId();
        const storedData = PropertiesService.getScriptProperties().getProperty('backup_' + triggerUid);
        
        if (storedData) {
          try {
            const triggerData = JSON.parse(storedData);
            
            // Determine next run time
            let nextRun = 'Unknown';
            if (triggerData.status !== 'paused') {
              if (triggerData.backupType === 'daily') {
                nextRun = 'Daily at 2:00 AM';
              } else if (triggerData.backupType === 'monthly') {
                nextRun = 'Monthly on 3rd at 2:00 AM';
              }
            }
            
            backupList.push({
              triggerUid: triggerUid,
              website: triggerData.website,
              backupType: triggerData.backupType,
              dimensions: triggerData.dimensions || [],
              searchType: triggerData.searchType || 'web',
              separateUngrouped: triggerData.separateUngrouped || false,
              emailNotification: triggerData.emailNotification || false,
              createdAt: triggerData.createdAt,
              status: triggerData.status || 'active',
              nextRun: nextRun
            });
          } catch (parseError) {
            console.error('Error parsing trigger data for UID:', triggerUid, parseError);
          }
        }
      }
    });
    
    // Sort by creation date (newest first)
    backupList.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    
    console.log('Retrieved backup list:', backupList);
    return backupList;
    
  } catch (error) {
    console.error('Error getting backup list:', error);
    throw new Error('Failed to load backup list: ' + error.message);
  }
}

/**
 * Runs a backup manually by trigger UID
 */
function runBackupManually(triggerUid) {
  try {
    console.log('Running backup manually for trigger UID:', triggerUid);
    
    // Add timeout protection
    const startTime = new Date();
    const maxExecutionTime = 5 * 60 * 1000; // 5 minutes max
    
    // Get backup configuration
    const storedData = PropertiesService.getScriptProperties().getProperty('backup_' + triggerUid);
    if (!storedData) {
      throw new Error('No backup configuration found for the specified backup');
    }
    
    const triggerData = JSON.parse(storedData);
    console.log('Retrieved trigger data:', triggerData);
    
    // Check if backup is paused
    if (triggerData.status === 'paused') {
      throw new Error('Cannot run a paused backup. Please resume it first.');
    }
    
    // Validate trigger data
    if (!triggerData.website) {
      throw new Error('Invalid backup configuration: website is missing');
    }
    
    if (!triggerData.dimensions || triggerData.dimensions.length === 0) {
      throw new Error('Invalid backup configuration: no dimensions specified');
    }
    
    console.log('Manual backup trigger data:', triggerData);
    
    // Test API access first
    console.log('Testing API access...');
    try {
      const testWebsites = getWebsites();
      const hasAccess = testWebsites.some(site => site.siteUrl === triggerData.website);
      if (!hasAccess) {
        throw new Error(`No access to website: ${triggerData.website}. Please verify the URL and permissions.`);
      }
      console.log('API access test passed');
    } catch (apiError) {
      console.error('API access test failed:', apiError);
      throw new Error(`API access failed: ${apiError.message}. Please check your Search Console permissions.`);
    }
    
    // Create a mock event object for the runBackup function
    const mockEvent = {
      triggerUid: triggerUid
    };
    
    // Check execution time before proceeding
    const currentTime = new Date();
    if (currentTime - startTime > maxExecutionTime) {
      throw new Error('Execution timeout reached before starting backup');
    }
    
    console.log('Starting backup execution...');
    
    // Call the existing runBackup function with timeout protection
    const result = runBackup(mockEvent);
    
    console.log('Manual backup completed:', result);
    return result;
    
  } catch (error) {
    console.error('Error running manual backup:', error);
    
    // Provide more specific error messages
    let errorMessage = error.message || error.toString();
    
    if (errorMessage.includes('server error') || errorMessage.includes('timeout')) {
      errorMessage = 'Server timeout or temporary error. Please try again in a few minutes. If the problem persists, check your internet connection and try again.';
    } else if (errorMessage.includes('quota') || errorMessage.includes('limit')) {
      errorMessage = 'Google Apps Script quota exceeded. Please wait a while before trying again.';
    } else if (errorMessage.includes('permission') || errorMessage.includes('access')) {
      errorMessage = 'Permission denied. Please ensure you have access to the website in Search Console and the necessary script permissions.';
    } else if (errorMessage.includes('API')) {
      errorMessage = `API Error: ${errorMessage}. Please check your Search Console setup and try again.`;
    }
    
    throw new Error('Failed to run backup: ' + errorMessage);
  }
}

/**
 * Pauses a backup by disabling its trigger
 */
function pauseBackup(triggerUid) {
  try {
    console.log('Pausing backup for trigger UID:', triggerUid);
    
    // Get and update backup configuration
    const storedData = PropertiesService.getScriptProperties().getProperty('backup_' + triggerUid);
    if (!storedData) {
      throw new Error('No backup configuration found');
    }
    
    const triggerData = JSON.parse(storedData);
    
    if (triggerData.status === 'paused') {
      throw new Error('Backup is already paused');
    }
    
    // Update status to paused
    triggerData.status = 'paused';
    triggerData.pausedAt = new Date().toISOString();
    
    // Store updated configuration
    PropertiesService.getScriptProperties().setProperty('backup_' + triggerUid, JSON.stringify(triggerData));
    
    // Find and delete the actual trigger
    const triggers = ScriptApp.getProjectTriggers();
    const trigger = triggers.find(t => t.getUniqueId() === triggerUid);
    
    if (trigger) {
      ScriptApp.deleteTrigger(trigger);
      console.log('Trigger deleted successfully');
    } else {
      console.log('Trigger not found, but status updated');
    }
    
    return `Backup for ${triggerData.website} has been paused successfully`;
    
  } catch (error) {
    console.error('Error pausing backup:', error);
    throw new Error('Failed to pause backup: ' + error.message);
  }
}

/**
 * Resumes a paused backup by recreating its trigger
 */
function resumeBackup(triggerUid) {
  try {
    console.log('Resuming backup for trigger UID:', triggerUid);
    
    // Get backup configuration
    const storedData = PropertiesService.getScriptProperties().getProperty('backup_' + triggerUid);
    if (!storedData) {
      throw new Error('No backup configuration found');
    }
    
    const triggerData = JSON.parse(storedData);
    
    if (triggerData.status !== 'paused') {
      throw new Error('Backup is not paused');
    }
    
    // Delete the existing trigger entry (in case it still exists)
    const existingTriggers = ScriptApp.getProjectTriggers();
    const existingTrigger = existingTriggers.find(t => t.getUniqueId() === triggerUid);
    if (existingTrigger) {
      ScriptApp.deleteTrigger(existingTrigger);
    }
    
    // Create new trigger
    let newTrigger;
    if (triggerData.backupType === 'daily') {
      newTrigger = ScriptApp.newTrigger('runBackup')
        .timeBased()
        .everyDays(1)
        .atHour(2)
        .create();
    } else {
      newTrigger = ScriptApp.newTrigger('runBackup')
        .timeBased()
        .onMonthDay(3)
        .atHour(2)
        .create();
    }
    
    // Get new trigger UID
    const newTriggerUid = newTrigger.getUniqueId();
    
    // Update configuration with new trigger UID and active status
    triggerData.status = 'active';
    triggerData.resumedAt = new Date().toISOString();
    delete triggerData.pausedAt;
    
    // Delete old configuration and store with new trigger UID
    PropertiesService.getScriptProperties().deleteProperty('backup_' + triggerUid);
    PropertiesService.getScriptProperties().setProperty('backup_' + newTriggerUid, JSON.stringify(triggerData));
    
    console.log('Backup resumed with new trigger UID:', newTriggerUid);
    
    return `Backup for ${triggerData.website} has been resumed successfully`;
    
  } catch (error) {
    console.error('Error resuming backup:', error);
    throw new Error('Failed to resume backup: ' + error.message);
  }
}

/**
 * Deletes a backup configuration and its trigger
 */
function deleteBackup(triggerUid) {
  try {
    console.log('Deleting backup for trigger UID:', triggerUid);
    
    // Get backup configuration to return website info
    const storedData = PropertiesService.getScriptProperties().getProperty('backup_' + triggerUid);
    let websiteName = 'Unknown';
    
    if (storedData) {
      try {
        const triggerData = JSON.parse(storedData);
        websiteName = triggerData.website;
      } catch (parseError) {
        console.error('Error parsing trigger data during deletion:', parseError);
      }
    }
    
    // Find and delete the trigger
    const triggers = ScriptApp.getProjectTriggers();
    const trigger = triggers.find(t => t.getUniqueId() === triggerUid);
    
    if (trigger) {
      ScriptApp.deleteTrigger(trigger);
      console.log('Trigger deleted successfully');
    } else {
      console.log('Trigger not found, continuing with configuration cleanup');
    }
    
    // Delete the stored configuration
    PropertiesService.getScriptProperties().deleteProperty('backup_' + triggerUid);
    
    console.log('Backup configuration deleted');
    
    return `Backup for ${websiteName} has been deleted successfully`;
    
  } catch (error) {
    console.error('Error deleting backup:', error);
    throw new Error('Failed to delete backup: ' + error.message);
  }
}

/**
 * Test function to diagnose backup issues
 */
function testBackupDiagnostics() {
  const results = {
    timestamp: new Date().toISOString(),
    tests: {},
    recommendations: []
  };
  
  try {
    // Test 1: Check if we can access PropertiesService
    console.log('Test 1: PropertiesService access');
    const testProperty = 'test_' + Date.now();
    PropertiesService.getScriptProperties().setProperty(testProperty, 'test');
    const retrieved = PropertiesService.getScriptProperties().getProperty(testProperty);
    PropertiesService.getScriptProperties().deleteProperty(testProperty);
    results.tests.propertiesService = retrieved === 'test' ? 'PASS' : 'FAIL';
    console.log('PropertiesService test:', results.tests.propertiesService);
  } catch (error) {
    results.tests.propertiesService = 'FAIL: ' + error.message;
    console.error('PropertiesService test failed:', error);
  }
  
  try {
    // Test 2: Check if we can access Search Console API
    console.log('Test 2: Search Console API access');
    const websites = getWebsites();
    results.tests.searchConsoleAPI = websites && websites.length > 0 ? 'PASS' : 'FAIL: No websites found';
    console.log('Search Console API test:', results.tests.searchConsoleAPI);
  } catch (error) {
    results.tests.searchConsoleAPI = 'FAIL: ' + error.message;
    console.error('Search Console API test failed:', error);
  }
  
  try {
    // Test 3: Check if we can create a spreadsheet
    console.log('Test 3: Spreadsheet creation');
    const testSS = SpreadsheetApp.create('Test_Backup_Diagnostic_' + Date.now());
    const testSheet = testSS.getSheets()[0];
    testSheet.getRange(1, 1).setValue('Test');
    const testValue = testSheet.getRange(1, 1).getValue();
    DriveApp.getFileById(testSS.getId()).setTrashed(true);
    results.tests.spreadsheetCreation = testValue === 'Test' ? 'PASS' : 'FAIL';
    console.log('Spreadsheet creation test:', results.tests.spreadsheetCreation);
  } catch (error) {
    results.tests.spreadsheetCreation = 'FAIL: ' + error.message;
    console.error('Spreadsheet creation test failed:', error);
    if (error.message.includes('permissions') || error.message.includes('auth')) {
      results.recommendations.push('Run "Authorize All Permissions" from the menu to fix spreadsheet creation issues.');
    }
  }
  
  try {
    // Test 4: Check if we can access Drive
    console.log('Test 4: Drive access');
    const rootFolder = DriveApp.getRootFolder();
    const folders = rootFolder.getFolders();
    results.tests.driveAccess = folders ? 'PASS' : 'FAIL';
    console.log('Drive access test:', results.tests.driveAccess);
  } catch (error) {
    results.tests.driveAccess = 'FAIL: ' + error.message;
    console.error('Drive access test failed:', error);
    if (error.message.includes('server error')) {
      results.recommendations.push('Drive access failed due to server error. This is usually temporary. Try again in a few minutes.');
    } else if (error.message.includes('permissions') || error.message.includes('auth')) {
      results.recommendations.push('Run "Authorize All Permissions" from the menu to fix Drive access issues.');
    }
  }
  
  try {
    // Test 5: Check if we can send emails
    console.log('Test 5: Email access');
    const userEmail = Session.getEffectiveUser().getEmail();
    results.tests.emailAccess = userEmail ? 'PASS' : 'FAIL: No user email found';
    console.log('Email access test:', results.tests.emailAccess);
  } catch (error) {
    results.tests.emailAccess = 'FAIL: ' + error.message;
    console.error('Email access test failed:', error);
  }
  
  try {
    // Test 6: Check existing backup configurations
    console.log('Test 6: Backup configurations');
    const backupList = getBackupList();
    results.tests.backupConfigurations = backupList && backupList.length > 0 ? 
      `PASS: ${backupList.length} backup(s) found` : 'PASS: No backups configured';
    console.log('Backup configurations test:', results.tests.backupConfigurations);
  } catch (error) {
    results.tests.backupConfigurations = 'FAIL: ' + error.message;
    console.error('Backup configurations test failed:', error);
  }
  
  // Add general recommendations based on test results
  const failedTests = Object.values(results.tests).filter(test => test.startsWith('FAIL'));
  if (failedTests.length > 0) {
    if (!results.recommendations.includes('Run "Authorize All Permissions"')) {
      results.recommendations.push('Some tests failed. Try running "Authorize All Permissions" from the menu.');
    }
  } else {
    results.recommendations.push('All tests passed! Your backup system should work correctly.');
  }
  
  console.log('Diagnostic results:', JSON.stringify(results, null, 2));
  return results;
}

/**
 * Utility function to clean up orphaned backup configurations
 */
function cleanupOrphanedBackups() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    const triggers = ScriptApp.getProjectTriggers();
    const activeTriggerUids = new Set(triggers.map(t => t.getUniqueId()));
    
    let cleanedCount = 0;
    
    // Find backup configurations
    Object.keys(allProperties).forEach(key => {
      if (key.startsWith('backup_')) {
        const triggerUid = key.replace('backup_', '');
        
        // If no corresponding trigger exists, remove the configuration
        if (!activeTriggerUids.has(triggerUid)) {
          properties.deleteProperty(key);
          cleanedCount++;
          console.log('Cleaned up orphaned backup configuration:', triggerUid);
        }
      }
    });
    
    console.log(`Cleaned up ${cleanedCount} orphaned backup configurations`);
    return `Cleaned up ${cleanedCount} orphaned backup configurations`;
    
  } catch (error) {
    console.error('Error cleaning up orphaned backups:', error);
    throw new Error('Failed to cleanup orphaned backups: ' + error.message);
  }
}

/**
 * Check email quota
 */
function checkEmailQuota() {
  try {
    const quota = MailApp.getRemainingDailyQuota();
    return {
      remaining: quota,
      message: `${quota} emails remaining today`
    };
  } catch (error) {
    return {
      remaining: 0,
      message: "Could not check email quota: " + error.message
    };
  }
} 