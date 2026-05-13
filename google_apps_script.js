/**
 * Google Apps Script for WAM Planning Dashboard
 * 
 * SETUP:
 * 1. Open your Google Sheet.
 * 2. Go to Extensions > Apps Script.
 * 3. Paste this code.
 * 4. Update the sheet names in the CONFIG section if needed.
 * 5. Click "Deploy" > "New Deployment".
 * 6. Select "Web App".
 * 7. Execute as: "Me".
 * 8. Who has access: "Anyone".
 * 9. Copy the Web App URL and paste it into index.html as WEB_APP_URL.
 */

const CONFIG = {
  SHEET_MAIN: "WAM PLANNING",
  SHEET_PROJECTS: "PROJECT LIST",
  SHEET_EMPLOYEES: "employee",
  SHEET_TASKS: "Task",
  SHEET_ZONE_SOURCE: "PROJECT SHEET 24", // This is usually the source for Zone Numbers
  SHEET_USER_ACCESS: "USER ACCESS"
};

/**
 * Access Control Helper
 */
function isAllowedUser(email, ss) {
  if (!email) return { allowed: false, role: "Guest" };
  const cache = CacheService.getScriptCache();
  const cacheKey = "user_access_" + email.toLowerCase().replace(/[^a-z0-9]/g, "_");
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_USER_ACCESS);
  if (!sheet) return { allowed: false, role: "Guest" };
  
  const finder = sheet.createTextFinder(email.trim()).matchCase(false);
  const cell = finder.findNext();
  
  let result = { allowed: false, role: "Guest" };
  if (cell) {
    const row = cell.getRow();
    const role = sheet.getRange(row, 2).getValue();
    result = { allowed: true, role: String(role || "Viewer").trim() };
  }
  
  try {
    // Cache for 5 minutes
    cache.put(cacheKey, JSON.stringify(result), 300); 
  } catch (e) {}
  
  return result;
}

/**
 * GET Handler
 */
function doGet(e) {
  try {
    if (!e || (!e.parameter && !e.parameters)) {
      return ContentService.createTextOutput("Error: No parameters provided").setMimeType(ContentService.MimeType.TEXT);
    }
    const action = e.parameter.action || (e.parameters && e.parameters.action && e.parameters.action[0]);
    const email = e.parameter.email || (e.parameters && e.parameters.email && e.parameters.email[0]);
    
    // Core data fetch
    if (action === 'getData') {
      return handleGetData(email);
    }
    
    // Delete action (GET fallback for simplicity in some triggers, though DELETE/POST preferred)
    if (action === 'deleteRow') {
      const access = isAllowedUser(email);
      if (!access.allowed || access.role.toUpperCase() === 'VIEWER') {
        return createJsonResponse({ status: "error", message: "Access Denied" });
      }
      return handleDelete(e.parameter);
    }
    
    // Access verification
    if (action === 'checkAccess') {
      return createJsonResponse(isAllowedUser(email));
    }
    
    return ContentService.createTextOutput("Action not found (doGet): " + (action || "undefined")).setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    console.error("GET Error: " + err.message);
    return ContentService.createTextOutput("Error: " + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * POST Handler
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const email = postData.email;
    
    if (action === 'saveRow') {
      const access = isAllowedUser(email);
      const isUpdate = postData.data.rowId && !String(postData.data.rowId).startsWith('TEMP-');
      
      // If updating, strictly require allowed status and non-viewer role
      if (isUpdate && (!access.allowed || access.role.toUpperCase() === 'VIEWER')) {
        return createJsonResponse({ status: "error", message: "Access Denied: You do not have permission to edit rows." });
      }
      
      return handleSave(postData.data);
    }
    
    return ContentService.createTextOutput("Action not found (doPost): " + (action || "undefined")).setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    console.error("POST Error: " + err.message);
    return ContentService.createTextOutput("Error: " + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Helper to create JSON response with CORS compatibility
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Data aggregation for frontend
 */
function handleGetData(email) {
  const t0 = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const access = isAllowedUser(email, ss);
  const t1 = new Date().getTime();
  
  const sheetsMap = {};
  ss.getSheets().forEach(s => sheetsMap[s.getName().toUpperCase().trim()] = s);

  const fetchFast = (name) => {
    const s = sheetsMap[name.toUpperCase().trim()];
    if (!s) {
      console.warn("Sheet not found: " + name);
      return { headers: [], values: [] };
    }
    const lr = s.getLastRow();
    const lc = s.getLastColumn();
    if (lr < 1 || lc < 1) return { headers: [], values: [] };
    const data = s.getRange(1, 1, lr, lc).getValues();
    return { headers: data[0].map(h => String(h || "").trim().toUpperCase()), values: data.slice(1) };
  };

  const getZones = (sheet) => {
    if (!sheet) return [];
    const lr = sheet.getLastRow();
    const lc = sheet.getLastColumn();
    if (lr <= 1 || lc < 1) return [];
    const data = sheet.getRange(1, 1, lr, lc).getValues();
    const headers = data[0].map(h => String(h || "").trim().toUpperCase());
    const idx = headers.indexOf("ZONE NUMBER");
    if (idx === -1) return [];
    return [...new Set(data.slice(1).map(r => String(r[idx] || "").trim()).filter(Boolean))].sort();
  };

  const result = {
    main: fetchFast(CONFIG.SHEET_MAIN),
    projects: fetchFast(CONFIG.SHEET_PROJECTS),
    employees: fetchFast(CONFIG.SHEET_EMPLOYEES),
    tasks: fetchFast(CONFIG.SHEET_TASKS),
    zones: getZones(sheetsMap[CONFIG.SHEET_ZONE_SOURCE.toUpperCase().trim()]),
    userAccess: access,
    isAllowedUser: access.allowed,
    metrics: {
      authTime: t1 - t0,
      totalServerTime: new Date().getTime() - t0
    }
  };
  
  return createJsonResponse(result);
}

/**
 * Unified save/update logic
 */
function handleSave(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_MAIN);
  if (!mainSheet) throw new Error("Main sheet not found: " + CONFIG.SHEET_MAIN);
  
  const rowId = data.rowId;
  const isUpdate = rowId && !String(rowId).startsWith('TEMP-');
  
  const headers = mainSheet.getRange(1, 1, 1, Math.max(1, mainSheet.getLastColumn())).getValues()[0];
  const values = headers.map(rh => {
    const h = String(rh || "").trim().toUpperCase();
    const matchKey = Object.keys(data).find(k => k.trim().toUpperCase() === h);
    return matchKey !== undefined ? data[matchKey] : "";
  });

  if (isUpdate) {
    mainSheet.getRange(parseInt(rowId), 1, 1, values.length).setValues([values]);
  } else {
    // Deduplication logic: Check for exact matches on all mapped values.
    const existingData = mainSheet.getDataRange().getValues();
    
    // Columns to ignore when checking for duplicates (automatic or row-specific fields)
    const ignoreHeaders = ['ROW ID', 'TIMESTAMP', 'CREATED AT', 'UPDATED AT', 'REVISION NO.'];
    const indicesToCompare = [];
    headers.forEach((h, i) => {
      const uh = String(h || "").trim().toUpperCase();
      if (!ignoreHeaders.includes(uh)) {
        indicesToCompare.push(i);
      }
    });

    let targetRow = -1;
    for (let i = 1; i < existingData.length; i++) {
        const row = existingData[i];
        let isMatch = true;
        
        for (const colIdx of indicesToCompare) {
            // Compare mapped value with existing sheet value
            const rowVal = String(row[colIdx] || "").trim().toLowerCase();
            const dataVal = String(values[colIdx] || "").trim().toLowerCase();
            if (rowVal !== dataVal) {
                isMatch = false;
                break;
            }
        }
        
        if (isMatch) {
            targetRow = i + 1;
            break;
        }
    }

    if (targetRow !== -1) {
      mainSheet.getRange(targetRow, 1, 1, values.length).setValues([values]);
    } else {
      const lastRow = mainSheet.getLastRow();
      mainSheet.getRange(lastRow + 1, 1, 1, values.length).setValues([values]);
    }
  }

  // CRITICAL: Synchronize NEW Project/Zone to PROJECT LIST sheet
  let syncResult = "Not triggered";
  if (data.updateProjectList && data.projectUpdateRow) {
    try {
      syncResult = syncToProjectList(ss, data.projectUpdateRow, data.isNewMainRow);
      // Catch validation errors from syncToProjectList
      if (syncResult && typeof syncResult === "string" && syncResult.startsWith("ERROR:")) {
        return createJsonResponse({ 
          status: "error", 
          message: syncResult.replace("ERROR:", "").trim(),
          syncStatus: syncResult 
        });
      }
    } catch (e) {
      syncResult = "Error: " + e.message;
      console.error("Project List Sync Error: " + e.message);
    }
  }

  return createJsonResponse({ 
    status: "Success", 
    message: "Row saved successfully", 
    syncStatus: syncResult 
  });
}

/**
 * Syncs project details to the "PROJECT LIST" sheet
 */
function syncToProjectList(ss, updateData, isNewMainRow) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_PROJECTS);
  if (!sheet) return "Sheet not found: " + CONFIG.SHEET_PROJECTS;
  
  const lastRow = sheet.getLastRow();
  
  // Use a smaller range to get headers if sheet is huge, but usually headers are in row 1
  const rawHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  const headers = rawHeaders.map(h => String(h || "").trim().toUpperCase());
  
  const projNameIdx = headers.indexOf('PROJECT NAME');
  const zoneNameIdx = headers.indexOf('ZONE NAME');
  const zoneNumIdx = headers.indexOf('ZONE NUMBER');
  
  const projectName = String(updateData['PROJECT NAME'] || '').trim().toUpperCase();
  const zoneName = String(updateData['ZONE NAME'] || '').trim().toUpperCase();
  const zoneNumber = String(updateData['ZONE NUMBER'] || '').trim().toUpperCase();
  const sqFt = String(updateData['SQ. FT.'] || '').trim().toUpperCase();

  if (!projectName) return "Missing project name";

  // Find existing row - Must match ALL key fields to be considered an exact match
  let foundRowIdx = -1;
  const ALLOWED_COLUMNS = ['PROJECTS UNIQUE ID', 'PROJECT NAME', 'ZONE NAME', 'ZONE NUMBER', 'SQ. FT.', 'DIR', 'TL', 'DPT'];
  const sqFtIdx = headers.indexOf('SQ. FT.');

  if (projNameIdx !== -1 && lastRow > 1) {
    const searchData = sheet.getRange(1, 1, lastRow, headers.length).getValues();
    for (let i = 1; i < searchData.length; i++) {
      const row = searchData[i];
      const rowProj = String(row[projNameIdx] || "").trim().toUpperCase();
      if (rowProj !== projectName) continue;

      const rowZoneNum = zoneNumIdx !== -1 ? String(row[zoneNumIdx] || "").trim().toUpperCase() : "";
      const rowZoneName = zoneNameIdx !== -1 ? String(row[zoneNameIdx] || "").trim().toUpperCase() : "";
      const rowSqFt = sqFtIdx !== -1 ? String(row[sqFtIdx] || "").trim().toUpperCase() : "";

      // Exact match check for identity AND SQ. FT.
      if (rowZoneNum === zoneNumber && rowZoneName === zoneName && rowSqFt === sqFt) {
        foundRowIdx = i + 1;
        break;
      }
    }
  }

  // If an existing record was found, determine if we need to update it
  const targetRowIdx = foundRowIdx !== -1 ? foundRowIdx : (sheet.getLastRow() + 1);
  const isAppending = targetRowIdx > sheet.getLastRow();
  
  // Fetch existing row data (or empty row if appending)
  let rowValues;
  if (isAppending) {
    rowValues = new Array(headers.length).fill("");
  } else {
    rowValues = sheet.getRange(targetRowIdx, 1, 1, headers.length).getValues()[0];
  }
  
  let hasActualChanges = false;
  const updatedRow = rowValues.map((val, colIdx) => {
    const h = headers[colIdx];
    
    // ONLY update if it's in the allowed set of columns requested by the user
    if (ALLOWED_COLUMNS.indexOf(h) !== -1) {
      const matchKey = Object.keys(updateData).find(k => k.trim().toUpperCase() === h);
      if (matchKey !== undefined) {
        const newVal = String(updateData[matchKey]).trim();
        const oldVal = String(val).trim();
        
        // We only consider it a change if the new value is not empty AND differs from old
        if (newVal !== "" && newVal !== oldVal) {
          hasActualChanges = true;
          return updateData[matchKey];
        }
      }
    }
    return val;
  });

  // If we found an existing row and nothing changed, skip the write entirely
  if (!isAppending && !hasActualChanges) {
    return "Existing exact match found, no changes to update.";
  }

  sheet.getRange(targetRowIdx, 1, 1, updatedRow.length).setValues([updatedRow]);
  
  return (targetRowIdx > lastRow ? "Appended new" : "Updated existing") + " record at row " + targetRowIdx;
}

/**
 * Handle deletion
 */
function handleDelete(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_MAIN);
  if (!sheet) return createJsonResponse({ status: "error", message: "Sheet not found" });
  
  const rowIdParam = parseInt(params.id);
  const lastRow = sheet.getLastRow();
  
  const checkFields = {
    'PROJECT NAME': params.p,
    'ZONE NUMBER': params.z,
    'TASK': params.t,
    'WORK STAGE / LIST OF DRAWING': params.d,
    'DOER': params.doer
  };

  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0]
    .map(h => String(h || "").trim().toUpperCase());
  
  function isMatch(rowData) {
    for (const [key, expected] of Object.entries(checkFields)) {
      const idx = headers.indexOf(key.toUpperCase());
      if (idx === -1) continue; // Skip if column doesn't exist
      
      const actualValue = String(rowData[idx] || '').trim().toLowerCase();
      const expectedValue = String(expected || '').trim().toLowerCase();
      
      if (actualValue !== expectedValue) {
        return false;
      }
    }
    return true;
  }

  // 1. Try direct match by rowId
  if (!isNaN(rowIdParam) && rowIdParam >= 2 && rowIdParam <= lastRow) {
    const data = sheet.getRange(rowIdParam, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (isMatch(data)) {
      sheet.deleteRow(rowIdParam);
      return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
    }
  }

  // 2. Search fallback
  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (isMatch(allData[i])) {
      sheet.deleteRow(i + 1);
      return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
    }
  }

  return ContentService.createTextOutput("Row not found").setMimeType(ContentService.MimeType.TEXT);
}
