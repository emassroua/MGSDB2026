// ============================================================
// MGS - Massroua General Services
// Google Apps Script - Backend API
// ============================================================

var SPREADSHEET_ID = '1l0Yv80mYhUnzgkFRL26PSbVbAOI5RBvh5Qb43vB4n_8';
var DRIVE_FOLDER_NAME = 'MGS_Documents';

// ============================================================
// ENTRY POINTS
// ============================================================

function sanitizeDates(obj) {
  if (obj === null || obj === undefined) return obj;
  if (obj instanceof Date) {
    try { return isNaN(obj.getTime()) ? '' : obj.toISOString(); } catch(e) { return ''; }
  }
  if (Array.isArray(obj)) return obj.map(sanitizeDates);
  if (typeof obj === 'object') {
    var result = {};
    for (var k in obj) {
      if (obj.hasOwnProperty(k)) result[k] = sanitizeDates(obj[k]);
    }
    return result;
  }
  return obj;
}

function doGet(e) {
  try {
    var params = e.parameter || {};

    if (!params.action) {
      var page = params.page || 'Reports';
      var validPages = {
        'Supplier_Manager_FINAL': 'Supplier_Manager_FINAL',
        'Worker_Approval': 'Worker_Approval',
        'Worker_Documents': 'Worker_Documents',
        'Reports_Manager': 'Reports_Manager',
        'Supplier_Portal': 'Supplier_Portal',
        'Processing_Costs': 'Processing_Costs',
        'Nationalities_Manager': 'Nationalities_Manager',
        'local_Agencies_Portal': 'local_Agencies_Portal',
        'Local_Agencies_Manager': 'Local_Agencies_Manager',
        'Rules_Review': 'Rules_Review',
        'Visas_Manager': 'Visas_Manager',
        'Workers_Manager': 'Workers_Manager',
        'Messaging_Manager': 'Messaging_Manager',
        'Message_Templates': 'Message_Templates',
        'Employers_Manager': 'Employers_Manager',
        'Insurance_Manager': 'Insurance_Manager',
        'Admin_Dashboard': 'Admin_Dashboard',
        'Employer_Documents': 'Employer_Documents',
        'Reports': 'Reports',
        'Database_Reset_Manager': 'Database_Reset_Manager',
        'Dashboard': 'Dashboard',
        'Processing_Manager': 'Processing_Manager',
        'Documents_Manager': 'Documents_Manager',
        'Accountant': 'Accountant',
        'Cleanup_Unlinked_Employers': 'Cleanup_Unlinked_Employers'
      };
      var fileName = validPages[page] || 'Reports';
      var titles = {
        'Supplier_Manager_FINAL': 'Supplier Manager - MGSDB2026',
        'Worker_Approval': 'Worker Approval Control - MGSDB2026',
        'Rules_Review': 'Rules Review - MGSDB2026',
        'Reports_Manager': 'Reports Manager - MGSDB2026',
        'Visas_Manager': 'Visas Manager - MGSDB2026',
        'Processing_Costs': 'Processing Costs - MGSDB2026',
        'Worker_Documents': 'Worker Documents - MGSDB2026',
        'Nationalities_Manager': 'Nationalities Manager - MGSDB2026',
        'local_Agencies_Portal': 'Local Agencies Portal - MGSDB2026',
        'Local_Agencies_Manager': 'Local Agencies Manager - MGSDB2026',
        'Supplier_Portal': 'Supplier Portal - MGSDB2026',
        'Admin_Dashboard': 'Admin Dashboard - MGSDB2026',
        'Insurance_Manager': 'Insurance Manager - MGSDB2026',
        'Workers_Manager': 'Workers Manager - MGSDB2026',
        'Messaging_Manager': 'Messaging Manager - MGSDB2026',
        'Message_Templates': 'Message Templates - MGSDB2026',
        'Employer_Documents': 'Employer Documents - MGSDB2026',
        'Employers_Manager': 'Employers Manager - MGSDB2026',
        'Reports': 'Reports - MGSDB2026',
        'Database_Reset_Manager': 'Database Reset Manager - MGSDB2026',
        'Dashboard': 'Dashboard - MGSDB2026',
        'Processing_Manager': 'Processing Manager - MGSDB2026',
        'Accountant': 'Accountant - MGSDB2026',
        'Documents_Manager': 'Documents Manager - MGSDB2026',
        'Cleanup_Unlinked_Employers': 'Cleanup Unlinked Employers - MGSDB2026'
      };
      var html = HtmlService.createTemplateFromFile(fileName).evaluate().getContent();
      if (page !== 'Message_Templates' && page !== 'Messaging_Manager') {
        var sharedCss = HtmlService.createHtmlOutputFromFile('Shared_CSS').getContent();
        html = html.replace('</head>', sharedCss + '\n</head>');
      }
      return HtmlService.createHtmlOutput(html)
        .setTitle(titles[fileName] || 'MGSDB2026')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    var result = handleRequest(params);
    return buildResponse(sanitizeDates(result));
  } catch (err) {
    return buildResponse({ success: false, error: err.message });
  }
}

try { body = JSON.parse(e.postData.contents); } catch (parseErr) {
  // Twilio sends form-encoded — parse it properly
  try {
    body = {};
    var pairs = e.postData.contents.split('&');
    pairs.forEach(function(pair) {
      var kv = pair.split('=');
      if (kv[0]) body[decodeURIComponent(kv[0])] = decodeURIComponent((kv[1]||'').replace(/\+/g,' '));
    });
    if (Object.keys(body).length === 0) body = e.parameter || {};
  } catch(e2) { body = {}; }
}

function buildResponse(data) {
  try {
    var jsonStr = JSON.stringify(data, function(key, value) {
      if (value instanceof Date) {
        try {
          return isNaN(value.getTime()) ? '' : value.toISOString();
        } catch(e) {
          return '';
        }
      }
      if (typeof value === 'number' && key.toLowerCase().indexOf('date') >= 0) {
        return value;
      }
      return value;
    });
    return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    var safe = JSON.stringify(data, function(key, value) {
      try {
        JSON.stringify(value);
        return value;
      } catch(err) {
        return '';
      }
    });
    return ContentService.createTextOutput(safe).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// REQUEST ROUTER  — FIX: removed all duplicate cases + stray token
// ============================================================

function handleRequest(params) {
  var action = params.action || '';
  switch (action) {
    case 'statusCallback': return handleStatusCallback(params);
    case 'statusCallback':   return handleStatusCallback(e.parameter || params);
    case 'saveQuickReplies': return saveQuickReplies(params.data);
    case 'getQuickReplies':  return getQuickReplies();
    case 'saveQuickReplies': return saveQuickReplies(params.data);
    case 'getQuickReplies':  return getQuickReplies();
    case 'getAll':                   return actionGetAll(params);
    case 'saveRow':                  return actionSaveRow(params);
    case 'saveAll':                  return actionSaveAll(params);
    case 'deleteRow':                return actionDeleteRow(params);
    case 'getNextID':                return actionGetNextID(params);
    case 'uploadFile':               return actionUploadFile(params);
    case 'uploadDocsChunk':          return actionUploadFile(params);
    case 'deleteOldReports':         return actionDeleteOldReports(params);
    case 'createFolder':             return actionCreateFolder(params);
    case 'listFiles':                return actionListFiles(params);
    case 'findFile':                 return actionFindFile(params);
    case 'getFileContent':           return actionGetFileContent(params);
    case 'getPDFThumbnail':          return actionGetPDFThumbnail(params);
    case 'setup':                    return actionSetup(params);
    case 'syncContacts':             return actionSyncContacts(params);
    case 'getSyncProgress':          return actionGetSyncProgress(params);
    case 'uploadEmployerFile':       return uploadEmployerFile(params);
    case 'cleanupEmployerFiles':     return cleanupEmployerFiles(params);
    case 'saveMessagingAssignments': return saveMessagingAssignments(params.data);
    case 'getMessagingAssignments':  return getMessagingAssignments();
    case 'getMessages':              return getMessages();
    case 'scheduleMessage':          return scheduleMessage(params);
    case 'sendNow':                  return sendNow(params);
    case 'deleteMessage':            return deleteMessage(params);
    case 'processPendingMessages':   return processPendingMessages();
    case 'getTemplates':             return getTemplates();
    case 'saveTemplate':             return saveTemplate(params);
    case 'toggleTemplate':           return toggleTemplate(params);
    case 'deleteTemplate':           return deleteTemplate(params);
    case 'getSystemStats':           return getSystemStats();
    default: return { success: false, error: 'Unknown action: ' + action };
  }
}

// ============================================================
// ACTION: getAll
// ============================================================

function actionGetAll(params) {
  var sheetName = params.sheet;
  if (!sheetName) return { success: false, error: 'Missing sheet name' };
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Sheet not found: ' + sheetName, data: [] };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { success: true, data: [] };
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var data = rows.map(function(row) {
    var obj = {};
    headers.forEach(function(header, i) {
      var key = String(header).trim();
      if (!key) return;
      var val = row[i];
      if (val instanceof Date) {
        try {
          val = isNaN(val.getTime()) ? '' : Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } catch(e) {
          val = '';
        }
      } else if (val === null || val === undefined) {
        val = '';
      }
      obj[key] = val;
    });
    return obj;
  });
  data = data.filter(function(obj) {
    return Object.values(obj).some(function(v) { return v !== '' && v !== null && v !== undefined; });
  });
  return { success: true, data: data, count: data.length };
}

// ============================================================
// ACTION: saveRow
// ============================================================

function actionSaveRow(params) {
  var sheetName = params.sheet;
  var rowData   = params.row || params.data;
  if (!sheetName) return { success: false, error: 'Missing sheet name' };
  if (!rowData)   return { success: false, error: 'Missing row data' };
  if (typeof rowData === 'string') {
    try { rowData = JSON.parse(rowData); } catch(e) { return { success: false, error: 'Invalid row data JSON' }; }
  }
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var headers = [];
  if (lastCol > 0 && lastRow > 0) {
    headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  }
  var keys = Object.keys(rowData);
  keys.forEach(function(key) {
    if (key && headers.indexOf(key) === -1) {
      headers.push(key);
      sheet.getRange(1, headers.length).setValue(key);
    }
  });
  var idField = getIdField(sheetName);
  var rowId   = rowData[idField] !== undefined ? String(rowData[idField]) : null;
  var values  = headers.map(function(h) {
    if (h === '') return '';
    var val = rowData[h];
    if (val === undefined || val === null) return '';
    if (typeof val === 'object') return JSON.stringify(val);
    return val;
  });
  var targetRow = -1;
  if (rowId && idField && lastRow > 1) {
    var idColIndex = headers.indexOf(idField);
    if (idColIndex >= 0) {
      var idColValues = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < idColValues.length; i++) {
        if (String(idColValues[i][0]).trim() === rowId) { targetRow = i + 2; break; }
      }
    }
  }
  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, values.length).setValues([values]);
    return { success: true, action: 'updated', id: rowId, row: targetRow };
  } else {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, values.length).setValues([values]);
    return { success: true, action: 'inserted', id: rowId, row: newRow };
  }
}

// ============================================================
// ACTION: deleteRow
// ============================================================

function actionDeleteRow(params) {
  var sheetName = params.sheet;
  var id        = params.id !== undefined ? String(params.id) : null;
  if (!sheetName) return { success: false, error: 'Missing sheet name' };
  if (!id)        return { success: false, error: 'Missing id' };
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Sheet not found: ' + sheetName };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { success: false, error: 'Sheet is empty' };
  var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var idField  = getIdField(sheetName);
  var idColIdx = headers.indexOf(idField);
  if (idColIdx < 0) return { success: false, error: 'ID column not found: ' + idField };
  var idColValues = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < idColValues.length; i++) {
    if (String(idColValues[i][0]).trim() === id) {
      sheet.deleteRow(i + 2);
      return { success: true, action: 'deleted', id: id };
    }
  }
  return { success: false, error: 'Row not found with ID: ' + id };
}

// ============================================================
// ACTION: getNextID
// ============================================================

function actionGetNextID(params) {
  var sheetName = params.sheet;
  if (!sheetName) return { success: false, error: 'Missing sheet name' };
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: true, nextID: 1 };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, nextID: 1 };
  var lastCol  = sheet.getLastColumn();
  var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var idField  = getIdField(sheetName);
  var idColIdx = headers.indexOf(idField);
  if (idColIdx < 0) return { success: true, nextID: lastRow };
  var idValues = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
  var maxId = 0;
  idValues.forEach(function(row) {
    var val = parseInt(row[0]);
    if (!isNaN(val) && val > maxId) maxId = val;
  });
  return { success: true, nextID: maxId + 1 };
}

// ============================================================
// HELPER: getMgsRootFolder
// ============================================================

function getMgsRootFolder() {
  var props  = PropertiesService.getScriptProperties();
  var rootId = props.getProperty('MGS_ROOT_FOLDER_ID');
  if (rootId) {
    try {
      var folder = DriveApp.getFolderById(rootId);
      if (!folder.isTrashed()) return folder;
    } catch(e) {}
  }
  var oldIter = DriveApp.getFoldersByName('Documents');
  while (oldIter.hasNext()) {
    var oldDoc = oldIter.next();
    if (oldDoc.isTrashed()) continue;
    var wCheck = oldDoc.getFoldersByName('Workers');
    if (wCheck.hasNext()) {
      rootId = oldDoc.getId();
      props.setProperty('MGS_ROOT_FOLDER_ID', rootId);
      return oldDoc;
    }
  }
  var iter = DriveApp.getFoldersByName('MGS_Documents');
  while (iter.hasNext()) {
    var candidate = iter.next();
    if (!candidate.isTrashed()) {
      rootId = candidate.getId();
      props.setProperty('MGS_ROOT_FOLDER_ID', rootId);
      return candidate;
    }
  }
  var newRoot = DriveApp.createFolder('MGS_Documents');
  rootId = newRoot.getId();
  props.setProperty('MGS_ROOT_FOLDER_ID', rootId);
  return newRoot;
}

function getOrCreateFolder(name, parent) {
  if (!parent) {
    if (name === 'Documents' || name === 'MGS_Documents') return getMgsRootFolder();
    var folders = DriveApp.getFoldersByName(name);
    if (folders.hasNext()) return folders.next();
    return DriveApp.createFolder(name);
  }
  var childFolders = parent.getFoldersByName(name);
  if (childFolders.hasNext()) return childFolders.next();
  return parent.createFolder(name);
}

// ============================================================
// ACTION: uploadFile
// ============================================================

function actionUploadFile(params) {
  var folderName = params.folderName || 'Unknown';
  var fileName   = params.fileName   || 'file';
  var mimeType   = params.mimeType   || 'application/octet-stream';
  var parentType = params.parentType || params.parentFolder || 'employers';
  var fileData   = params.fileData   || params.content;
  if (!fileData) return { success: false, error: 'Missing file data' };
  var WORKER_PARENT   = 'Workers';
  var EMPLOYER_PARENT = 'Employers_Documents';
  try {
    var entityFolder = null;
    var fid = params.folderId ? String(params.folderId).trim() : '';
    if (fid !== '') {
      try {
        entityFolder = DriveApp.getFolderById(fid);
        if (entityFolder.isTrashed()) entityFolder = null;
      } catch (e) { entityFolder = null; }
    }
    if (!entityFolder) {
      var mgsRoot = getMgsRootFolder();
      if (parentType === 'workers') {
        var workersFolder = getOrCreateFolder(WORKER_PARENT, mgsRoot);
        var lock = LockService.getScriptLock();
        lock.waitLock(15000);
        try {
          var iter = workersFolder.getFoldersByName(folderName);
          if (iter.hasNext()) { entityFolder = iter.next(); }
          if (!entityFolder) {
            var uniqueKey = folderName.match(/^[A-Za-z0-9\-]+\s*\([^)]+\)/);
            if (uniqueKey) {
              var allSubs = workersFolder.getFolders();
              while (allSubs.hasNext()) {
                var c = allSubs.next();
                if (c.getName().indexOf(uniqueKey[0]) === 0 && !c.isTrashed()) { entityFolder = c; break; }
              }
            }
          }
          if (!entityFolder) entityFolder = workersFolder.createFolder(folderName);
        } finally { lock.releaseLock(); }
      } else {
        var empFolder = getOrCreateFolder(EMPLOYER_PARENT, mgsRoot);
        var lock2 = LockService.getScriptLock();
        lock2.waitLock(15000);
        try {
          var iter2 = empFolder.getFoldersByName(folderName);
          if (iter2.hasNext()) { entityFolder = iter2.next(); }
          if (!entityFolder) {
            var uniqueKey2 = folderName.match(/^[A-Za-z0-9\-]+\s*\([^)]+\)/);
            if (uniqueKey2) {
              var allSubs2 = empFolder.getFolders();
              while (allSubs2.hasNext()) {
                var cand2 = allSubs2.next();
                if (cand2.getName().indexOf(uniqueKey2[0]) === 0 && !cand2.isTrashed()) { entityFolder = cand2; break; }
              }
            }
          }
          if (!entityFolder) entityFolder = empFolder.createFolder(folderName);
        } finally { lock2.releaseLock(); }
      }
    }
    var isReport = fileName.indexOf('Report_') === 0;
    if (!isReport) {
      var existingFiles = entityFolder.getFilesByName(fileName);
      while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);
    }
    var blob    = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
    var file    = entityFolder.createFile(blob);
    var fileUrl = file.getUrl();
    var folderId = entityFolder.getId();
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    return { success: true, fileUrl: fileUrl, fileName: fileName, folderId: folderId, folderUrl: 'https://drive.google.com/drive/folders/' + folderId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// ACTION: deleteOldReports
// ============================================================

function actionDeleteOldReports(params) {
  var fid        = params.folderId   ? String(params.folderId).trim()   : '';
  var folderName = params.folderName ? String(params.folderName).trim() : '';
  var prefix     = params.prefix     || 'Report_';
  var deleted = [], errors = [];
  try {
    var folder = null;
    if (fid) { try { folder = DriveApp.getFolderById(fid); } catch(e) { folder = null; } }
    if (!folder && folderName) {
      var mgsRoot   = getMgsRootFolder();
      var empFolder = getOrCreateFolder('Employers_Documents', mgsRoot);
      var iter = empFolder.getFolders();
      while (iter.hasNext()) { var f = iter.next(); if (f.getName() === folderName) { folder = f; break; } }
    }
    if (!folder) return { success: false, error: 'Folder not found', deleted: [] };
    var prefixes = Array.isArray(prefix) ? prefix : [prefix];
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (prefixes.some(function(p) { return name.indexOf(p) === 0; })) {
        try { file.setTrashed(true); deleted.push(name); } catch(e) { errors.push(name + ': ' + e.message); }
      }
    }
    return { success: true, deleted: deleted, count: deleted.length, errors: errors };
  } catch(e) { return { success: false, error: e.message, deleted: [] }; }
}

// ============================================================
// ACTION: createFolder
// ============================================================

function actionCreateFolder(params) {
  var folderName = params.folderName;
  var parentType = params.parentType || 'employers';
  if (!folderName) return { success: false, error: 'Missing folderName' };
  try {
    var mgsRoot    = getMgsRootFolder();
    var parentName = parentType === 'workers' ? 'Workers' : 'Employers_Documents';
    var parent     = getOrCreateFolder(parentName, mgsRoot);
    var newFolder  = getOrCreateFolder(folderName, parent);
    return { success: true, folderName: folderName, folderId: newFolder.getId(), folderUrl: 'https://drive.google.com/drive/folders/' + newFolder.getId() };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// ACTION: saveAll
// ============================================================

function actionSaveAll(params) {
  var sheetName = params.sheet;
  var rows      = params.rows || params.data;
  if (!sheetName) return { success: false, error: 'Missing sheet name' };
  if (!rows)      return { success: false, error: 'Missing rows data' };
  if (typeof rows === 'string') {
    try { rows = JSON.parse(rows); } catch(e) { return { success: false, error: 'Invalid rows JSON' }; }
  }
  if (!Array.isArray(rows)) return { success: false, error: 'Invalid rows data' };
  if (rows.length === 0) {
    var ss2 = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh2 = ss2.getSheetByName(sheetName);
    if (sh2 && sh2.getLastRow() > 1) sh2.deleteRows(2, sh2.getLastRow() - 1);
    return { success: true, count: 0 };
  }
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  var headerSet = [];
  rows.forEach(function(row) {
    Object.keys(row).forEach(function(k) { if (k && headerSet.indexOf(k) === -1) headerSet.push(k); });
  });
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headerSet.length).setValues([headerSet]);
  var dataRows = rows.map(function(row) {
    return headerSet.map(function(h) {
      var val = row[h];
      if (val === undefined || val === null) return '';
      if (typeof val === 'object') return JSON.stringify(val);
      return val;
    });
  });
  sheet.getRange(2, 1, dataRows.length, headerSet.length).setValues(dataRows);
  return { success: true, count: rows.length, sheet: sheetName };
}

// ============================================================
// ACTION: setup
// ============================================================

function actionSetup(params) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = {
    'Local_Agencies': ['AgencyID','AgencyCode','AgencyName','LicenseNumber','OwnerName','AgencyPassword','ContactNumber','EmailAddress','Address','City','Country','Notes'],
    'Employers': ['EmployerID','AgencyID','AgencyCode','EmployerCode','EmployerName','ContactNumber','EmailAddress','Address','City','Country','Notes','FolderUrl','EmployerFolderID','EmployerFolderName','EmployerFolderUrl'],
    'Workers': ['WorkerID','AgencyID','AgencyCode','WorkerFirstName','WorkerMiddleName','WorkerLastName','WFullName','Nationality','Gender','DateOfBirth','ContactNumber','PassportNumber','PassportIssueDate','PassportExpiryDate'],
    'Suppliers': ['SupplierID','SupplierCODE','SupplierBusinessName','SupplierOwnerName','ContactNumber1','EmailAddress1'],
    'Worker_Documents': ['DocumentID','DocumentNameEN','DocumentNameAR','Required'],
    'Nationalities': ['NationalityID','NationalityName','CountryCode'],
    'Settings': ['Key','Value']
  };
  var created = [], existing = [];
  Object.keys(sheets).forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      var headers = sheets[name];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      created.push(name);
    } else {
      existing.push(name);
    }
  });
  return { success: true, created: created, existing: existing };
}

// ============================================================
// ACTION: listFiles
// ============================================================

function actionListFiles(params) {
  var folderId   = params.folderId   || '';
  var folderName = params.folderName || '';
  var parentType = params.parentType || 'employers';
  try {
    var folder = null;
    if (folderId && String(folderId).trim() !== '') {
      try {
        folder = DriveApp.getFolderById(String(folderId).trim());
        if (folder.isTrashed()) folder = null;
      } catch(e) { folder = null; }
    }
    if (!folder && folderName) {
      var mgsRoot      = getMgsRootFolder();
      var parentName   = parentType === 'workers' ? 'Workers' : 'Employers_Documents';
      var parentFolder = getOrCreateFolder(parentName, mgsRoot);
      var iter         = parentFolder.getFoldersByName(folderName);
      if (iter.hasNext()) { folder = iter.next(); }
      if (!folder) {
        var uniqueKey = folderName.match(/^[A-Za-z0-9\-]+\s*\([^)]+\)/);
        if (uniqueKey) {
          var allSubs = parentFolder.getFolders();
          while (allSubs.hasNext()) {
            var c = allSubs.next();
            if (c.getName().indexOf(uniqueKey[0]) === 0 && !c.isTrashed()) { folder = c; break; }
          }
        }
      }
    }
    if (!folder) return { success: false, error: 'Folder not found', files: [] };
    var files = [];
    function collectFiles(f, subfolderName) {
      var iter2 = f.getFiles();
      while (iter2.hasNext()) {
        var file = iter2.next();
        files.push({ id: file.getId(), name: file.getName(), url: file.getUrl(), mimeType: file.getMimeType(), size: file.getSize(), date: file.getDateCreated().toISOString(), subfolder: subfolderName || '' });
      }
      var subIter = f.getFolders();
      while (subIter.hasNext()) { var sub = subIter.next(); collectFiles(sub, (subfolderName ? subfolderName + '/' : '') + sub.getName()); }
    }
    collectFiles(folder, '');
    return { success: true, files: files, folderName: folder.getName(), folderId: folder.getId() };
  } catch(e) { return { success: false, error: e.message, files: [] }; }
}

// ============================================================
// ACTION: findFile
// ============================================================

function actionFindFile(params) {
  var folderName = params.folderName || '';
  var fileName   = params.fileName   || '';
  var parentType = params.parentFolder || 'workers';
  if (!fileName && !folderName) return { fileId: null, error: 'Missing fileName and folderName' };
  try {
    var mgsRoot      = getMgsRootFolder();
    var parentFolder = getOrCreateFolder(parentType === 'workers' ? 'Workers' : 'Employers_Documents', mgsRoot);
    if (folderName) {
      var parts = folderName.split('/').filter(function(p) { return p.length > 0; });
      var targetFolder = parentFolder;
      var folderFound  = true;
      for (var i = 0; i < parts.length; i++) {
        var subs = targetFolder.getFoldersByName(parts[i]);
        if (subs.hasNext()) { targetFolder = subs.next(); } else { folderFound = false; break; }
      }
      if (folderFound && fileName) {
        var files = targetFolder.getFilesByName(fileName);
        if (files.hasNext()) { var file = files.next(); return { fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName() }; }
      }
    }
    if (fileName) {
      var subFolders = parentFolder.getFolders();
      while (subFolders.hasNext()) {
        var sub = subFolders.next();
        var filesInSub = sub.getFilesByName(fileName);
        if (filesInSub.hasNext()) { var found = filesInSub.next(); return { fileId: found.getId(), fileUrl: found.getUrl(), fileName: found.getName() }; }
      }
    }
    return { fileId: null, error: 'File not found' };
  } catch(e) { return { fileId: null, error: e.message }; }
}

// ============================================================
// ACTION: getFileContent
// ============================================================

function actionGetFileContent(params) {
  var fileId = params.fileId || '';
  if (!fileId) return { success: false, error: 'Missing fileId' };
  try {
    var file   = DriveApp.getFileById(fileId);
    var blob   = file.getBlob();
    return { success: true, content: Utilities.base64Encode(blob.getBytes()), mimeType: blob.getContentType() || file.getMimeType(), name: file.getName(), size: file.getSize() };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// ACTION: getPDFThumbnail
// ============================================================

function actionGetPDFThumbnail(params) {
  var fileId = params.fileId || '';
  if (!fileId) return { success: false, error: 'Missing fileId' };
  var token = ScriptApp.getOAuthToken();
  try {
    var file = DriveApp.getFileById(fileId);
    var name = file.getName();
    var url1 = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1600-h2200';
    var r1   = UrlFetchApp.fetch(url1, { muteHttpExceptions: true, headers: { 'Authorization': 'Bearer ' + token } });
    var ct1  = r1.getBlob().getContentType() || '';
    if (r1.getResponseCode() === 200 && ct1.indexOf('image/') === 0) {
      return { success: true, content: Utilities.base64Encode(r1.getBlob().getBytes()), mimeType: ct1, name: name };
    }
    return { success: false, error: 'Thumbnail not available', name: name };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// HELPERS
// ============================================================

function getIdField(sheetName) {
  var map = {
    'Local_Agencies': 'AgencyID', 'Employers': 'EmployerID', 'Workers': 'WorkerID',
    'Suppliers': 'SupplierID', 'Worker_Documents': 'DocumentID', 'Employer_Documents': 'DocumentID',
    'Nationalities': 'NationalityID', 'Invoices': 'invoiceNumber', 'Receipts': 'receiptNumber',
    'OutgoingPayments': 'voucherNumber', 'Settings': 'Key'
  };
  return map[sheetName] || 'ID';
}

function getWebAppUrl() { return ScriptApp.getService().getUrl(); }

function clearDashboardCache() { CacheService.getScriptCache().remove('dashboard_data'); }

// ============================================================
// HtmlService bridge functions
// ============================================================

function getEmployers() { return actionGetAll({ sheet: 'Employers' }); }
function getWorkers()   { return actionGetAll({ sheet: 'Workers' }); }
function getAgencies()  { return actionGetAll({ sheet: 'Local_Agencies' }); }

function listFilesForEmployer(folderId, folderName) {
  return actionListFiles({ folderId: folderId, folderName: folderName, parentType: 'employers' });
}

function uploadReport(folderId, folderName, fileName, fileData) {
  return actionUploadFile({ folderId: folderId, folderName: folderName, fileName: fileName, fileData: fileData, mimeType: 'application/pdf', parentType: 'employers' });
}

function deleteOldReportsForEmployer(folderId, folderName) {
  return actionDeleteOldReports({ folderId: folderId, folderName: folderName, prefix: ['Report_', 'Docs_'] });
}

// ============================================================
// gsr bridge functions
// ============================================================

function gsrGetAll(sheet) {
  var result = actionGetAll({ sheet: sheet });
  if (result && result.data && sheet === 'Employers') {
    result.data = result.data.map(function(row) {
      return {
        EmployerID: row.EmployerID,
        EmployerCode: row.EmployerCode,
        EFN: row.EFN,
        MobileNumber: row.MobileNumber,
        City: row.City,
        AgencyCode: row.AgencyCode,
        AgencyID: row.AgencyID,
        SyncStatus: row.SyncStatus,
        LastSyncDate: row.LastSyncDate,
        EmployerFolderName: row.EmployerFolderName,
        EmployerFolderID: row.EmployerFolderID,
        EmployerFolderUrl: row.EmployerFolderUrl,
        Disabled: row.Disabled,
        DisabledAt: row.DisabledAt,
        GoogleContactID: row.GoogleContactID
      };
    });
  }
  return result;
}
function gsrDeleteRow(sheetName, id) { return actionDeleteRow({ sheet: sheetName, id: id }); }
function gsrSaveRow(sheet, data)     { return actionSaveRow({ sheet: sheet, data: data }); }

function gsrClearSheet(sheetName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Sheet not found: ' + sheetName };
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return { success: true, cleared: sheetName };
}

function gsrGetPageUrl(page) {
    return ScriptApp.getService().getUrl() + '?page=' + page;
}

function gsrUploadFile(fileNameOrData, base64Data, mimeType, folderId) {
  if (typeof fileNameOrData === 'object' && fileNameOrData !== null) return actionUploadFile(fileNameOrData);
  return actionUploadFile({ fileName: fileNameOrData, fileData: base64Data, mimeType: mimeType || 'application/octet-stream', folderId: folderId || '', parentType: 'workers' });
}

function gsrPost(jsonString) {
  var data = {};
  try { data = typeof jsonString === 'string' ? JSON.parse(jsonString) : jsonString; } catch(e) { return { success: false, error: 'Invalid JSON: ' + e.message }; }

  // Handle deleteFolder here directly
  if (data.action === 'deleteFolder') {
    var fId = data.folderId || '';
    if (!fId && data.folderName) {
      try {
        var parentFolders = DriveApp.getFoldersByName('workers');
        if (parentFolders.hasNext()) {
          var subFolders = parentFolders.next().getFoldersByName(data.folderName);
          if (subFolders.hasNext()) fId = subFolders.next().getId();
        }
      } catch(e) { return { success: false, error: 'Folder search failed: ' + e.message }; }
    }
    if (fId) {
      try { DriveApp.getFolderById(fId).setTrashed(true); return { success: true, message: 'Folder deleted' }; }
      catch(e) { return { success: false, error: 'Could not delete folder: ' + e.message }; }
    }
    return { success: true, message: 'No folder found — nothing to delete' };
  }

  return handleRequest(data);
}

function gsrGetAllDashboard() {
  function countRows(sheet) { var all = actionGetAll({ sheet: sheet }); return (all && all.data) ? all.data.length : 0; }
  function slim(sheet, fields) {
    var all = actionGetAll({ sheet: sheet });
    if (!all || !all.data) return { data: [] };
    return { data: all.data.map(function(row) { var obj = {}; fields.forEach(function(f) { obj[f] = row[f]; }); return obj; }) };
  }
  return {
    workers:   slim('Workers',       ['WorkerID','Processing','ProcessingStage','WorkerOutcomes','EmployerCode','Locked','MovedEmployerHistory']),
    employers: slim('Employers',     ['EmployerID','EmployerCode','EmployerName','EFN','AgencyCode']),
    agencies:  slim('Local_Agencies',['AgencyCode','AgencyName']),
    suppliers: { data: new Array(countRows('Suppliers')) },
    nationalities: slim('Nationalities', ['Active'])
  };
}

function getAllFileContents(fileIds) {
  if (!fileIds || !fileIds.length) return { success: false, error: 'No fileIds provided' };
  var results = {}, errors = [];
  for (var i = 0; i < fileIds.length; i++) {
    var fid = fileIds[i];
    if (!fid) continue;
    try {
      var file = DriveApp.getFileById(fid);
      var blob = file.getBlob();
      results[fid] = { content: Utilities.base64Encode(blob.getBytes()), mimeType: blob.getContentType() || file.getMimeType() };
    } catch(e) { errors.push(fid + ':' + e.message); results[fid] = null; }
  }
  return { success: true, files: results, errors: errors };
}

// ============================================================
// uploadEmployerFile
// ============================================================

function uploadEmployerFile(params) {
  try {
    var employerId = String(params.employerId || '');
    var documentId = String(params.documentId || '');
    var fileName   = params.fileName   || 'file';
    var mimeType   = params.mimeType   || 'application/octet-stream';
    var fileData   = params.fileData   || '';
    if (!employerId || !documentId || !fileData) return { success: false, message: 'Missing employerId, documentId or fileData' };
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName('Employers');
    if (!sheet) return { success: false, message: 'Employers sheet not found' };
    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];
    var colID       = headers.indexOf('EmployerID');
    var colFolderID = headers.indexOf('EmployerFolderID');
    var colUploaded = headers.indexOf('UploadedDocuments');
    if (colID === -1 || colFolderID === -1 || colUploaded === -1) return { success: false, message: 'Required column not found' };
    var empRow = -1, empRecord = null;
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][colID]) === employerId) { empRow = i + 1; empRecord = allData[i]; break; }
    }
    if (!empRecord) return { success: false, message: 'Employer ' + employerId + ' not found' };
    var folderId = String(empRecord[colFolderID] || '').trim();
    if (!folderId) return { success: false, message: 'EmployerFolderID is empty' };
    var employerFolder;
    try { employerFolder = DriveApp.getFolderById(folderId); } catch(e) { return { success: false, message: 'Cannot access Drive folder: ' + e.message }; }
    var now      = new Date();
    var tz       = Session.getScriptTimeZone();
    var datePart = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var timePart = Utilities.formatDate(now, tz, 'HH-mm-ss') + '-' + String(now.getMilliseconds()).padStart(3, '0');
    var docName  = String(params.documentName || documentId).trim().replace(/[\/\\:*?"<>|]/g, '').trim();
    var ext      = fileName.lastIndexOf('.') !== -1 ? fileName.split('.').pop().toLowerCase() : '';
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    var copyNumber, uploadName, fileId, fileUrl;
    try {
      var freshRow     = sheet.getRange(empRow, 1, 1, headers.length).getValues()[0];
      var existingRaw  = freshRow[colUploaded] || '';
      var uploadedDocs = {};
      try { uploadedDocs = JSON.parse(existingRaw); } catch(e) { uploadedDocs = {}; }
      if (typeof uploadedDocs !== 'object' || Array.isArray(uploadedDocs)) uploadedDocs = {};
      var existingArr  = uploadedDocs[documentId] || [];
      if (!Array.isArray(existingArr)) existingArr = [existingArr];
      copyNumber  = existingArr.length + 1;
      uploadName  = docName + '_' + copyNumber + '_' + datePart + '_' + timePart + (ext ? '.' + ext : '');
      var decoded  = Utilities.base64Decode(fileData);
      var blob     = Utilities.newBlob(decoded, mimeType, uploadName);
      var uploaded = employerFolder.createFile(blob);
      uploaded.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileId  = uploaded.getId();
      fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
      existingArr.push({ copyNumber: copyNumber, fileName: uploadName, originalName: fileName, fileId: fileId, fileUrl: fileUrl, uploadedAt: now.toISOString() });
      uploadedDocs[documentId] = existingArr;
      sheet.getRange(empRow, colUploaded + 1).setValue(JSON.stringify(uploadedDocs));
    } finally { lock.releaseLock(); }
    return { success: true, fileId: fileId, fileUrl: fileUrl, fileName: uploadName, copyNumber: copyNumber, folder: employerFolder.getName() };
  } catch(e) { return { success: false, message: 'Upload error: ' + e.message }; }
}

// ============================================================
// cleanupEmployerFiles
// ============================================================

function cleanupEmployerFiles(params) {
  try {
    var employerId = String(params.employerId || '');
    if (!employerId) return { success: false, message: 'Missing employerId' };
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName('Employers');
    if (!sheet) return { success: false, message: 'Employers sheet not found' };
    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];
    var colID       = headers.indexOf('EmployerID');
    var colFolderID = headers.indexOf('EmployerFolderID');
    var colUploaded = headers.indexOf('UploadedDocuments');
    var empRow = -1, empRecord = null;
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][colID]) === employerId) { empRow = i + 1; empRecord = allData[i]; break; }
    }
    if (!empRecord) return { success: false, message: 'Employer not found' };
    var deleted = 0;
    var folderId = String(empRecord[colFolderID] || '').trim();
    if (folderId) {
      try {
        var folder = DriveApp.getFolderById(folderId);
        var files  = folder.getFiles();
        while (files.hasNext()) { files.next().setTrashed(true); deleted++; }
      } catch(e) {}
    }
    if (colUploaded !== -1) sheet.getRange(empRow, colUploaded + 1).setValue('');
    return { success: true, deleted: deleted };
  } catch(e) { return { success: false, message: 'Cleanup error: ' + e.message }; }
}

// ============================================================
// TEMPLATE RULES
// ============================================================

function saveTemplateRules(rulesJson) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TemplateRules');
  if (!sheet) { sheet = ss.insertSheet('TemplateRules'); sheet.getRange('A1').setValue('RulesJSON'); }
  sheet.getRange('A2').setValue(rulesJson);
  return 'ok';
}

function getTemplateRules() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TemplateRules');
  if (!sheet) return '{}';
  var val = sheet.getRange('A2').getValue();
  return val || '{}';
}

// ============================================================
// sendNow — FIX: added missing var ss declaration
// ============================================================

function sendNow(params) {
  try {
    var phone    = (params.recipientPhone || '').toString().trim();
    var text     = (params.messageText   || '').toString().trim();
    var name     = (params.recipientName || '').toString().trim();
    var mediaURL = (params.mediaURL      || '').toString().trim();

    if (!phone) return { success: false, error: 'No phone number provided.' };
    if (!text && !mediaURL) return { success: false, error: 'No message text or media provided.' };

    var result = sendTwilioMessage_(phone, text, mediaURL);

    var ss    = SpreadsheetApp.getActiveSpreadsheet();  // FIX: was missing
    var log   = ss.getSheetByName('Messages');
    var now   = new Date();
    var msgID = 'MSG' + now.getTime();

    if (result.success) {
      if (log) {
        log.appendRow([
          msgID, now,
          params.recipientType || 'Employer',
          params.recipientID   || '',
          name, phone, text,
          mediaURL,
          params.mediaType     || '',
          '', '', 'Sent', now, '',
          params.sentBy        || 'Admin'
        ]);
      }
    } else {
      if (log) {
        log.appendRow([
          msgID, now,
          params.recipientType || 'Employer',
          params.recipientID   || '',
          name, phone, text,
          mediaURL,
          params.mediaType     || '',
          '', '', 'Failed', '', result.error || 'Unknown error',
          params.sentBy        || 'Admin'
        ]);
      }
    }

    return result;
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// Twilio
// ============================================================

function sendTwilioMessage_(to, body, mediaURL) {
  var TWILIO_SID   = 'AC47f2e9614496b77eaba98866a6f8ef02';
  var TWILIO_TOKEN = '1db1564cc75f9f74fc60677ee7031d1f';
  var FROM = 'whatsapp:+14155238886';

  var toPhone = to.toString()
    .replace(/\s/g, '')
    .replace(/-/g, '')
    .replace(/\(/g, '')
    .replace(/\)/g, '')
    .replace(/whatsapp:/gi, '');
  if (toPhone.startsWith('0')) toPhone = '+961' + toPhone.substring(1);
  else if (!toPhone.startsWith('+')) {
    if (toPhone.length <= 8) toPhone = '+961' + toPhone;
    else toPhone = '+' + toPhone;
  }
  toPhone = toPhone.replace(/^\+\+/, '+');
  Logger.log('📱 Sending to: whatsapp:' + toPhone);

  try {
    var payload = { From: FROM, To: 'whatsapp:' + toPhone, Body: body || '' };
    if (mediaURL && mediaURL.toString().trim() !== '') {
      payload.MediaUrl = mediaURL.toString().trim();
    }
    var resp = UrlFetchApp.fetch(
      'https://api.twilio.com/2010-04-01/Accounts/' + TWILIO_SID + '/Messages.json',
      { method: 'post', payload: payload, headers: { Authorization: 'Basic ' + Utilities.base64Encode(TWILIO_SID + ':' + TWILIO_TOKEN) }, muteHttpExceptions: true }
    );
    var json = JSON.parse(resp.getContentText());
    Logger.log('📨 Twilio response: ' + JSON.stringify(json));
    return json.sid ? { success: true, sid: json.sid } : { success: false, error: json.message || 'Unknown' };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// DELETE MESSAGE — NEW
// ============================================================

function deleteMessage(params) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Messages');
    if (!sheet) return { success: false, error: 'Messages sheet not found' };

    var msgID = (params.messageID || '').toString().trim();
    if (!msgID) return { success: false, error: 'Missing messageID' };

    var data     = sheet.getDataRange().getValues();
    var headers  = data[0];
    var colID    = headers.indexOf('MessageID');
    var colStatus= headers.indexOf('Status');
    if (colID === -1) return { success: false, error: 'MessageID column not found' };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colID]).trim() === msgID) {
        if (colStatus !== -1) {
          sheet.getRange(i + 1, colStatus + 1).setValue('Cancelled');
        } else {
          sheet.deleteRow(i + 1);
        }
        return { success: true, messageID: msgID };
      }
    }
    return { success: false, error: 'Message not found: ' + msgID };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// PROCESS PENDING MESSAGES — NEW
// ============================================================

function processPendingMessages() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Messages');
    if (!sheet) return { success: false, error: 'Messages sheet not found', sent: 0, failed: 0 };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var colStatus    = headers.indexOf('Status');
    var colPhone     = headers.indexOf('RecipientPhone');
    var colText      = headers.indexOf('MessageText');
    var colMediaURL  = headers.indexOf('MediaURL');
    var colSentDate  = headers.indexOf('SentDate');
    var colErrMsg    = headers.indexOf('ErrorMessage');
    var colSchedDate = headers.indexOf('ScheduledDate');

    if (colStatus === -1 || colPhone === -1)
      return { success: false, error: 'Required columns not found', sent: 0, failed: 0 };

    var today = new Date(); today.setHours(0,0,0,0);
    var sent  = 0, failed = 0;

    for (var i = 1; i < data.length; i++) {
      var status = (data[i][colStatus] || '').toString().trim();
      if (status !== 'Pending') continue;

      // Only send if scheduled date <= today
      if (colSchedDate !== -1 && data[i][colSchedDate]) {
        var sd = new Date(data[i][colSchedDate]); sd.setHours(0,0,0,0);
        if (sd > today) continue;
      }

      var phone    = (data[i][colPhone]   || '').toString().trim();
      var text     = (data[i][colText]    || '').toString().trim();
      var mediaURL = colMediaURL !== -1 ? (data[i][colMediaURL] || '').toString().trim() : '';

      if (!phone) { failed++; continue; }

      var result = sendTwilioMessage_(phone, text, mediaURL);

      if (result.success) {
        sent++;
        sheet.getRange(i + 1, colStatus  + 1).setValue('Sent');
        if (colSentDate !== -1) sheet.getRange(i + 1, colSentDate + 1).setValue(new Date());
        if (colErrMsg   !== -1) sheet.getRange(i + 1, colErrMsg   + 1).setValue('');
      } else {
        failed++;
        sheet.getRange(i + 1, colStatus + 1).setValue('Failed');
        if (colErrMsg !== -1) sheet.getRange(i + 1, colErrMsg + 1).setValue(result.error || 'Unknown');
      }
    }
    return { success: true, sent: sent, failed: failed };
  } catch(e) { return { success: false, error: e.message, sent: 0, failed: 0 }; }
}

// ============================================================
// SCHEDULE MESSAGE
// ============================================================

function scheduleMessage(params) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Messages');
    if (!sheet) return { success: false, error: 'Messages sheet not found' };

    var now   = new Date();
    var msgID = 'MSG' + now.getTime();

    sheet.appendRow([
      msgID,
      now,
      params.recipientType || 'Employer',
      params.recipientID   || '',
      params.recipientName || '',
      params.recipientPhone|| '',
      params.messageText   || '',
      params.mediaURL      || '',
      params.mediaType     || '',
      params.scheduledDate || '',
      params.scheduledTime || '09:00',
      'Pending',
      '',
      '',
      params.sentBy        || 'Admin'
    ]);

    return { success: true, messageID: msgID };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// GET MESSAGES
// ============================================================

function getMessages() {
  Logger.log('getMessages CALLED');
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Messages');
  Logger.log('Sheet rows: ' + (sheet ? sheet.getLastRow() : 'NOT FOUND'));
  var data = [];
  if (sheet && sheet.getLastRow() > 1) {
    var vals = sheet.getDataRange().getValues();
    var headers = vals[0];
    for (var i = 1; i < vals.length; i++) {
      var obj = {};
      headers.forEach(function(h, j) {
        var val = vals[i][j];
        if (val instanceof Date) {
          try { val = isNaN(val.getTime()) ? '' : val.toISOString(); } catch(e) { val = ''; }
        }
        obj[h] = val;
      });
      data.push(obj);
    }
  }
  data.reverse();
  Logger.log('Returning: ' + data.length + ' rows, first status: ' + (data[0] ? data[0].Status : 'none'));
  return { success: true, data: data };
}

// ============================================================
// MESSAGING ASSIGNMENTS
// ============================================================

function getMessagingAssignments() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('MessagingAssignments');
    if (!sheet) return { success: true, data: {} };
    var val = sheet.getRange('A2').getValue();
    return { success: true, data: val ? JSON.parse(val) : {} };
  } catch(e) { return { success: false, data: {} }; }
}

// FIX: was completely missing
function saveMessagingAssignments(data) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('MessagingAssignments');
    if (!sheet) {
      sheet = ss.insertSheet('MessagingAssignments');
      sheet.getRange('A1').setValue('AssignmentsJSON');
    }
    sheet.getRange('A2').setValue(JSON.stringify(data));
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function saveQuickReplies(data) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('MessagingAssignments');
    if (!sheet) {
      sheet = ss.insertSheet('MessagingAssignments');
      sheet.getRange('A1').setValue('AssignmentsJSON');
    }
    sheet.getRange('B1').setValue('QuickRepliesJSON');
    sheet.getRange('B2').setValue(JSON.stringify(data));
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function getQuickReplies() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('MessagingAssignments');
    if (!sheet) return { success: true, data: [] };
    var val = sheet.getRange('B2').getValue();
    return { success: true, data: val ? JSON.parse(val) : [] };
  } catch(e) { return { success: false, data: [] }; }
}

// ============================================================
// GET TEMPLATES PUBLIC
// ============================================================

function getTemplatesPublic() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('MessageTemplates') || ss.getSheetByName('Message_Templates');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows    = data.slice(1);
    var templates = rows.map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) {
        var val = row[i];
        if (val instanceof Date) {
          try { val = isNaN(val.getTime()) ? '' : val.toISOString(); } catch(e) { val = ''; }
        }
        obj[h] = val;
      });
      return obj;
    });
    return { success: true, data: templates };
  } catch(e) {
    return { success: false, error: e.message, data: [] };
  }
}

// ============================================================
// TEMPLATE CRUD — NEW
// ============================================================

function getTemplates() { return getTemplatesPublic(); }

function saveTemplate(params) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('MessageTemplates') || ss.getSheetByName('Message_Templates');
    if (!sheet) return { success: false, error: 'MessageTemplates sheet not found' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return (h||'').toString().trim(); });
    var tpl     = params.template || params;
    var tplID   = (tpl.TemplateID || '').toString().trim();
    var colID   = headers.indexOf('TemplateID');
    if (colID === -1) return { success: false, error: 'TemplateID column not found' };

    var values = headers.map(function(h) {
      var v = tpl[h]; return (v === undefined || v === null) ? '' : v;
    });

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colID]).trim() === tplID) {
        sheet.getRange(i + 1, 1, 1, values.length).setValues([values]);
        return { success: true, action: 'updated', templateID: tplID };
      }
    }
    sheet.appendRow(values);
    return { success: true, action: 'inserted', templateID: tplID };
  } catch(e) { return { success: false, error: e.message }; }
}

function toggleTemplate(params) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('MessageTemplates') || ss.getSheetByName('Message_Templates');
    if (!sheet) return { success: false, error: 'MessageTemplates sheet not found' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return (h||'').toString().trim(); });
    var colID     = headers.indexOf('TemplateID');
    var colActive = headers.indexOf('Active');
    if (colID === -1 || colActive === -1)
      return { success: false, error: 'TemplateID or Active column not found' };

    var tplID = (params.templateID || '').toString().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colID]).trim() === tplID) {
        var cur    = data[i][colActive];
        var newVal = !(cur === true || cur === 'true' || cur === 'TRUE');
        sheet.getRange(i + 1, colActive + 1).setValue(newVal);
        return { success: true, templateID: tplID, active: newVal };
      }
    }
    return { success: false, error: 'Template not found: ' + tplID };
  } catch(e) { return { success: false, error: e.message }; }
}

function deleteTemplate(params) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('MessageTemplates') || ss.getSheetByName('Message_Templates');
    if (!sheet) return { success: false, error: 'MessageTemplates sheet not found' };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return (h||'').toString().trim(); });
    var colID   = headers.indexOf('TemplateID');
    if (colID === -1) return { success: false, error: 'TemplateID column not found' };

    var tplID = (params.templateID || '').toString().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colID]).trim() === tplID) {
        sheet.deleteRow(i + 1);
        return { success: true, templateID: tplID };
      }
    }
    return { success: false, error: 'Template not found: ' + tplID };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// uploadMediaToDrive
// ============================================================

function uploadMediaToDrive(params) {
  try {
    var folder = DriveApp.getFoldersByName('MGS_Media').hasNext()
      ? DriveApp.getFoldersByName('MGS_Media').next()
      : DriveApp.createFolder('MGS_Media');
    var decoded  = Utilities.base64Decode(params.base64);
    var blob     = Utilities.newBlob(decoded, params.mimeType, params.fileName);
    var file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    var fileId    = file.getId();
    var publicUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
    return { success: true, publicUrl: publicUrl };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// syncSingleContact
// ============================================================

function syncSingleContact(params) {
  try {
    var phone = (params.phone || '').toString().trim();
    var name  = (params.name  || '').toString().trim().toLowerCase();
    var contacts = getAllGoogleContacts_();
    if (!contacts || contacts.length === 0) return { success: false, error: 'No contacts found in Google Contacts.' };
    var normalizePhone = function(p) {
      var n = (p || '').toString().replace(/[\s\-\(\)\+]/g, '');
      if (n.indexOf('961') === 0) n = n.substring(3);
      if (n.charAt(0) === '0') n = n.substring(1);
      return n;
    };
    var cleanPhone = normalizePhone(phone);
    var match = null;
    if (cleanPhone && cleanPhone.length >= 6) {
      match = contacts.find(function(c) {
        return (c.allPhones || []).some(function(p) { return normalizePhone(p) === cleanPhone; });
      });
    }
    var stripPrefix = function(n) { return (n || '').toLowerCase().replace(/^(ncs|mgs|mgsc|mr|mrs|ms|dr)\s+/i, '').trim(); };
    if (!match && name) {
      var cleanSearchName = stripPrefix(name);
      match = contacts.find(function(c) {
        var cName = stripPrefix(c.name);
        return cName.indexOf(cleanSearchName) !== -1 || cleanSearchName.indexOf(cName) !== -1;
      });
    }
    if (!match) return { success: false, error: 'No matching contact found for phone "' + phone + '" or name "' + params.name + '".' };
    var contactData = { EFN: match.name || '', City: match.city || '', Email: match.emails && match.emails[0] ? match.emails[0] : '', Address: match.address || '', Notes: match.notes || '', Organization: match.org || '', JobTitle: match.title || '', BirthDate: match.birthday || '', GoogleContactID: match.resourceName || '' };
    var allPhones = match.allPhones || [];
    allPhones.forEach(function(p, i) { var key = i === 0 ? 'MobileNumber' : 'MobileNumber' + (i + 1); contactData[key] = p; });
    return { success: true, contact: contactData, totalPhones: allPhones.length };
  } catch(e) { return { success: false, error: e.message }; }
}

function getAllGoogleContacts_() {
  var contacts = [], pageToken = null;
  do {
    var url = 'https://people.googleapis.com/v1/people/me/connections'
      + '?personFields=names,phoneNumbers,emailAddresses,addresses,biographies,organizations,birthdays'
      + '&pageSize=1000'
      + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    var response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    var data = JSON.parse(response.getContentText());
    if (!data.connections) break;
    data.connections.forEach(function(person) {
      var names  = person.names || [], phones = person.phoneNumbers || [], emails = person.emailAddresses || [];
      var addrs  = person.addresses || [], bios = person.biographies || [], orgs = person.organizations || [], bdays = person.birthdays || [];
      var primaryName = names.find(function(n){ return n.metadata && n.metadata.primary; }) || names[0];
      var primaryAddr = addrs.find(function(a){ return a.metadata && a.metadata.primary; }) || addrs[0];
      var primaryOrg  = orgs.find(function(o){ return o.metadata && o.metadata.primary; })  || orgs[0];
      var allPhones   = phones.map(function(p){ return p.canonicalForm || p.value || ''; }).filter(Boolean);
      var emailList   = emails.map(function(e){ return e.value || ''; }).filter(Boolean);
      var birthday = '';
      if (bdays.length > 0 && bdays[0].date) {
        var bd = bdays[0].date;
        var mo = bd.month ? ('0' + bd.month).slice(-2) : '';
        var dy = bd.day   ? ('0' + bd.day).slice(-2)   : '';
        var yr = bd.year  ? bd.year : '';
        if (mo && dy) birthday = (yr ? yr + '-' : '') + mo + '-' + dy;
      }
      contacts.push({ resourceName: person.resourceName || '', name: primaryName ? (primaryName.displayName || '') : '', allPhones: allPhones, emails: emailList, city: primaryAddr ? (primaryAddr.city || '') : '', address: primaryAddr ? (primaryAddr.formattedValue || '') : '', notes: bios.length ? (bios[0].value || '') : '', org: primaryOrg ? (primaryOrg.name || '') : '', title: primaryOrg ? (primaryOrg.title || '') : '', birthday: birthday });
    });
    pageToken = data.nextPageToken || null;
  } while (pageToken);
  return contacts;
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================

function fixEmployerPhoneColumns() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Employers');
  if (!sheet) { Logger.log('Employers sheet not found'); return; }
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIndex = {};
  headers.forEach(function(h, i) { var key = (h || '').toString().trim(); if (key && !colIndex[key]) colIndex[key] = i; });
  var phoneKeys = ['MobileNumber','MobileNumber2','MobileNumber3','MobileNumber4','MobileNumber5','MobileNumber6','MobileNumber7','MobileNumber8','MobileNumber9','MobileNumber10'];
  var nonPhoneHeaders = [], seen = {};
  headers.forEach(function(h) {
    var key = (h || '').toString().trim();
    if (!key || phoneKeys.indexOf(key) !== -1 || key === 'Email' || seen[key]) return;
    seen[key] = true; nonPhoneHeaders.push(key);
  });
  var newHeaders = nonPhoneHeaders.concat(phoneKeys);
  var newData = [newHeaders];
  for (var r = 1; r < data.length; r++) {
    var oldRow = data[r];
    newData.push(newHeaders.map(function(h) { var oldIdx = colIndex[h]; return (oldIdx !== undefined) ? (oldRow[oldIdx] || '') : ''; }));
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newHeaders.length).setValues(newData);
  Logger.log('Done! ' + newHeaders.length + ' columns, ' + newData.length + ' rows.');
}

function loadMessageLog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MessageLog');
  var sent = {}, lastDate = {};
  if (!sheet) return { sent: sent, lastDate: lastDate };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var empID = (data[i][3] || '').toString();
    var tplID = (data[i][4] || '').toString();
    if (!empID || !tplID) continue;
    var key = empID + '_' + tplID;
    sent[key] = true;
    var d = new Date(data[i][0]);
    if (!lastDate[key] || d > lastDate[key]) lastDate[key] = d;
  }
  return { sent: sent, lastDate: lastDate };
}

function getAllEmployers_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Employers');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return (h||'').toString().trim(); });
  var result  = [];
  for (var r = 1; r < data.length; r++) {
    if (!data[r][0]) continue;
    var emp = {};
    headers.forEach(function(h, i){ emp[h] = data[r][i] !== undefined ? data[r][i] : ''; });
    result.push(emp);
  }
  return result;
}

function parseDateFlexible_(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val) ? null : new Date(val);
  var s  = val.toString().trim();
  var m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m1) return new Date(parseInt(m1[3]), parseInt(m1[2])-1, parseInt(m1[1]));
  var m2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m2) return new Date(parseInt(m2[1]), parseInt(m2[2])-1, parseInt(m2[3]));
  var d = new Date(s);
  return isNaN(d) ? null : d;
}

function buildPersonalisedMessage_(template, emp, today) {
  return (template || '').replace(/\{(\w+)\}/g, function(_, key) {
    if (emp[key] !== undefined && emp[key] !== '') return emp[key];
    return '{' + key + '}';
  });
}

function logSentMessage_(emp, tpl, msg, sid) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName('MessageLog');
  if (!log) {
    log = ss.insertSheet('MessageLog');
    log.getRange('1:1').setValues([['Timestamp','Name','Phone','EmployerID','TemplateID','TemplateName','TriggerEvent','Message','TwilioSID','Status']]);
    log.setFrozenRows(1);
  }
  log.appendRow([ new Date(), (emp.EFN || emp.EmployerName || '').toString(), (emp.MobileNumber || emp.ContactNumber || '').toString(), (emp.EmployerID || '').toString(), (tpl.TemplateID || '').toString(), (tpl.TemplateName || '').toString(), (tpl.TriggerEvent || '').toString(), msg, sid || '', 'Sent' ]);
}

// ============================================================
// RULES REVIEW
// ============================================================

function saveRulesReviewNotes(jsonStr) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RulesReviewNotes') || ss.insertSheet('RulesReviewNotes');
  sheet.clearContents();
  sheet.getRange('A1').setValue('Notes');
  sheet.getRange('A2').setValue(jsonStr);
  return { success: true };
}

function getRulesReviewNotes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RulesReviewNotes');
  if (!sheet) return {};
  var val = sheet.getRange('A2').getValue();
  try { return JSON.parse(val || '{}'); } catch(e) { return {}; }
}

// ============================================================
// SYSTEM STATS
// ============================================================

function getSystemStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ['Workers','Employers','Suppliers','Local_Agencies','Nationalities'];
  var result = {};
  sheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    var count = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
    result[name] = count;
  });
  var natSheet = ss.getSheetByName('Nationalities');
  if (natSheet && natSheet.getLastRow() > 1) {
    var data    = natSheet.getDataRange().getValues();
    var headers = data[0];
    var actIdx  = headers.indexOf('Active');
    var active  = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][actIdx] === true || data[i][actIdx] === 'true' || data[i][actIdx] === 'TRUE') active++;
    }
    result['ActiveNationalities'] = active;
  }
  return result;
}

// ============================================================
// AUTO TRIGGERS — BATCH VERSION
// ============================================================

var BATCH_SIZE = 50;

function processAutoTriggers() {
  Logger.log('BATCH START');
  var props      = PropertiesService.getScriptProperties();
  var startIndex = parseInt(props.getProperty('TRIGGER_BATCH_INDEX') || '0');
  var result        = getTemplatesPublic();
  var templates     = (result && result.data) ? result.data : [];
  var autoTemplates = templates.filter(function(t) {
    return t.Active && t.TriggerEvent && t.TriggerEvent !== 'Manual';
  });
  if (!autoTemplates.length) { Logger.log('No auto-trigger templates.'); return; }
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheet     = ss.getSheetByName('Employers');
  if (!sheet) { Logger.log('Employers sheet not found'); return; }
  var totalRows = sheet.getLastRow() - 1;
  if (startIndex >= totalRows) { startIndex = 0; Logger.log('Batch reset to 0'); }
  var endIndex  = Math.min(startIndex + BATCH_SIZE, totalRows);
  Logger.log('Batch: rows ' + startIndex + ' to ' + (endIndex-1) + ' of ' + totalRows);
  var headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){ return (h||'').toString().trim(); });
  var batchData = sheet.getRange(startIndex + 2, 1, endIndex - startIndex, sheet.getLastColumn()).getValues();
  var batch = batchData.map(function(row) {
    var emp = {};
    headers.forEach(function(h, i){ emp[h] = row[i] !== undefined ? row[i] : ''; });
    return emp;
  }).filter(function(emp){ return emp.EmployerID; });
  Logger.log('Batch loaded: ' + batch.length + ' employers');
  var sentLog = loadMessageLog_();
  var today   = new Date(); today.setHours(0,0,0,0);
  var totalSent = 0, errors = [];
  var TRIGGER_FIELD = {
    'Birthday':         'BirthDate',
    'Anniversary':      'Anniversary',
    'PassportExpiry':   'PassportExpiry',
    'IkamaExpiry':      'IkamaExpiry',
    'InsuranceExpiry':  'InsuranceExpiry',
    'WorkPermitExpiry': 'WorkPermitExpiry',
    'WorkerArrival':    'ArrivalDate',
    'Scheduled':        'ScheduledDate',
    'DocumentReady':    null
  };
  autoTemplates.forEach(function(tpl) {
    var trigger    = tpl.TriggerEvent;
    var daysBefore = parseInt(tpl.DaysBefore)    || 0;
    var daysAfter  = parseInt(tpl.DaysAfter)      || 0;
    var sendOnce   = tpl.SendOnce === true || tpl.SendOnce === 'true' || tpl.SendOnce === 'TRUE';
    var minDays    = parseInt(tpl.MinDaysBetween) || 0;
    var dateField  = TRIGGER_FIELD[trigger];
    if (trigger === 'DocumentReady') return;
    batch.forEach(function(emp) {
      if (emp.EmployerDisabled === 'true' || emp.EmployerDisabled === true) return;
      var phone = '';
      for (var pi = 1; pi <= 10; pi++) {
        var pkey = pi === 1 ? 'MobileNumber' : 'MobileNumber' + pi;
        var pval = (emp[pkey] || '').toString().trim();
        if (pval && pval.length >= 7) { phone = pval; break; }
      }
      if (!phone) phone = (emp.ContactNumber || '').toString().trim();
      if (!phone) return;
      var shouldSend = false;
      if (dateField) {
        var rawDate = emp[dateField];
        if (!rawDate) return;
        var targetDate = parseDateFlexible_(rawDate);
        if (!targetDate) return;
        targetDate.setHours(0,0,0,0);
        if (trigger === 'Birthday' || trigger === 'Anniversary') {
          var checkDate = new Date(today);
          checkDate.setDate(checkDate.getDate() + daysBefore);
          shouldSend = (targetDate.getMonth() === checkDate.getMonth()) &&
                       (targetDate.getDate()  === checkDate.getDate());
        } else {
          var diffBefore = Math.round((targetDate - today) / 86400000);
          var diffAfter  = Math.round((today - targetDate) / 86400000);
          shouldSend = (diffBefore === daysBefore) || (daysAfter > 0 && diffAfter === daysAfter);
        }
      }
      if (!shouldSend) return;
      var empID  = (emp.EmployerID || '').toString();
      var tplID  = (tpl.TemplateID || '').toString();
      var logKey = empID + '_' + tplID;
      if (sendOnce && sentLog.sent[logKey]) return;
      if (minDays > 0) {
        var last = sentLog.lastDate[logKey];
        if (last && Math.floor((new Date() - last) / 86400000) < minDays) return;
      }
      var name = (emp.EFN || emp.EmployerName || '').split('(')[0].trim();
      var msg  = buildPersonalisedMessage_(tpl.MessageBody || '', emp, today);
      var res  = sendTwilioMessage_(phone, msg);
      if (res.success) {
        totalSent++;
        Logger.log('Sent [' + tpl.TemplateName + '] to ' + name + ' (' + phone + ')');
        logSentMessage_(emp, tpl, msg, res.sid);
        sentLog.sent[logKey] = true;
        sentLog.lastDate[logKey] = new Date();
      } else {
        errors.push(name + ': ' + res.error);
        Logger.log('FAILED [' + tpl.TemplateName + '] ' + name + ' — ' + res.error);
      }
    });
  });
  props.setProperty('TRIGGER_BATCH_INDEX', String(endIndex));
  Logger.log('Next batch at: ' + endIndex);
  Logger.log('DONE. Sent: ' + totalSent + ' | Errors: ' + errors.length);
}

function installBatchTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processAutoTriggers') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Removed old trigger');
    }
  });
  ScriptApp.newTrigger('processAutoTriggers').timeBased().everyMinutes(30).create();
  PropertiesService.getScriptProperties().setProperty('TRIGGER_BATCH_INDEX', '0');
  Logger.log('Batch trigger installed — every 30 min, 50 employers per run.');
}

function resetBatchIndex() {
  PropertiesService.getScriptProperties().setProperty('TRIGGER_BATCH_INDEX', '0');
  Logger.log('Batch index reset to 0');
}

// ============================================================
// DEBUG / FIX UTILITIES
// ============================================================

function debugFindBadDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;
    var data = sheet.getDataRange().getValues();
    data.forEach(function(row, ri) {
      row.forEach(function(val, ci) {
        if (val instanceof Date && isNaN(val.getTime())) {
          Logger.log('BAD DATE — Sheet: ' + name + ' | Row: ' + (ri+1) + ' | Col: ' + (ci+1));
        }
      });
    });
  });
  Logger.log('Debug complete');
}

function fixBadDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var msgSheet = ss.getSheetByName('Messages');
  if (msgSheet && msgSheet.getLastRow() > 1) {
    var col6 = msgSheet.getRange(2, 6, msgSheet.getLastRow() - 1, 1).getValues();
    col6.forEach(function(row, i) {
      if (row[0] instanceof Date && isNaN(row[0].getTime())) { col6[i][0] = ''; }
    });
    msgSheet.getRange(2, 6, col6.length, 1).setValues(col6);
    Logger.log('Fixed Messages column 6');
  }
  var empSheet = ss.getSheetByName('Employers');
  if (empSheet && empSheet.getLastRow() > 1) {
    var col35 = empSheet.getRange(2, 35, empSheet.getLastRow() - 1, 1).getValues();
    col35.forEach(function(row, i) {
      if (row[0] instanceof Date && isNaN(row[0].getTime())) { col35[i][0] = ''; }
    });
    empSheet.getRange(2, 35, col35.length, 1).setValues(col35);
    Logger.log('Fixed Employers column 35');
  }
  Logger.log('All bad dates fixed!');
}

function gsrCleanDriveFolders(parentType) {
  var mgsRoot    = getMgsRootFolder();
  var targetName = parentType === 'workers' ? 'Workers' : 'Employers_Documents';
  var parentFolder = null;
  var iter = mgsRoot.getFoldersByName(targetName);
  if (iter.hasNext()) parentFolder = iter.next();
  if (!parentFolder) return { success: false, error: targetName + ' folder not found inside ' + mgsRoot.getName() };
  var subs = parentFolder.getFolders();
  var deleted = 0;
  while (subs.hasNext()) { subs.next().setTrashed(true); deleted++; }
  var files = parentFolder.getFiles();
  var filesDeleted = 0;
  while (files.hasNext()) { files.next().setTrashed(true); filesDeleted++; }
  return { success: true, type: targetName, foldersDeleted: deleted, filesDeleted: filesDeleted };
}

// ============================================================
// TEST FUNCTIONS
// ============================================================

function testTwilioSend() {
  var result = sendTwilioMessage_('+9613823339', 'Test message from MGSDB2026');
  Logger.log(JSON.stringify(result));
}

function testGetMessages() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Messages');
  Logger.log('TEST - Sheet: ' + (sheet ? sheet.getName() : 'NULL'));
  Logger.log('TEST - Rows: ' + (sheet ? sheet.getLastRow() : 0));
}

function testGetTemplates() { Logger.log(JSON.stringify(getTemplates())); }

function testAutoDebug() {
  var result    = getTemplatesPublic();
  var templates = (result && result.data) ? result.data : [];
  Logger.log('Templates count: ' + templates.length);
  templates.forEach(function(t) {
    Logger.log(t.TemplateID + ' | ' + t.TemplateName + ' | Active:' + t.Active + ' | Trigger:' + t.TriggerEvent);
  });
}

function testGsrGetAll() {
  var result = gsrGetAll('Suppliers');
  Logger.log('Result type: ' + typeof result);
  Logger.log('Result: ' + JSON.stringify(result));
}

function testActionGetAll() {
  var result = actionGetAll({ sheet: 'Suppliers' });
  Logger.log('Result type: ' + typeof result);
  Logger.log('Result: ' + JSON.stringify(result));
}

function handleIncomingWhatsApp(params) {
  Logger.log('INCOMING CALLED — params: ' + JSON.stringify(params));
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Messages');
    if (!sheet) { Logger.log('Messages sheet NOT FOUND'); return; }
    var from  = (params.From || '').replace('whatsapp:','').replace(/\s/g,'');
    var body  = params.Body || '';
    var name  = params.ProfileName || from;
    var now   = new Date();
    var msgID = 'IN' + now.getTime();
    sheet.appendRow([
      msgID, now, 'Incoming', '', name, from,
      body, params.MediaUrl0||'', params.MediaContentType0||'',
      '','', 'Received', now, '', 'Incoming'
    ]);
    Logger.log('INCOMING SAVED — from: ' + from + ' | body: ' + body);
  } catch(e) {
    Logger.log('incomingWhatsApp error: ' + e.message);
  }
  return ContentService
    .createTextOutput('<?xml version="1.0"?><Response></Response>')
    .setMimeType(ContentService.MimeType.XML);
}

function handleStatusCallback(params) {
  try {
    var status  = params.MessageStatus || '';
    var sid     = params.MessageSid    || '';
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet   = ss.getSheetByName('Messages');
    if (!sheet || !sid) return { success: true };
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var colStatus = headers.indexOf('Status');
    var colErr    = headers.indexOf('ErrorMessage');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).indexOf(sid) !== -1 || String(data[i][0]) === sid) {
        if (colStatus !== -1) sheet.getRange(i+1, colStatus+1).setValue(status);
        if (status === 'failed' || status === 'undelivered') {
          if (colErr !== -1) sheet.getRange(i+1, colErr+1).setValue('Delivery failed: ' + status);
        }
        break;
      }
    }
  } catch(e) { Logger.log('statusCallback error: ' + e.message); }
  return { success: true };
}

function doPost(e) {
  try {
    if (e.parameter && e.parameter.method === 'OPTIONS') {
      return buildResponse({ success: true });
    }
    var body = {};
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      try {
        var pairs = (e.postData && e.postData.contents ? e.postData.contents : '').split('&');
        pairs.forEach(function(pair) {
          var kv = pair.split('=');
          if (kv[0]) body[decodeURIComponent(kv[0])] = decodeURIComponent((kv[1]||'').replace(/\+/g,' '));
        });
      } catch(parseErr2) {}
      if (e.parameter) {
        Object.keys(e.parameter).forEach(function(k){
          if (!body[k]) body[k] = e.parameter[k];
        });
      }
    }
    if (!body.action && e.parameter && e.parameter.action) {
      body.action = e.parameter.action;
    }
    if (!body.action && (body.From || body.MessageSid)) {
      if (body.MessageStatus) {
        body.action = 'statusCallback';
      } else {
        body.action = 'incomingWhatsApp';
      }
    }
    Logger.log('doPost action: ' + body.action + ' | From: ' + (body.From||'none'));
    if (body.action === 'incomingWhatsApp') {
      return handleIncomingWhatsApp(body);
    }
    if (body.action === 'statusCallback') {
      return buildResponse(handleStatusCallback(body));
    }
    var result = handleRequest(body);
    return buildResponse(result);
  } catch (err) {
    return buildResponse({ success: false, error: err.message });
  }
}

function testIncoming() {
  handleIncomingWhatsApp({
    From: 'whatsapp:+9613755199',
    Body: 'Test incoming message',
    ProfileName: 'Test Contact'
  });
}

function handleStatusCallback(params) {
  try {
    var status = params.MessageStatus || '';
    var sid    = params.MessageSid    || '';
    var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet  = ss.getSheetByName('Messages');
    if (!sheet || !sid) return { success: true };
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var colStatus = headers.indexOf('Status');
    var colErr    = headers.indexOf('ErrorMessage');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === sid) {
        if (colStatus !== -1) sheet.getRange(i+1, colStatus+1).setValue(status);
        if ((status === 'failed' || status === 'undelivered') && colErr !== -1)
          sheet.getRange(i+1, colErr+1).setValue('Delivery failed: ' + status);
        break;
      }
    }
  } catch(e) { Logger.log('statusCallback error: ' + e.message); }
  return { success: true };
}

function testGetMessages() {
  var result = getMessages();
  Logger.log('Count: ' + result.data.length);
  Logger.log('First: ' + JSON.stringify(result.data[0]));
}

function testMessagesDebug() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Messages');
  Logger.log('SS ID: ' + ss.getId());
  Logger.log('Sheet found: ' + (sheet ? 'YES' : 'NO'));
  Logger.log('Last row: ' + (sheet ? sheet.getLastRow() : 0));
}

function testDoPost() {
  var fakeEvent = {
    postData: {
      contents: 'From=whatsapp%3A%2B9613755199&Body=Hello+test&ProfileName=Test+User&MessageSid=SM123456'
    },
    parameter: {}
  };
  var result = doPost(fakeEvent);
  Logger.log('Result: ' + result.getContent());
}

function testContactsAccess() {
  var token = ScriptApp.getOAuthToken();
  var url = 'https://people.googleapis.com/v1/people/me/connections?personFields=names,phoneNumbers&pageSize=10';
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  Logger.log('Status: ' + response.getResponseCode());
  Logger.log('Response: ' + response.getContentText().substring(0, 500));
}

function testGetEmployers() {
  var result = actionGetAll({ sheet: 'Employers' });
  Logger.log('Count: ' + (result.data ? result.data.length : 0));
  Logger.log('First row: ' + JSON.stringify(result.data ? result.data[0] : null));
}

function testGsrGetAll() {
  var result = gsrGetAll('Employers');
  Logger.log('Success: ' + result.success);
  Logger.log('Count: ' + (result.data ? result.data.length : 0));
}

function testPayloadSize() {
  var result = actionGetAll({ sheet: 'Employers' });
  var json = JSON.stringify(result.data);
  Logger.log('Total bytes: ' + json.length);
  Logger.log('Per row avg: ' + Math.round(json.length / result.data.length));
  
  // Test with just 10 rows
  var small = result.data.slice(-10);
  Logger.log('10 rows bytes: ' + JSON.stringify(small).length);
}
function testReducedSize() {
  var result = actionGetAll({ sheet: 'Employers' });
  var reduced = result.data.map(function(row) {
    return {
      EmployerID: row.EmployerID,
      EmployerCode: row.EmployerCode,
      EFN: row.EFN,
      MobileNumber: row.MobileNumber,
      City: row.City,
      AgencyCode: row.AgencyCode,
      SyncStatus: row.SyncStatus,
      EmployerFolderID: row.EmployerFolderID,
      RequiredDocuments: row.RequiredDocuments,
      UploadedDocuments: row.UploadedDocuments,
      Disabled: row.Disabled
    };
  });
  Logger.log('Reduced bytes: ' + JSON.stringify(reduced).length);
  Logger.log('Reduced per row: ' + Math.round(JSON.stringify(reduced).length / reduced.length));
}
function testMinimalSize() {
  var result = actionGetAll({ sheet: 'Employers' });
  var reduced = result.data.map(function(row) {
    return {
      EmployerID: row.EmployerID,
      EmployerCode: row.EmployerCode,
      EFN: row.EFN,
      MobileNumber: row.MobileNumber,
      City: row.City,
      AgencyCode: row.AgencyCode,
      SyncStatus: row.SyncStatus,
      EmployerFolderID: row.EmployerFolderID,
      RequiredDocuments: row.RequiredDocuments,
      Disabled: row.Disabled
    };
  });
  Logger.log('Minimal bytes: ' + JSON.stringify(reduced).length);
}
function testTinySize() {
  var result = actionGetAll({ sheet: 'Employers' });
  var reduced = result.data.map(function(row) {
    return {
      EmployerID: row.EmployerID,
      EmployerCode: row.EmployerCode,
      EFN: row.EFN,
      MobileNumber: row.MobileNumber,
      City: row.City,
      AgencyCode: row.AgencyCode,
      SyncStatus: row.SyncStatus,
      Disabled: row.Disabled
    };
  });
  Logger.log('Tiny bytes: ' + JSON.stringify(reduced).length);
}
function testGsrReturn() {
  var result = gsrGetAll('Employers');
  Logger.log('Type: ' + typeof result);
  Logger.log('Success: ' + result.success);
  Logger.log('Count: ' + (result.data ? result.data.length : 'no data'));
  Logger.log('JSON size: ' + JSON.stringify(result).length);
}
