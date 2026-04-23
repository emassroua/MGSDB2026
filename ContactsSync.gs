// ============================================================
// MGS - Google Contacts Sync Module v5.0
// Syncs MGSC contacts from Personal Google Account → Employers Sheet
//
// FIX: No longer stores contacts in ScriptProperties (quota fix).
//      Uses API pageToken to resume between runs — zero storage issues.
//
// SETUP ORDER (do this once):
// 1. Run getAuthorizationUrl()   → open URL in browser, sign in with PERSONAL account
// 2. Copy "code=" from redirect URL → paste into saveAuthCode() → run it
// 3. Run syncGoogleContactsToDB() → starts sync
// 4. Run setupContactsSyncTrigger() → auto-runs every 15 min
// ============================================================

var SYNC_PREFIX = 'MGSC'; // fallback only

function getAgencyCodes() {
  try {
    var result = actionGetAll({ sheet: 'Local_Agencies' });
    if (!result || !result.data || result.data.length === 0) return ['MGSC'];
    var codes = result.data
      .filter(function(a) { return a.AgencyCode && a.AgencyStatus !== 'disabled'; })
      .map(function(a) { return a.AgencyCode.trim().toUpperCase(); });
    return codes.length > 0 ? codes : ['MGSC'];
  } catch(e) {
    Logger.log('getAgencyCodes error: ' + e.message);
    return ['MGSC'];
  }
}

function detectAgencyCode(contactName, agencyCodes) {
  if (!contactName) return null;
  var name = contactName.trim().toUpperCase();
  for (var i = 0; i < agencyCodes.length; i++) {
    if (name.indexOf(agencyCodes[i]) === 0) return agencyCodes[i];
  }
  return null;
}

var SYNC_SHEET      = 'Employers';
var SYNC_ID_FIELD   = 'GoogleContactID';
var SYNC_STATUS_COL = 'SyncStatus';
var SYNC_LOG_SHEET  = 'Sync_Log';
var PAGE_SIZE       = 200; // contacts per API page = contacts processed per run

var CONTACTS_CLIENT_ID     = '1071315446230-4hecmnn9f83pcm054epck06snv3mqr57.apps.googleusercontent.com';
var CONTACTS_CLIENT_SECRET = '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';
var CONTACTS_REDIRECT_URI = 'https://developers.google.com/oauthplayground';
var CONTACTS_SCOPE         = 'https://www.googleapis.com/auth/contacts.readonly';

// ScriptProperties keys (small values only — no contact data stored!)
var PROP_SYNC_TOKEN  = 'CONTACTS_SYNC_TOKEN';     // saved after full sync completes
var PROP_PAGE_TOKEN  = 'CONTACTS_PAGE_TOKEN';     // current page position during full sync
var PROP_BATCH_STATS = 'BATCH_ACCUMULATED_STATS'; // running totals across runs
var PROP_FULL_DONE   = 'CONTACTS_FULL_SYNC_DONE'; // 'true' when full sync completed

// ============================================================
// AUTHORIZATION
// ============================================================

function getAuthorizationUrl() {
  var url = 'https://accounts.google.com/o/oauth2/auth'
    + '?client_id='    + encodeURIComponent(CONTACTS_CLIENT_ID)
    + '&redirect_uri=' + encodeURIComponent(CONTACTS_REDIRECT_URI)
    + '&response_type=code'
    + '&scope='        + encodeURIComponent(CONTACTS_SCOPE)
    + '&access_type=offline'
    + '&prompt=consent';

  Logger.log('=== OPEN THIS URL IN YOUR BROWSER ===');
  Logger.log(url);
  Logger.log('=====================================');
  Logger.log('After authorizing → copy value after "code=" → paste into saveAuthCode()');
  return url;
}

function saveAuthCode() {
var AUTH_CODE = '4/0Aci98E8kbhfcJjcYS3APmg76yqyyjIoEjOp3cDBuh8hCEDnvLmdQCCNMwSXImxlEhRnZnA';
  if (AUTH_CODE === 'PASTE_YOUR_CODE_HERE' || !AUTH_CODE.trim()) {
    Logger.log('Paste your auth code first, then run again!');
    return;
  }

  var response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      code:          AUTH_CODE,
      client_id:     CONTACTS_CLIENT_ID,
      client_secret: CONTACTS_CLIENT_SECRET,
      redirect_uri:  CONTACTS_REDIRECT_URI,
      grant_type:    'authorization_code'
    },
    muteHttpExceptions: true
  });

  var data = JSON.parse(response.getContentText());
  if (data.error) {
    Logger.log('Error: ' + data.error + ' — ' + data.error_description);
    return;
  }

  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    'CONTACTS_ACCESS_TOKEN':  data.access_token,
    'CONTACTS_REFRESH_TOKEN': data.refresh_token || '',
    'CONTACTS_TOKEN_EXPIRY':  String(new Date().getTime() + (data.expires_in * 1000))
  });

  clearSyncState(props);
  Logger.log('Authorization saved! Now run syncGoogleContactsToDB()');
}

function getValidAccessToken() {
  var props        = PropertiesService.getScriptProperties();
  var accessToken  = props.getProperty('CONTACTS_ACCESS_TOKEN');
  var expiry       = parseInt(props.getProperty('CONTACTS_TOKEN_EXPIRY') || '0');
  var refreshToken = props.getProperty('CONTACTS_REFRESH_TOKEN');

  if (accessToken && new Date().getTime() < expiry - 300000) return accessToken;

  if (!refreshToken) {
    Logger.log('No refresh token. Run getAuthorizationUrl() to re-authorize.');
    return null;
  }

  Logger.log('Refreshing access token...');
  var response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      refresh_token: refreshToken,
      client_id:     CONTACTS_CLIENT_ID,
      client_secret: CONTACTS_CLIENT_SECRET,
      grant_type:    'refresh_token'
    },
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());
  if (result.error) {
    Logger.log('Token refresh failed: ' + result.error_description);
    return null;
  }

  props.setProperty('CONTACTS_ACCESS_TOKEN', result.access_token);
  props.setProperty('CONTACTS_TOKEN_EXPIRY', String(new Date().getTime() + (result.expires_in * 1000)));
  return result.access_token;
}

function revokeContactsAuth() {
  var props = PropertiesService.getScriptProperties();
  ['CONTACTS_ACCESS_TOKEN', 'CONTACTS_REFRESH_TOKEN', 'CONTACTS_TOKEN_EXPIRY'].forEach(function(k) {
    props.deleteProperty(k);
  });
  clearSyncState(props);
  Logger.log('Authorization revoked. Run getAuthorizationUrl() to re-authorize.');
  return { success: true };
}

function clearSyncState(props) {
  if (!props) props = PropertiesService.getScriptProperties();
  [PROP_SYNC_TOKEN, PROP_PAGE_TOKEN, PROP_BATCH_STATS, PROP_FULL_DONE].forEach(function(k) {
    props.deleteProperty(k);
  });
}

// ============================================================
// MAIN SYNC — auto-detects mode
// ============================================================

function syncGoogleContactsToDB() {
  var startTime = new Date();
  var props     = PropertiesService.getScriptProperties();
  var syncToken = props.getProperty(PROP_SYNC_TOKEN);

  if (syncToken) {
    return runIncrementalSync(startTime, props, syncToken);
  } else {
    return runFullSyncPage(startTime, props);
  }
}

// ============================================================
// FULL SYNC — fetches ONE page (50 contacts) per run
// Saves pageToken and resumes on next trigger run
// ============================================================

function runFullSyncPage(startTime, props) {
  var token = getValidAccessToken();
  if (!token) return { success: false, error: 'Not authorized. Run getAuthorizationUrl().' };

  var pageToken = props.getProperty(PROP_PAGE_TOKEN) || null;
  var statsRaw  = props.getProperty(PROP_BATCH_STATS);
  var stats     = statsRaw
    ? JSON.parse(statsRaw)
    : { inserted: 0, updated: 0, inactive: 0, folders: 0, errors: 0, total: 0, mode: 'full_batch' };

  if (!pageToken) Logger.log('=== FULL SYNC START ===');

  // Fetch ONE page from People API
  var url = 'https://people.googleapis.com/v1/people/me/connections'
  + '?personFields=names,emailAddresses,phoneNumbers,addresses,organizations,biographies,userDefined,birthdays'
    + '&pageSize=' + PAGE_SIZE
    + '&requestSyncToken=true'
    + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('People API error: ' + response.getContentText());
    return { success: false, error: 'People API error' };
  }

  var data        = JSON.parse(response.getContentText());
  var connections = data.connections || [];
  var nextPage    = data.nextPageToken || null;
  var newToken    = data.nextSyncToken  || null;

  // Filter MGSC contacts from this page
  var _agencyCodes = getAgencyCodes();
  var mgscContacts = connections.filter(function(p) {
    var name = (p.names && p.names[0]) ? (p.names[0].displayName || '') : '';
    return detectAgencyCode(name, _agencyCodes) !== null;
  });

  Logger.log('Page: ' + mgscContacts.length + ' MGSC | nextPage: ' + (nextPage ? 'yes' : 'done'));

  // Process this page's contacts
  if (mgscContacts.length > 0) {
    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet       = ss.getSheetByName(SYNC_SHEET) || ss.insertSheet(SYNC_SHEET);
    var existingMap = buildExistingEmployerMap(sheet);
    processContacts(mgscContacts, existingMap, stats);
  }

  stats.total = (stats.total || 0) + connections.length;

  if (nextPage) {
    // More pages — save position and wait for next trigger run
    props.setProperty(PROP_PAGE_TOKEN,  nextPage);
    props.setProperty(PROP_BATCH_STATS, JSON.stringify(stats));
    Logger.log('Progress: ' + stats.inserted + ' inserted, ' + stats.updated + ' updated so far...');
    logSyncResult(stats, startTime, 'IN_PROGRESS');
    return {
      success:    true,
      inProgress: true,
      mode:       'full_batch',
      inserted:   stats.inserted,
      updated:    stats.updated,
      folders:    stats.folders,
      errors:     stats.errors
    };
  } else {
    // All pages done!
    if (newToken) props.setProperty(PROP_SYNC_TOKEN, newToken);
    props.deleteProperty(PROP_PAGE_TOKEN);
    props.deleteProperty(PROP_BATCH_STATS);
    props.setProperty(PROP_FULL_DONE, 'true');

    Logger.log('FULL SYNC COMPLETE!');
    Logger.log('  Inserted : ' + stats.inserted);
    Logger.log('  Updated  : ' + stats.updated);
    Logger.log('  Folders  : ' + stats.folders);
    Logger.log('  Errors   : ' + stats.errors);

    logSyncResult(stats, startTime, 'OK');
    return { success: true, inProgress: false, mode: 'full_batch', inserted: stats.inserted, updated: stats.updated, inactive: stats.inactive, folders: stats.folders, errors: stats.errors, total: stats.total };
  }
}

// ============================================================
// INCREMENTAL SYNC — only changed contacts since last sync
// ============================================================

function runIncrementalSync(startTime, props, savedSyncToken) {
  var token = getValidAccessToken();
  if (!token) return { success: false, error: 'Not authorized.' };

  var stats    = { inserted: 0, updated: 0, inactive: 0, deleted: 0, folders: 0, errors: 0, total: 0, mode: 'incremental' };
  var changed  = [], deleted = [], nextPage = null, newSyncToken = savedSyncToken;

  Logger.log('=== INCREMENTAL SYNC ===');

  do {
    var url = 'https://people.googleapis.com/v1/people/me/connections'
     + '?personFields=names,emailAddresses,phoneNumbers,addresses,organizations,biographies,userDefined,metadata,birthdays'
      + '&pageSize=1000'
      + '&requestSyncToken=true'
      + '&syncToken=' + encodeURIComponent(savedSyncToken)
      + (nextPage ? '&pageToken=' + encodeURIComponent(nextPage) : '');

    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 410) {
      Logger.log('Sync token expired — starting fresh full sync');
      clearSyncState(props);
      return runFullSyncPage(startTime, props);
    }

    if (response.getResponseCode() !== 200) {
      Logger.log('API error: ' + response.getContentText());
      break;
    }

    var data = JSON.parse(response.getContentText());
    (data.connections || []).forEach(function(person) {
      if (person.metadata && person.metadata.deleted === true) {
        deleted.push(person.resourceName);
      } else {
        var name = (person.names && person.names[0]) ? (person.names[0].displayName || '') : '';
        var _incAgencyCodes = getAgencyCodes();
         if (detectAgencyCode(name, _incAgencyCodes) !== null) changed.push(person);
      }
    });

    nextPage     = data.nextPageToken || null;
    newSyncToken = data.nextSyncToken  || newSyncToken;

  } while (nextPage);

  stats.total = changed.length + deleted.length;

  if (changed.length > 0 || deleted.length > 0) {
    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet       = ss.getSheetByName(SYNC_SHEET) || ss.insertSheet(SYNC_SHEET);
    var existingMap = buildExistingEmployerMap(sheet);

    processContacts(changed, existingMap, stats);

    deleted.forEach(function(resourceName) {
      try {
        var r = existingMap[resourceName];
        if (r && r.data[SYNC_STATUS_COL] !== 'Inactive') {
          r.data[SYNC_STATUS_COL] = 'Inactive';
          r.data['LastSyncDate']  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
          actionSaveRow({ sheet: SYNC_SHEET, row: r.data });
          stats.inactive++;
        }
      } catch(e) { stats.errors++; }
    });
  }

  if (newSyncToken) props.setProperty(PROP_SYNC_TOKEN, newSyncToken);

  Logger.log('Incremental done | changed:' + changed.length + ' deleted:' + deleted.length);
  logSyncResult(stats, startTime, 'OK');
  return { success: true, inProgress: false, mode: 'incremental', inserted: stats.inserted, updated: stats.updated, inactive: stats.inactive, folders: stats.folders, errors: stats.errors, total: stats.total };
}

// ============================================================
// PROCESS CONTACTS
// ============================================================

function processContacts(contacts, existingMap, stats) {
  contacts.forEach(function(person) {
    try {
      var contactId = person.resourceName || '';
      if (!contactId) return;

      var employerRow = mapContactToEmployer(person);
      var existingRow = existingMap[contactId];

      if (existingRow) {
        var merged = mergeEmployerData(existingRow.data, employerRow);
        actionSaveRow({ sheet: SYNC_SHEET, row: merged });
        stats.updated++;
      } else {
        var nextId = actionGetNextID({ sheet: SYNC_SHEET });
         employerRow['EmployerID'] = nextId.nextID || 1;
        if (!employerRow['EmployerCode']) {
         employerRow['EmployerCode'] = (employerRow['AgencyCode']||'MGSC') + '-' + employerRow['EmployerID'];
}
     actionSaveRow({ sheet: SYNC_SHEET, row: employerRow });
        stats.inserted++;

        try {
          var folderName   = buildEmployerFolderName(employerRow);
          var folderResult = actionCreateFolder({ folderName: folderName, parentType: 'employers' });
          if (folderResult.success) {
            employerRow['FolderUrl']          = folderResult.folderUrl;
            employerRow['EmployerFolderID']   = folderResult.folderId;
            employerRow['EmployerFolderName'] = folderName;
            employerRow['EmployerFolderUrl']  = folderResult.folderUrl;
            actionSaveRow({ sheet: SYNC_SHEET, row: employerRow });
            stats.folders++;
          }
        } catch(fe) {
          Logger.log('Folder error: ' + fe.message);
          stats.errors++;
        }
      }
    } catch(e) {
      Logger.log('Contact error: ' + e.message);
      stats.errors++;
    }
  });
}

// ============================================================
// MAP GOOGLE CONTACT → EMPLOYER ROW
// ============================================================

function mapContactToEmployer(person) {
  var row = {};
  row[SYNC_ID_FIELD] = person.resourceName || '';

  if (person.names && person.names[0]) {
    var fullName = person.names[0].displayName || '';
    var _codes = getAgencyCodes();
    var _detected = detectAgencyCode(fullName, _codes) || 'MGSC';
    row['EFN'] = fullName.replace(new RegExp('^' + _detected + '\\s*[-:]?\\s*', 'i'), '').trim();
    row['AgencyCode'] = _detected;
    row['EmployerName'] = row['EFN'];

    try {
      var agencies = actionGetAll({ sheet: 'Local_Agencies' });
      if (agencies && agencies.data) {
        var matchedAgency = agencies.data.find(function(a) {
          return a.AgencyCode && a.AgencyCode.trim().toUpperCase() === _detected;
        });
        if (matchedAgency) {
          row['AgencyID'] = matchedAgency.AgencyID;
          Logger.log('Linked to agency: ' + _detected + ' (ID: ' + matchedAgency.AgencyID + ')');
        }
      }
    } catch(ae) {
      Logger.log('Agency lookup error: ' + ae.message);
    }
  }

  if (person.phoneNumbers && person.phoneNumbers.length > 0) {
    person.phoneNumbers.forEach(function(ph) {
      var type = (ph.type || ph.canonicalForm || '').toLowerCase();
      var val  = ph.value || '';
      if (!val) return;
      if (type === 'mobile' || type === 'cell') {
        row['MobileNumber'] = row['MobileNumber'] || val;
      } else if (type === 'work' || type === 'main') {
        row['ContactNumber'] = row['ContactNumber'] || val;
      } else {
        row['MobileNumber'] = row['MobileNumber'] || val;
      }
    });
    if (!row['MobileNumber'] && !row['ContactNumber']) {
      row['MobileNumber'] = person.phoneNumbers[0].value || '';
    }
  }

  if (person.birthdays && person.birthdays[0] && person.birthdays[0].date) {
    var bd = person.birthdays[0].date;
    var month = String(bd.month).padStart(2,'0');
    var day   = String(bd.day).padStart(2,'0');
    var year  = bd.year ? bd.year : 1900;
    row['BirthDate'] = year + '-' + month + '-' + day;
  }
  if (person.emailAddresses && person.emailAddresses[0]) {
    row['EmailAddress'] = person.emailAddresses[0].value || '';
  }
  if (person.addresses && person.addresses[0]) {
    var addr = person.addresses[0];
    row['Address'] = addr.streetAddress || '';
    row['City']    = addr.city          || '';
    row['Country'] = addr.country       || '';
  }
  if (person.biographies && person.biographies[0]) {
    row['Notes'] = person.biographies[0].value || '';
  }
  if (person.organizations && person.organizations[0] && !row['EFN']) {
    row['EFN'] = row['EmployerName'] = person.organizations[0].name || '';
  }
  if (person.userDefined) {
    person.userDefined.forEach(function(f) {
      var k = (f.key || '').trim();
      var v = (f.value || '').trim();
      if (k) row[k] = v;
    });
  }

  row[SYNC_STATUS_COL] = 'Active';
  row['LastSyncDate'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  return row;
}

// ============================================================
// HELPERS
// ============================================================

function buildExistingEmployerMap(sheet) {
  var map     = {};
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return map;

  var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var idColIdx = headers.indexOf(SYNC_ID_FIELD);
  if (idColIdx < 0) return map;

  sheet.getRange(2, 1, lastRow - 1, lastCol).getValues().forEach(function(row, i) {
    var cid = String(row[idColIdx] || '').trim();
    if (!cid) return;
    var obj = {};
    headers.forEach(function(h, j) { if (h) obj[h] = row[j] == null ? '' : row[j]; });
    map[cid] = { data: obj, rowIndex: i + 2 };
  });
  return map;
}

function mergeEmployerData(existing, incoming) {
  var PROTECTED = ['EmployerID', 'FolderUrl', 'EmployerFolderID', 'EmployerFolderName',
                   'EmployerFolderUrl', 'AgencyID', 'RequiredDocuments', 'UploadedDocuments'];
  var merged = {};
  Object.keys(existing).forEach(function(k) { merged[k] = existing[k]; });
  Object.keys(incoming).forEach(function(k) {
    if (PROTECTED.indexOf(k) !== -1 && merged[k] && merged[k] !== '') return;
    merged[k] = incoming[k];
  });
  return merged;
}

function buildEmployerFolderName(employerRow) {
  var agencyCode = employerRow['AgencyCode'] || 'MGS';
  var employerId = employerRow['EmployerID'] || 0;
  var name       = (employerRow['EFN'] || employerRow['EmployerName'] || 'Unknown').trim();
  var date       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return agencyCode + '(' + employerId + ') - ' + name + ' - ' + date;
}

// ============================================================
// SYNC LOG
// ============================================================

function logSyncResult(stats, startTime, status) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SYNC_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(SYNC_LOG_SHEET);
      sheet.getRange(1, 1, 1, 9).setValues([[
        'Timestamp', 'Mode', 'Status', 'Total', 'Inserted', 'Updated', 'Inactive', 'Folders', 'Errors'
      ]]);
    }
    sheet.appendRow([
      Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      stats.mode || 'full_batch', status,
      stats.total || 0, stats.inserted || 0, stats.updated || 0,
      stats.inactive || 0, stats.folders || 0, stats.errors || 0
    ]);
    var lastRow = sheet.getLastRow();
    if (lastRow > 501) sheet.deleteRows(2, lastRow - 501);
  } catch(e) { Logger.log('Log error: ' + e.message); }
}

// ============================================================
// PROGRESS (polled by frontend)
// ============================================================

function getSyncProgress() {
  var props     = PropertiesService.getScriptProperties();
  var pageToken = props.getProperty(PROP_PAGE_TOKEN);
  var syncToken = props.getProperty(PROP_SYNC_TOKEN);
  var statsRaw  = props.getProperty(PROP_BATCH_STATS);
  var stats     = statsRaw ? JSON.parse(statsRaw) : { inserted: 0, updated: 0, folders: 0, errors: 0 };
  return {
    inProgress: !!pageToken,
    mode:       pageToken ? 'full_batch' : (syncToken ? 'incremental' : 'not_started'),
    inserted:   stats.inserted || 0,
    updated:    stats.updated  || 0,
    folders:    stats.folders  || 0,
    errors:     stats.errors   || 0
  };
}

// ============================================================
// TRIGGER SETUP
// ============================================================

function setupContactsSyncTrigger() {
  removeContactsSyncTrigger();
  ScriptApp.newTrigger('syncGoogleContactsToDB').timeBased().everyMinutes(15).create();
  Logger.log('Trigger set: every 15 minutes.');
  return { success: true };
}

function removeContactsSyncTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncGoogleContactsToDB') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  if (removed) Logger.log('Removed ' + removed + ' trigger(s)');
  return { success: true, removed: removed };
}

function checkSyncTriggerStatus() {
  var found = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncGoogleContactsToDB') found++;
  });
  var progress = getSyncProgress();
  Logger.log('Triggers: ' + found + ' | Mode: ' + progress.mode);
  return { active: found > 0, mode: progress.mode };
}

function forceFullSync() {
  clearSyncState();
  Logger.log('Reset done — next sync will be a fresh full sync.');
  return { success: true };
}

// ============================================================
// FRONTEND ACTIONS
// ============================================================

function actionSyncContacts(params) {
  try {
    var result = syncGoogleContactsToDB();
    return {
      success:    true,
      inProgress: result.inProgress || false,
      mode:       result.mode       || 'full_batch',
      inserted:   result.inserted   || 0,
      updated:    result.updated    || 0,
      inactive:   result.inactive   || 0,
      folders:    result.folders    || 0,
      errors:     result.errors     || 0,
      total:      result.total      || 0
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function actionGetSyncProgress(params) {
  try {
    var p = getSyncProgress();
    return { success: true, inProgress: p.inProgress, mode: p.mode, inserted: p.inserted, updated: p.updated, folders: p.folders, errors: p.errors };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function debugFirstContacts() {
  var token = getValidAccessToken();
  var url = 'https://people.googleapis.com/v1/people/me/connections'
    + '?personFields=names&pageSize=20';
  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });
  var data = JSON.parse(response.getContentText());
  (data.connections || []).forEach(function(p) {
    var name = (p.names && p.names[0]) ? p.names[0].displayName : '(no name)';
    Logger.log(name);
  });
}

function debugWhoAmI() {
  var token = getValidAccessToken();
  var response = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v1/userinfo', {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });
  Logger.log(response.getContentText());
}

function debugTokens() {
  var props = PropertiesService.getScriptProperties();
  Logger.log('Access Token: ' + (props.getProperty('CONTACTS_ACCESS_TOKEN') ? 'EXISTS' : 'MISSING'));
  Logger.log('Refresh Token: ' + (props.getProperty('CONTACTS_REFRESH_TOKEN') ? 'EXISTS' : 'MISSING'));
  Logger.log('Expiry: ' + props.getProperty('CONTACTS_TOKEN_EXPIRY'));
  Logger.log('Now: ' + new Date().getTime());
}

function debugTokenInfo() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('CONTACTS_ACCESS_TOKEN');
  var response = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?access_token=' + token, {
    muteHttpExceptions: true
  });
  Logger.log(response.getContentText());
}

function debugAfterSync() {
  var props = PropertiesService.getScriptProperties();
  Logger.log('PAGE_TOKEN: ' + props.getProperty('CONTACTS_PAGE_TOKEN'));
  Logger.log('SYNC_TOKEN: ' + props.getProperty('CONTACTS_SYNC_TOKEN'));
  Logger.log('BATCH_STATS: ' + props.getProperty('BATCH_ACCUMULATED_STATS'));
}

function checkSyncState() {
  var props = PropertiesService.getScriptProperties().getProperties();
  Logger.log('PAGE_TOKEN: ' + props['CONTACTS_PAGE_TOKEN']);
  Logger.log('SYNC_TOKEN: ' + props['CONTACTS_SYNC_TOKEN']);
  Logger.log('FULL_SYNC_DONE: ' + props['CONTACTS_FULL_SYNC_DONE']);
  Logger.log('ACCUMULATED_STATS: ' + props['BATCH_ACCUMULATED_STATS']);
  Logger.log('TOKEN_EXPIRY: ' + props['CONTACTS_TOKEN_EXPIRY']);
}

function testAgencyCodes() {
  var codes = getAgencyCodes();
  Logger.log('Agency codes found: ' + codes.join(', '));
}

function testFolderCreation() {
  var result = actionCreateFolder({ 
    folderName: 'GSA(1315) - Toni Abou Assi - 2026-03-10', 
    parentType: 'employers' 
  });
  Logger.log(JSON.stringify(result));
}

function debugMissingFolders() {
  var employers = actionGetAll({ sheet: 'Employers' });
  [1315, 1316, 1317].forEach(function(id) {
    var emp = employers.data.find(function(e) { return e.EmployerID == id; });
    if (emp) {
      Logger.log('ID:' + id + ' | AgencyCode:' + emp.AgencyCode + ' | FolderID:' + emp.EmployerFolderID + ' | EFN:' + emp.EFN);
      Logger.log('FolderName would be: ' + buildEmployerFolderName(emp));
    }
  });
}

function linkAllEmployersToAgencies() {
  var employers = actionGetAll({ sheet: 'Employers' });
  var agencies  = actionGetAll({ sheet: 'Local_Agencies' });
  if (!employers || !agencies) return;
  
  var fixed = 0;
  employers.data.forEach(function(emp) {
    if (emp.AgencyID) return; // already linked
    if (!emp.AgencyCode) return; // no agency code
    var ag = agencies.data.find(function(a) {
      return a.AgencyCode && a.AgencyCode.trim().toUpperCase() === emp.AgencyCode.trim().toUpperCase();
    });
    if (ag) {
      emp.AgencyID = ag.AgencyID;
      actionSaveRow({ sheet: 'Employers', row: emp });
      fixed++;
      Logger.log('Fixed: ' + emp.EFN + ' → ' + emp.AgencyCode + ' (ID:' + ag.AgencyID + ')');
    }
  });
  Logger.log('Done! Fixed: ' + fixed + ' employers');
}

function debugFindMGSC() {
  var token = getValidAccessToken();
  var found = 0;
  var pageToken = null;
  var page = 0;

  do {
    page++;
    var url = 'https://people.googleapis.com/v1/people/me/connections'
      + '?personFields=names,birthdays&pageSize=1000'
      + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    var data = JSON.parse(response.getContentText());
    Logger.log('Page ' + page + ': ' + (data.connections || []).length + ' contacts');

    (data.connections || []).forEach(function(p) {
      var name = (p.names && p.names[0]) ? p.names[0].displayName : '';
      if (name.indexOf('MGSC') === 0 || name.indexOf('GSA') === 0 || name.indexOf('KHO') === 0) {
        var bday = (p.birthdays && p.birthdays[0] && p.birthdays[0].date) ? JSON.stringify(p.birthdays[0].date) : 'NO BIRTHDAY';
        Logger.log(name + ' | ' + bday);
        found++;
      }
    });

    pageToken = data.nextPageToken || null;
  } while (pageToken);

  Logger.log('Total MGSC/GSA/KHO found: ' + found);
}

function findNada() {
  var token = getValidAccessToken();
  var url = 'https://people.googleapis.com/v1/people/me/connections'
    + '?personFields=names,phoneNumbers&pageSize=1000';
  var found = false;
  do {
    var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());
    (data.connections || []).forEach(function(p) {
      var name = (p.names && p.names[0]) ? p.names[0].displayName : '';
      if (name.indexOf('Nada') !== -1) {
        Logger.log('FOUND: ' + name + ' | ' + JSON.stringify(p.phoneNumbers));
        found = true;
      }
    });
    url = data.nextPageToken 
      ? 'https://people.googleapis.com/v1/people/me/connections?personFields=names,phoneNumbers&pageSize=1000&pageToken=' + encodeURIComponent(data.nextPageToken)
      : null;
  } while (url);
  if (!found) Logger.log('Nada not found in contacts');
}

function findNada() {
  var token = getValidAccessToken();
  var url = 'https://people.googleapis.com/v1/people/me/connections'
    + '?personFields=names,phoneNumbers&pageSize=1000';
  var found = false;
  do {
    var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());
    (data.connections || []).forEach(function(p) {
      var name = (p.names && p.names[0]) ? p.names[0].displayName : '';
      if (name.indexOf('Nada') !== -1) {
        Logger.log('FOUND: ' + name + ' | ' + JSON.stringify(p.phoneNumbers));
        found = true;
      }
    });
    url = data.nextPageToken 
      ? 'https://people.googleapis.com/v1/people/me/connections?personFields=names,phoneNumbers&pageSize=1000&pageToken=' + encodeURIComponent(data.nextPageToken)
      : null;
  } while (url);
  if (!found) Logger.log('Nada not found in contacts');
}

// ============================================================
// FIND & SYNC SINGLE CONTACT BY NAME
// Usage: findAndSyncContact('Nada') or findAndSyncContact('MGSC John')
// ============================================================
function findAndSyncContact(searchName) {
  if (!searchName) {
    Logger.log('Please provide a name to search for.');
    return;
  }

  var token = getValidAccessToken();
  if (!token) { Logger.log('Not authorized. Run getAuthorizationUrl()'); return; }

  Logger.log('🔍 Searching for: ' + searchName);

  var found   = [];
  var pageToken = null;

  do {
    var url = 'https://people.googleapis.com/v1/people/me/connections'
      + '?personFields=names,phoneNumbers,emailAddresses,addresses,organizations,biographies,birthdays,userDefined'
      + '&pageSize=1000'
      + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    var data = JSON.parse(resp.getContentText());

    (data.connections || []).forEach(function(p) {
      var name = (p.names && p.names[0]) ? p.names[0].displayName : '';
      if (name.toLowerCase().indexOf(searchName.toLowerCase()) !== -1) {
        found.push(p);
        Logger.log('✅ Found: ' + name);
      }
    });

    pageToken = data.nextPageToken || null;
  } while (pageToken);

  if (found.length === 0) {
    Logger.log('❌ No contact found matching: ' + searchName);
    return { success: false, error: 'Contact not found' };
  }

  Logger.log('📋 Found ' + found.length + ' match(es) — syncing now...');

  // Sync each found contact to Employers sheet
  var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet       = ss.getSheetByName(SYNC_SHEET) || ss.insertSheet(SYNC_SHEET);
  var existingMap = buildExistingEmployerMap(sheet);
  var stats       = { inserted: 0, updated: 0, inactive: 0, folders: 0, errors: 0, total: 0 };

  processContacts(found, existingMap, stats);

  Logger.log('✅ Sync complete!');
  Logger.log('   Inserted : ' + stats.inserted);
  Logger.log('   Updated  : ' + stats.updated);
  Logger.log('   Folders  : ' + stats.folders);
  Logger.log('   Errors   : ' + stats.errors);

  return { success: true, inserted: stats.inserted, updated: stats.updated, errors: stats.errors };
}

// Quick shortcut — run this and change the name
function syncByName() {
  var NAME = 'Nada Samir'; // ← change this to any name
  findAndSyncContact(NAME);
}

function fixMissingEmployerCodes() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName('Employers');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol   = headers.indexOf('EmployerID');
  var codeCol = headers.indexOf('EmployerCode');
  var efnCol  = headers.indexOf('EFN');
  var agCol   = headers.indexOf('AgencyCode');
  var fixed   = 0;

  // Get all agency codes to strip from EFN
  var agencyCodes = getAgencyCodes();

  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][codeCol]||'').trim();
    if (code) continue; // already has code
    var id      = data[i][idCol];
    var efn     = String(data[i][efnCol]||'').trim();
    var agCode  = String(data[i][agCol]||'').trim().toUpperCase();
    if (!id || !efn) continue;

    // Strip agency prefix from EFN if present
    var cleanEFN = efn;
    var prefixes = agCode ? [agCode] : agencyCodes;
    prefixes.forEach(function(prefix) {
      cleanEFN = cleanEFN.replace(new RegExp('^' + prefix + '\\s*[-:]?\\s*', 'i'), '').trim();
    });

    var parts   = cleanEFN.split(' ').filter(function(p){return p.length>0;});
    var newCode = parts.map(function(p){return p.charAt(0).toUpperCase();}).join('') + id;
    sheet.getRange(i+1, codeCol+1).setValue(newCode);
    fixed++;
    Logger.log('Fixed: ' + efn + ' → ' + newCode);
  }
  Logger.log('Total fixed: ' + fixed);
}

// ============================================================
// MONTHLY TOKEN HEALTH CHECK
// ============================================================
function monthlyTokenCheck() {
  var props  = PropertiesService.getScriptProperties();
  var expiry = parseInt(props.getProperty('CONTACTS_TOKEN_EXPIRY') || '0');
  var token  = getValidAccessToken();

  if (!token) {
    MailApp.sendEmail({
      to: 'elias.j.massroua@gmail.com',
      subject: 'MGSDB Alert: Contacts Sync Token Expired',
      body: 'The Google Contacts sync token has expired and needs re-authorization.\n\n' +
            'Please open Apps Script and run getAuthorizationUrl() to fix.\n\n' +
            'MGSDB2026 Auto Monitor'
    });
    Logger.log('Token expired — email alert sent!');
  } else {
    Logger.log('Token OK — expires: ' + new Date(expiry).toLocaleString());
  }
}

function setupMonthlyTokenCheck() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'monthlyTokenCheck') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('monthlyTokenCheck')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
  Logger.log('Monthly token check set for 1st of every month at 9am');
}

function checkMissingEFN() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Employers');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var efnCol = headers.indexOf('EFN');
  var codeCol = headers.indexOf('EmployerCode');
  var idCol = headers.indexOf('EmployerID');
  var missing = 0;
  for(var i=1;i<data.length;i++){
    if(!data[i][efnCol]){
      Logger.log('Missing EFN: ID='+data[i][idCol]+' Code='+data[i][codeCol]);
      missing++;
    }
  }
  Logger.log('Total missing EFN: '+missing);
}
function fixInitialsEmployerCodes() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Employers');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol   = headers.indexOf('EmployerID');
  var codeCol = headers.indexOf('EmployerCode');
  var agCol   = headers.indexOf('AgencyCode');
  var fixed = 0;

  for (var i = 1; i < data.length; i++) {
    var code   = String(data[i][codeCol]||'').trim();
    var id     = data[i][idCol];
    var agCode = String(data[i][agCol]||'').trim().toUpperCase();
    if (!id || !agCode) continue;
    // Fix if code does NOT contain a dash (initials format has no dash)
    if (code && code.indexOf('-') === -1) {
      var newCode = agCode + '-' + id;
      sheet.getRange(i+1, codeCol+1).setValue(newCode);
      fixed++;
      Logger.log('Fixed: ' + code + ' → ' + newCode);
    }
  }
  Logger.log('Total fixed: ' + fixed);
}
