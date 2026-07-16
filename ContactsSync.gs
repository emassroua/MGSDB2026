// ============================================================
// MGS - Google Contacts Sync Module v5.1
// Syncs MGSC contacts from Personal Google Account → Employers Sheet
//                                                 → Firestore 'employers' collection (dual-write)
//
// FIX: No longer stores contacts in ScriptProperties (quota fix).
//      Uses API pageToken to resume between runs — zero storage issues.
//
// SETUP ORDER (do this once):
// 1. Run getAuthorizationUrl()   → open URL in browser, sign in with PERSONAL account
// 2. Copy "code=" from redirect URL → paste into saveAuthCode() → run it
// 3. Run syncGoogleContactsToDB() → starts sync
// 4. Run setupContactsSyncTrigger() → auto-runs every 15 min
//
// v5.1 — adds Firestore mirror via firestoreSaveDoc()
//        Requires FirestoreClient.gs + FirebaseSecrets.gs in same project.
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

function detectAgencyCodeFromPerson(person, agencyCodes) {
  if (person.names && person.names[0] && person.names[0].honorificPrefix) {
    var prefixRaw = String(person.names[0].honorificPrefix).trim().toUpperCase();
    if (prefixRaw) {
      var tags = prefixRaw.split(/[\s,;\-\/|]+/).filter(function(t) { return t.length > 0; });
      for (var i = 0; i < tags.length; i++) {
        if (agencyCodes.indexOf(tags[i]) !== -1) return tags[i];
      }
      return null;
    }
  }
  var displayName = (person.names && person.names[0] && person.names[0].displayName) ? person.names[0].displayName : '';
  return detectAgencyCode(displayName, agencyCodes);
}

var SYNC_SHEET      = 'Employers';
var SYNC_ID_FIELD   = 'GoogleContactID';
var SYNC_STATUS_COL = 'SyncStatus';
var SYNC_LOG_SHEET  = 'Sync_Log';
var PAGE_SIZE       = 200; // contacts per API page = contacts processed per run

var CONTACTS_CLIENT_ID     = '1071315446230-4hecmnn9f83pcm054epck06snv3mqr57.apps.googleusercontent.com';
var CONTACTS_CLIENT_SECRET = 'STORED_IN_APPS_SCRIPT_ONLY';
var CONTACTS_REDIRECT_URI = 'https://script.google.com/macros/d/1Ea1TUNZbk7T1OuzIrt4DGefEiS62Y2CgbkwitcwVo4QaYaMmYOLSU9rW/usercallback';
var CONTACTS_SCOPE         = 'https://www.googleapis.com/auth/contacts.readonly';

// ScriptProperties keys (small values only — no contact data stored!)
var PROP_SYNC_TOKEN  = 'CONTACTS_SYNC_TOKEN';     // saved after full sync completes
var PROP_PAGE_TOKEN  = 'CONTACTS_PAGE_TOKEN';     // current page position during full sync
var PROP_BATCH_STATS = 'BATCH_ACCUMULATED_STATS'; // running totals across runs
var PROP_FULL_DONE   = 'CONTACTS_FULL_SYNC_DONE'; // 'true' when full sync completed

// Firestore collection name (matches Firebase page's SHEET_MAP entry)
var FIRESTORE_EMPLOYERS_COLLECTION = 'employers';

// ============================================================
// FIRESTORE MIRROR HELPER
// ------------------------------------------------------------
// Writes an employer row to Firestore 'employers' collection.
// Wrapped in try/catch so Firestore failures never break Sheets writes.
// Uses EmployerID as the document ID (must be a string).
// ============================================================
function mirrorEmployerToFirestore(employerRow) {
  try {
    if (!employerRow || !employerRow.EmployerID) return;
    var docId = String(employerRow.EmployerID);
    // firestoreSaveDoc lives in FirestoreClient.gs
    var result = firestoreSaveDoc(FIRESTORE_EMPLOYERS_COLLECTION, docId, employerRow);
    if (!result || !result.success) {
      Logger.log('🔥 Firestore mirror failed for EmployerID ' + docId + ': ' + (result && result.error));
    }
  } catch (e) {
    Logger.log('🔥 Firestore mirror exception: ' + e.message);
  }
}

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
var AUTH_CODE = '4/0AfrIepC1VmYBadBnLnbs7cUqrcYFpET19BRlkS8R1GyKN0asYU6lzXX1A4ovtg9xsEN7yQ';
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
// FULL SYNC — fetches ONE page (200 contacts) per run
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
    return detectAgencyCodeFromPerson(p, _agencyCodes) !== null;
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

    if (response.getResponseCode() === 410 ||
        (response.getResponseCode() === 400 && response.getContentText().indexOf('EXPIRED_SYNC_TOKEN') !== -1)) {
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
        var _incAgencyCodes = getAgencyCodes();
        if (detectAgencyCodeFromPerson(person, _incAgencyCodes) !== null) changed.push(person);
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
          mirrorEmployerToFirestore(r.data); // ← Firestore mirror (dual-write)
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

      // Get ALL matching agency codes (multi-agency support)
      var allCodes = getAgencyCodes();
      var matchedAgencies = detectAllAgencyCodesFromPerson(person, allCodes);
      if (matchedAgencies.length === 0) return;  // no matching agency — skip

      // Process ONE record per matching agency
      matchedAgencies.forEach(function(agencyCode) {
        try {
          // Build employer row for this specific agency
          var employerRow = mapContactToEmployerForAgency(person, agencyCode);
          if (!employerRow) return;

          // Composite key: contactId + agencyCode (so each agency tracks its own copy)
          var mapKey = contactId + ':' + agencyCode;
          var existingRow = existingMap[mapKey];

          if (existingRow) {
            var merged = mergeEmployerData(existingRow.data, employerRow);
            actionSaveRow({ sheet: SYNC_SHEET, row: merged });
            mirrorEmployerToFirestore(merged);
            stats.updated++;
          } else {
            var nextId = actionGetNextID({ sheet: SYNC_SHEET });
            employerRow['EmployerID'] = nextId.nextID || 1;
            if (employerRow['AgencyCode']) {
              employerRow['EmployerCode'] = employerRow['AgencyCode'] + '-' + employerRow['EmployerID'];
            }
            actionSaveRow({ sheet: SYNC_SHEET, row: employerRow });
            existingMap[mapKey] = { data: employerRow };
            stats.inserted++;

            // ── Folder creation — UNTOUCHED, same logic as before
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

            mirrorEmployerToFirestore(employerRow);
          }
        } catch(ae) {
          Logger.log('Agency-loop error for ' + agencyCode + ': ' + ae.message);
          stats.errors++;
        }
      });
    } catch(e) {
      Logger.log('Contact error: ' + e.message);
      stats.errors++;
    }
  });
}

// ============================================================
// MAP GOOGLE CONTACT → EMPLOYER ROW
// ============================================================
// ============================================================
// MULTI-AGENCY: Returns ALL agency codes matching the prefix tags
// ============================================================
function detectAllAgencyCodesFromPerson(person, agencyCodes) {
  if (!person.names || !person.names[0] || !person.names[0].honorificPrefix) {
    return [];
  }
  var prefixRaw = String(person.names[0].honorificPrefix).trim().toUpperCase();
  if (!prefixRaw) return [];
  var tags = prefixRaw.split(/[\s,;\-\/|]+/).filter(function(t) { return t.length > 0; });
  var matches = [];
  for (var i = 0; i < tags.length; i++) {
    if (agencyCodes.indexOf(tags[i]) !== -1 && matches.indexOf(tags[i]) === -1) {
      matches.push(tags[i]);
    }
  }
  return matches;
}
function mapContactToEmployer(person) {
  var row = {};
  row[SYNC_ID_FIELD] = person.resourceName || '';

  if (person.names && person.names[0]) {
    var fullName        = person.names[0].displayName || '';
    var _codes = getAgencyCodes();
    var _detected = detectAgencyCodeFromPerson(person, _codes);
    if (!_detected) return null;  // skip contacts without a recognized agency prefix
    row['EFN']      = fullName.replace(new RegExp('^' + _detected + '\\s*[-:]?\\s*', 'i'), '').trim();
    row['AgencyCode'] = _detected;
    // ── Auto-generate EmployerCode — simple format: AgencyCode + "-" + EmployerID
    // e.g. MGSC + "-" + 1355 = MGSC-1355
    // (EmployerID is assigned later in processContacts for new inserts; for updates
    //  the existing code is preserved via mergeEmployerData PROTECTED logic.)
    if (!row['EmployerCode'] && row['AgencyCode'] && row['EmployerID']) {
      row['EmployerCode'] = row['AgencyCode'] + '-' + row['EmployerID'];
    }
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
    row['EmployerName'] = row['EFN'];
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
    /// fallback — if still empty use first number for Mobile only
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
  row['LastSyncDate']  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
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

  var headers      = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
  var idColIdx     = headers.indexOf(SYNC_ID_FIELD);
  var agencyColIdx = headers.indexOf('AgencyCode');
  if (idColIdx < 0) return map;

  sheet.getRange(2, 1, lastRow - 1, lastCol).getValues().forEach(function(row, i) {
    var cid = String(row[idColIdx] || '').trim();
    if (!cid) return;
    var agencyCode = agencyColIdx >= 0 ? String(row[agencyColIdx] || '').trim().toUpperCase() : '';
    var obj = {};
    headers.forEach(function(h, j) { if (h) obj[h] = row[j] == null ? '' : row[j]; });
    
    // New composite key: contactId:agencyCode
    var compositeKey = agencyCode ? cid + ':' + agencyCode : cid;
    map[compositeKey] = { data: obj, rowIndex: i + 2 };
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
      mirrorEmployerToFirestore(emp); // ← also mirror the agency link fix
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

// ============================================================
// FIND & SYNC SINGLE CONTACT BY NAME
// Usage: findAndSyncContact('Nada') or findAndSyncContact('MGSC Elias')
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
  var NAME = 'PutTheContactNameHere';  // ← change this
  findAndSyncContact(NAME);
}

function fixMissingEmployerCodes() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName('Employers');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol   = headers.indexOf('EmployerID');
  var codeCol = headers.indexOf('EmployerCode');
  var agCol   = headers.indexOf('AgencyCode');
  var fixed   = 0;

  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][codeCol]||'').trim();
    if (code) continue; // already has code
    var id     = data[i][idCol];
    var agCode = String(data[i][agCol]||'').trim().toUpperCase();
    if (!id || !agCode) continue;

    // New format: AgencyCode-EmployerID (e.g. MGSC-1355)
    var newCode = agCode + '-' + id;
    sheet.getRange(i+1, codeCol+1).setValue(newCode);
    fixed++;
    Logger.log('Fixed: ID ' + id + ' → ' + newCode);
  }
  Logger.log('Total fixed: ' + fixed);
}

// ============================================================
// FIRESTORE MIRROR — TEST HELPER
// ------------------------------------------------------------
// Run this manually to test that the Firestore dual-write works
// before running a real sync. Picks one existing employer and
// re-writes it to Firestore.
// ============================================================
function testFirestoreMirror() {
  Logger.log('── Firestore Mirror Test ──');

  // Get one employer from Sheets
  var employers = actionGetAll({ sheet: 'Employers' });
  if (!employers || !employers.data || employers.data.length === 0) {
    Logger.log('❌ No employers in Sheets to test with.');
    return;
  }

  var emp = employers.data[0];
  Logger.log('Test employer: ID=' + emp.EmployerID + ' | Name=' + (emp.EFN || emp.EmployerName));

  // Mirror to Firestore
  var before = new Date();
  mirrorEmployerToFirestore(emp);
  var after  = new Date();

  Logger.log('✅ Mirror call completed in ' + (after - before) + 'ms');

  // Read it back from Firestore to verify
  var readBack = firestoreGetDoc(FIRESTORE_EMPLOYERS_COLLECTION, String(emp.EmployerID));
  if (readBack) {
    Logger.log('✅ Read-back verified from Firestore:');
    Logger.log('   EmployerID: ' + readBack.EmployerID);
    Logger.log('   EFN:        ' + readBack.EFN);
    Logger.log('   AgencyCode: ' + readBack.AgencyCode);
  } else {
    Logger.log('❌ Read-back failed — document not found in Firestore');
  }
}

function debugContactPrefixes() {
  var token = getValidAccessToken();
  var codes = getAgencyCodes();
  Logger.log('Agencies: ' + codes.join(', '));
  Logger.log('---');

  var url = 'https://people.googleapis.com/v1/people/me/connections?personFields=names&pageSize=50';
  var resp = UrlFetchApp.fetch(url, { headers:{Authorization:'Bearer '+token}, muteHttpExceptions:true });
  var data = JSON.parse(resp.getContentText());

  (data.connections||[]).forEach(function(p){
    var name = (p.names && p.names[0]) ? p.names[0].displayName : '(no name)';
    var detected = detectAgencyCode(name, codes);
    Logger.log((detected ? '✅ '+detected : '❌ NULL') + ' | ' + name);
  });
}

function testSyncAfterFix() {
  var result = syncGoogleContactsToDB();
  Logger.log(JSON.stringify(result));
}

function debugFirstPage() {
  var token = getValidAccessToken();
  var codes = getAgencyCodes();
  Logger.log('Agency codes: ' + codes.join(', '));
  Logger.log('---');

  var url = 'https://people.googleapis.com/v1/people/me/connections'
    + '?personFields=names&pageSize=200';
  var resp = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + getValidAccessToken() },
    muteHttpExceptions: true
  });
  var data = JSON.parse(resp.getContentText());
  
  var matched = 0;
  var skipped = 0;
  (data.connections || []).forEach(function(p) {
    var name = (p.names && p.names[0]) ? p.names[0].displayName : '';
    var detected = detectAgencyCode(name, codes);
    if (detected) {
      matched++;
      if (matched <= 5) Logger.log('✅ MATCH: ' + detected + ' | ' + name);
    } else {
      skipped++;
      if (skipped <= 5) Logger.log('❌ SKIP: ' + name);
    }
  });
  Logger.log('---');
  Logger.log('Total: ' + (data.connections || []).length);
  Logger.log('Matched: ' + matched);
  Logger.log('Skipped: ' + skipped);
}

function debugLocalAgencies() {
  var result = actionGetAll({ sheet: 'Local_Agencies' });
  if (!result || !result.data) { Logger.log('No data'); return; }
  Logger.log('Total agencies: ' + result.data.length);
  Logger.log('---');
  result.data.forEach(function(a) {
    Logger.log(
      'Code: "' + (a.AgencyCode || '(empty)') + '" | ' +
      'Name: ' + (a.AgencyName || '(empty)') + ' | ' +
      'Status: "' + (a.AgencyStatus || '(empty)') + '"'
    );
  });
}

function debugAllAgencyData() {
  // 1. Read from Sheets (what ContactSync uses)
  var sheetResult = actionGetAll({ sheet: 'Local_Agencies' });
  Logger.log('=== FROM GOOGLE SHEETS ===');
  Logger.log('Count: ' + (sheetResult.data ? sheetResult.data.length : 0));
  (sheetResult.data || []).forEach(function(a) {
    Logger.log('  Code: ' + a.AgencyCode + ' | Name: ' + a.AgencyName + ' | Status: ' + (a.AgencyStatus || '(empty)'));
  });
  
  // 2. Read from Firestore (what Local_Agencies_Manager UI uses)
  Logger.log('=== FROM FIRESTORE ===');
  try {
    var fsAgencies = firestoreGetAll('Local_Agencies');
    Logger.log('Count: ' + (fsAgencies ? fsAgencies.length : 0));
    (fsAgencies || []).forEach(function(a) {
      Logger.log('  Code: ' + a.AgencyCode + ' | Name: ' + a.AgencyName + ' | Status: ' + (a.AgencyStatus || '(empty)'));
    });
  } catch(e) {
    Logger.log('Firestore read error: ' + e.message);
  }
}

// ============================================================
// DEBUG: Scan all contacts and group by prefix (first word of name)
// Shows which tags exist in Google Contacts and how many contacts use each
// ============================================================
function debugAllContactPrefixes() {
  var token = getValidAccessToken();
  if (!token) { Logger.log('Not authorized.'); return; }

  var prefixCount = {};   // { 'MGSS': 47, 'NCS': 23, ... }
  var noPrefix = 0;        // contacts that start with a non-letter or have no name
  var totalContacts = 0;
  var pageToken = null;
  var pageNum = 0;

  do {
    pageNum++;
    var url = 'https://people.googleapis.com/v1/people/me/connections'
      + '?personFields=names&pageSize=1000'
      + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    var resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('API error on page ' + pageNum + ': ' + resp.getContentText());
      break;
    }

    var data = JSON.parse(resp.getContentText());
    (data.connections || []).forEach(function(p) {
      totalContacts++;
      var name = (p.names && p.names[0]) ? (p.names[0].displayName || '').trim() : '';
      if (!name) { noPrefix++; return; }

      // Extract the first "word" — letters/digits before space, dash, colon, or punctuation
      var match = name.match(/^([A-Z0-9]+)/i);
      if (!match) { noPrefix++; return; }

      var prefix = match[1].toUpperCase();
      prefixCount[prefix] = (prefixCount[prefix] || 0) + 1;
    });

    pageToken = data.nextPageToken || null;
  } while (pageToken);

  // Sort prefixes by count (descending)
  var sorted = Object.keys(prefixCount).sort(function(a, b) {
    return prefixCount[b] - prefixCount[a];
  });

  // Compare against active agency codes
  var agencyCodes = getAgencyCodes();
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('GOOGLE CONTACTS PREFIX SCAN');
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('Total contacts scanned: ' + totalContacts);
  Logger.log('Contacts with a prefix: ' + (totalContacts - noPrefix));
  Logger.log('Contacts without prefix: ' + noPrefix);
  Logger.log('Unique prefixes found: ' + sorted.length);
  Logger.log('Active agency codes in sheet: ' + agencyCodes.join(', '));
  Logger.log('───────────────────────────────────────────────');
  Logger.log('PREFIXES (sorted by count):');
  Logger.log('───────────────────────────────────────────────');

  sorted.forEach(function(prefix) {
    var inAgencies = agencyCodes.indexOf(prefix) !== -1;
    var marker = inAgencies ? '✅ SYNC' : '❌ SKIP';
    Logger.log(marker + ' | ' + prefix + ' (' + prefixCount[prefix] + ' contacts)');
  });
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('Pages scanned: ' + pageNum);
}

// ============================================================
// DEBUG: Scan only the "Prefix" field of Google Contacts
// (honorificPrefix — the actual Name Prefix dropdown field)
// ============================================================
function debugContactPrefixField() {
  var token = getValidAccessToken();
  if (!token) { Logger.log('Not authorized.'); return; }

  var prefixCount = {};
  var noPrefix = 0;
  var totalContacts = 0;
  var pageToken = null;
  var pageNum = 0;

  do {
    pageNum++;
    var url = 'https://people.googleapis.com/v1/people/me/connections'
      + '?personFields=names&pageSize=1000'
      + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    var resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('API error on page ' + pageNum + ': ' + resp.getContentText());
      break;
    }

    var data = JSON.parse(resp.getContentText());
    (data.connections || []).forEach(function(p) {
      totalContacts++;
      var prefix = '';
      if (p.names && p.names[0] && p.names[0].honorificPrefix) {
        prefix = String(p.names[0].honorificPrefix).trim().toUpperCase();
      }
      if (!prefix) {
        noPrefix++;
        return;
      }
      prefixCount[prefix] = (prefixCount[prefix] || 0) + 1;
    });

    pageToken = data.nextPageToken || null;
  } while (pageToken);

  var sorted = Object.keys(prefixCount).sort(function(a, b) {
    return prefixCount[b] - prefixCount[a];
  });

  var agencyCodes = getAgencyCodes();
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('GOOGLE CONTACTS — PREFIX FIELD ONLY');
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('Total contacts scanned: ' + totalContacts);
  Logger.log('Contacts with prefix field: ' + (totalContacts - noPrefix));
  Logger.log('Contacts WITHOUT prefix: ' + noPrefix);
  Logger.log('Unique prefixes found: ' + sorted.length);
  Logger.log('Active agency codes in sheet: ' + agencyCodes.join(', '));
  Logger.log('───────────────────────────────────────────────');
  sorted.forEach(function(prefix) {
    var inAgencies = agencyCodes.indexOf(prefix) !== -1;
    var marker = inAgencies ? '✅ SYNC' : '❌ SKIP';
    Logger.log(marker + ' | ' + prefix + ' (' + prefixCount[prefix] + ' contacts)');
  });
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('Pages scanned: ' + pageNum);
}

function testDetect() {
  var token = getValidAccessToken();
  var url = 'https://people.googleapis.com/v1/people/me/connections?personFields=names&pageSize=10';
  var resp = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
  var data = JSON.parse(resp.getContentText());
  var codes = getAgencyCodes();
  
  (data.connections || []).forEach(function(p) {
    var name = (p.names && p.names[0] && p.names[0].displayName) ? p.names[0].displayName : '(no name)';
    var prefix = (p.names && p.names[0] && p.names[0].honorificPrefix) ? p.names[0].honorificPrefix : '(empty)';
    var detected = detectAgencyCodeFromPerson(p, codes);
    Logger.log(name + ' | prefix: "' + prefix + '" | detected: ' + (detected || 'NULL'));
  });
}

function detectAgencyCode(contactName, agencyCodes) {
  if (!contactName) return null;
  var name = contactName.trim().toUpperCase();
  for (var i = 0; i < agencyCodes.length; i++) {
    if (name.indexOf(agencyCodes[i]) === 0) return agencyCodes[i];
  }
  return null;
}

function checkResult() {
  var result = syncGoogleContactsToDB();
  Logger.log(JSON.stringify(result));
}

// ============================================================
// MAP GOOGLE CONTACT → EMPLOYER ROW (for a specific agency)
// Used by multi-agency processContacts loop
// ============================================================
function mapContactToEmployerForAgency(person, agencyCode) {
  var row = {};
  row[SYNC_ID_FIELD] = person.resourceName || '';

  if (person.names && person.names[0]) {
    var fullName = person.names[0].displayName || '';
    row['EFN'] = fullName.replace(new RegExp('^' + agencyCode + '\\s*[-:,]?\\s*', 'i'), '').trim();
    row['AgencyCode'] = agencyCode;
    
    // Look up AgencyID from Local_Agencies (Firestore)
    try {
      var agencyDocs = firestoreListCollection('Local_Agencies');
      if (agencyDocs && agencyDocs.length > 0) {
        var matchedAgency = agencyDocs.find(function(a) {
          return a.AgencyCode && String(a.AgencyCode).trim().toUpperCase() === agencyCode;
        });
        if (matchedAgency) {
          row['AgencyID'] = matchedAgency.AgencyID;
        }
      }
    } catch(ae) {
      Logger.log('Agency lookup error: ' + ae.message);
    }
    row['EmployerName'] = row['EFN'];
  }

  if (person.phoneNumbers && person.phoneNumbers.length > 0) {
    person.phoneNumbers.forEach(function(ph) {
      var type = (ph.type || ph.canonicalForm || '').toLowerCase();
      var val = ph.value || '';
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
    var month = String(bd.month).padStart(2, '0');
    var day = String(bd.day).padStart(2, '0');
    var year = bd.year ? bd.year : 1900;
    row['BirthDate'] = year + '-' + month + '-' + day;
  }
  if (person.emailAddresses && person.emailAddresses[0]) {
    row['EmailAddress'] = person.emailAddresses[0].value || '';
  }
  if (person.addresses && person.addresses[0]) {
    var addr = person.addresses[0];
    row['Address'] = addr.streetAddress || '';
    row['City'] = addr.city || '';
    row['Country'] = addr.country || '';
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
function batchSyncRunner() {
  var startTime = new Date().getTime();
  var maxRunMs = 5 * 60 * 1000;
  for (var i = 0; i < 30; i++) {
    if (new Date().getTime() - startTime > maxRunMs) {
      Logger.log('Time limit reached after ' + i + ' batches.');
      return;
    }
    var result = syncGoogleContactsToDB();
    Logger.log('Batch ' + (i+1) + ': mode=' + result.mode + ' inProgress=' + result.inProgress + ' inserted=' + result.inserted + ' updated=' + result.updated);
    if (!result.inProgress) {
      Logger.log('Full sync complete!');
      return;
    }
  }
}
function checkAutoSyncStatus() {
  var props = PropertiesService.getScriptProperties();
  var pageToken = props.getProperty('CONTACTS_PAGE_TOKEN');
  var syncToken = props.getProperty('CONTACTS_SYNC_TOKEN');
  var statsRaw  = props.getProperty('BATCH_ACCUMULATED_STATS');
  Logger.log('Page token: ' + (pageToken ? 'in progress' : 'done'));
  Logger.log('Sync token: ' + (syncToken ? 'have one' : 'none'));
  Logger.log('Stats: ' + statsRaw);
}
function debugMultiAgency() {
  var token = getValidAccessToken();
  // Find your test contacts by name
  var url = 'https://people.googleapis.com/v1/people/me/connections?personFields=names&pageSize=1000';
  var pageToken = null;
  var found = [];
  do {
    var u = url + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    var resp = UrlFetchApp.fetch(u, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());
    (data.connections || []).forEach(function(p) {
      var prefix = (p.names && p.names[0] && p.names[0].honorificPrefix) || '';
      if (/MGSC.*SALI|SALI.*MGSC/i.test(prefix)) {
        var name = (p.names[0].displayName || '');
        found.push({ name: name, prefix: prefix });
      }
    });
    pageToken = data.nextPageToken || null;
  } while (pageToken);
  
  Logger.log('Found ' + found.length + ' multi-tag contacts:');
  found.forEach(function(c) {
    Logger.log('  Prefix: "' + c.prefix + '" | Name: ' + c.name);
    var agencies = detectAllAgencyCodesFromPerson({names:[{honorificPrefix:c.prefix}]}, ['MGSC','SALI']);
    Logger.log('  → Detected agencies: ' + JSON.stringify(agencies));
  });
}

function testMultiAgencyOnTestContacts() {
  var token = getValidAccessToken();
  var url = 'https://people.googleapis.com/v1/people/me/connections?personFields=names,emailAddresses,phoneNumbers,addresses,organizations,biographies,userDefined,birthdays&pageSize=1000';
  var pageToken = null;
  var testContacts = [];
  
  do {
    var u = url + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    var resp = UrlFetchApp.fetch(u, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());
    (data.connections || []).forEach(function(p) {
      var prefix = (p.names && p.names[0] && p.names[0].honorificPrefix) || '';
      if (/MGSC.*SALI|SALI.*MGSC/i.test(prefix)) {
        testContacts.push(p);
      }
    });
    pageToken = data.nextPageToken || null;
  } while (pageToken);
  
  Logger.log('Found ' + testContacts.length + ' test contacts to process');
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SYNC_SHEET);
  var existingMap = buildExistingEmployerMap(sheet);
  var stats = { inserted: 0, updated: 0, inactive: 0, folders: 0, errors: 0, total: 0 };
  
  processContacts(testContacts, existingMap, stats);
  
  Logger.log('=== RESULT ===');
  Logger.log('Inserted: ' + stats.inserted);
  Logger.log('Updated: ' + stats.updated);
  Logger.log('Folders: ' + stats.folders);
  Logger.log('Errors: ' + stats.errors);
}
// ============================================================
// CHECK CONTACTS EXIST FOR AGENCY (used by auto-sync popup)
// ------------------------------------------------------------
// Read-only. Scans Google Contacts for the given agency code in
// the Prefix field. Short-circuits at the first match for speed.
// Used by local_agencies_manager.html to decide whether to show
// the "sync now?" popup after creating a new agency.
// ============================================================
function hasContactsForAgency(agencyCode) {
  try {
    if (!agencyCode) return { success: false, error: 'No agency code provided' };
    var target = String(agencyCode).trim().toUpperCase();
    if (!target) return { success: false, error: 'Empty agency code' };

    var token = getValidAccessToken();
    if (!token) return { success: false, error: 'Not authorized' };

    var pageToken = null;
    do {
      var url = 'https://people.googleapis.com/v1/people/me/connections'
        + '?personFields=names&pageSize=1000'
        + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

      var resp = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      if (resp.getResponseCode() !== 200) {
        return { success: false, error: 'People API error: ' + resp.getResponseCode() };
      }

      var data = JSON.parse(resp.getContentText());
      var connections = data.connections || [];

      for (var i = 0; i < connections.length; i++) {
        var p = connections[i];
        if (!p.names || !p.names[0] || !p.names[0].honorificPrefix) continue;
        var prefixRaw = String(p.names[0].honorificPrefix).trim().toUpperCase();
        if (!prefixRaw) continue;
        var tags = prefixRaw.split(/[\s,;\-\/|]+/).filter(function(t) { return t.length > 0; });
        if (tags.indexOf(target) !== -1) {
          return { success: true, hasContacts: true };
        }
      }

      pageToken = data.nextPageToken || null;
    } while (pageToken);

    return { success: true, hasContacts: false };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function actionHasContactsForAgency(params) {
  return hasContactsForAgency(params && params.agencyCode);
}
function listAllScriptProperties() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var keys = Object.keys(props);
  Logger.log('Total properties stored: ' + keys.length);
  Logger.log('───────────────────────────────────');
  keys.forEach(function(k) {
    var v = props[k] || '';
    var size = v.length;
    var preview = size > 80 ? v.substring(0, 80) + '... [TRUNCATED]' : v;
    Logger.log('[' + size + ' chars] ' + k + ' = ' + preview);
  });
}

function cleanupOldScriptProperties() {
  var props = PropertiesService.getScriptProperties();
  var toDelete = [
    'BATCH_CONTACTS_0', 'BATCH_CONTACTS_1', 'BATCH_CONTACTS_2',
    'BATCH_CONTACTS_3', 'BATCH_CONTACTS_4', 'BATCH_CONTACTS_5',
    'AS_PROGRESS_MGSC_STATS', 'AS_PROGRESS_MGSC_QUEUE',
    'AS_PROGRESS_MGSC_POS',   'AS_PROGRESS_MGSC_TOTAL',
    'TRIGGER_BATCH_INDEX'
  ];
  var deleted = 0, notFound = 0;
  toDelete.forEach(function(k) {
    if (props.getProperty(k) !== null) {
      props.deleteProperty(k);
      deleted++;
      Logger.log('✅ Deleted: ' + k);
    } else {
      notFound++;
      Logger.log('— Not present: ' + k);
    }
  });
  Logger.log('───────────────────────────────────');
  Logger.log('Cleanup complete. Deleted: ' + deleted + ' | Not found: ' + notFound);
}
// ============================================================
// DAILY JANITOR — auto-cleanup stale ScriptProperties
// ------------------------------------------------------------
// Runs once a day via time-based trigger. Cleans up:
//   1. AS_PROGRESS_* keys where POS >= TOTAL (sync finished but
//      _clearProgress didn't run for some reason)
//   2. Any legacy BATCH_CONTACTS_* keys (pre-v5.1 leftovers)
//   3. TRIGGER_BATCH_INDEX (legacy)
// 
// Read-only on important keys: NEVER touches OAuth tokens,
// Twilio config, MGS_ROOT_FOLDER_ID, or any current sync state.
// ============================================================
function dailyJanitorCleanup() {
  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  var keys = Object.keys(allProps);
  var deleted = [];

  keys.forEach(function(k) {
    // 1. Legacy BATCH_CONTACTS_* keys — always safe to delete
    if (/^BATCH_CONTACTS_\d+$/.test(k)) {
      props.deleteProperty(k);
      deleted.push(k + ' (legacy batch)');
      return;
    }

    // 2. TRIGGER_BATCH_INDEX legacy
    if (k === 'TRIGGER_BATCH_INDEX') {
      props.deleteProperty(k);
      deleted.push(k + ' (legacy)');
      return;
    }

    // 3. AS_PROGRESS_<CODE>_POS — only delete the GROUP if POS >= TOTAL
    var posMatch = k.match(/^AS_PROGRESS_(.+)_POS$/);
    if (posMatch) {
      var code = posMatch[1];
      var pos = parseInt(allProps[k] || '0');
      var total = parseInt(allProps['AS_PROGRESS_' + code + '_TOTAL'] || '0');
      // If sync finished (pos >= total > 0), clean the whole group
      if (total > 0 && pos >= total) {
        ['QUEUE', 'POS', 'STATS', 'TOTAL'].forEach(function(suffix) {
          var fullKey = 'AS_PROGRESS_' + code + '_' + suffix;
          if (props.getProperty(fullKey) !== null) {
            props.deleteProperty(fullKey);
            deleted.push(fullKey + ' (completed sync)');
          }
        });
      }
    }
  });

  Logger.log('🧹 Daily janitor — deleted ' + deleted.length + ' stale keys');
  deleted.forEach(function(k) { Logger.log('  • ' + k); });
  return { success: true, deletedCount: deleted.length, deletedKeys: deleted };
}

// Setup function — run ONCE manually to install the daily trigger
function setupDailyJanitorTrigger() {
  // Remove any existing janitor triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyJanitorCleanup') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Install fresh: runs once a day around 3 AM (low-traffic time)
  ScriptApp.newTrigger('dailyJanitorCleanup')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
  Logger.log('✅ Daily janitor trigger installed (runs ~3 AM daily)');
  return { success: true };
}
function debugCheckLahoud() {
  var token = getValidAccessToken();
  var url = 'https://people.googleapis.com/v1/people/me/connections?personFields=names&pageSize=1000';
  var pageToken = null;
  var found = false;
  
  do {
    var u = url + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    var resp = UrlFetchApp.fetch(u, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());
    (data.connections || []).forEach(function(p) {
      var name = (p.names && p.names[0]) ? p.names[0].displayName : '';
      if (name.toLowerCase().indexOf('lahoud louis') !== -1) {
        found = true;
        Logger.log('=== FOUND LAHOUD ===');
        Logger.log('displayName: "' + p.names[0].displayName + '"');
        Logger.log('honorificPrefix: "' + (p.names[0].honorificPrefix || '(EMPTY)') + '"');
        Logger.log('givenName: "' + (p.names[0].givenName || '') + '"');
        Logger.log('familyName: "' + (p.names[0].familyName || '') + '"');
        Logger.log('Full names object: ' + JSON.stringify(p.names));
      }
    });
    pageToken = data.nextPageToken || null;
  } while (pageToken);
  
  if (!found) Logger.log('Lahoud not found in Google Contacts!');
}

function debugSyncLahoud() {
  var token = getValidAccessToken();
  var url = 'https://people.googleapis.com/v1/people/me/connections?personFields=names,phoneNumbers,emailAddresses,addresses,organizations,biographies,birthdays,userDefined&pageSize=1000';
  var pageToken = null;
  var lahoud = null;
  
  do {
    var u = url + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    var resp = UrlFetchApp.fetch(u, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());
    (data.connections || []).forEach(function(p) {
      var name = (p.names && p.names[0]) ? p.names[0].displayName : '';
      if (name.toLowerCase().indexOf('lahoud louis') !== -1) lahoud = p;
    });
    pageToken = data.nextPageToken || null;
  } while (pageToken);
  
  if (!lahoud) { Logger.log('Lahoud not found!'); return; }
  
  Logger.log('Found Lahoud. Testing detection...');
  var codes = getAgencyCodes();
  Logger.log('Agency codes loaded: ' + JSON.stringify(codes));
  
  var matched = detectAllAgencyCodesFromPerson(lahoud, codes);
  Logger.log('Matched agencies: ' + JSON.stringify(matched));
  
  if (matched.length === 0) { Logger.log('NO MATCH — sync would skip him'); return; }
  
  Logger.log('Trying mapContactToEmployerForAgency for ' + matched[0]);
  try {
    var row = mapContactToEmployerForAgency(lahoud, matched[0]);
    Logger.log('Mapped row: ' + JSON.stringify(row));
  } catch(e) {
    Logger.log('MAP FAILED: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
}

function debugRunFullSync() {
  var props = PropertiesService.getScriptProperties();
  // Clear sync token to force full sync
  props.deleteProperty('PEOPLE_SYNC_TOKEN');
  props.deleteProperty('SYNC_IN_PROGRESS');
  props.deleteProperty('SYNC_PAGE_TOKEN');
  Logger.log('Cleared sync state. Now running full sync...');
  syncGoogleContactsToDB();
  Logger.log('Done.');
}
function debugWhyLahoudSkipped() {
  // Manually replay processContacts on just Lahoud
  var token = getValidAccessToken();
  var url = 'https://people.googleapis.com/v1/people/me/connections?personFields=names,phoneNumbers,emailAddresses,addresses,organizations,biographies,birthdays,userDefined,metadata&pageSize=1000';
  var pageToken = null;
  var lahoud = null;
  
  do {
    var u = url + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    var resp = UrlFetchApp.fetch(u, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());
    (data.connections || []).forEach(function(p) {
      var name = (p.names && p.names[0]) ? p.names[0].displayName : '';
      if (name.toLowerCase().indexOf('lahoud louis') !== -1) lahoud = p;
    });
    pageToken = data.nextPageToken || null;
  } while (pageToken);
  
  if (!lahoud) { Logger.log('Lahoud not found in API!'); return; }
  Logger.log('Lahoud resourceName: ' + lahoud.resourceName);
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Employers');
  var existingMap = buildExistingEmployerMap(sheet);
  Logger.log('existingMap keys count: ' + Object.keys(existingMap).length);
  
  // Check if Lahoud is already in existingMap
  var hasLahoud = false;
  Object.keys(existingMap).forEach(function(k) {
    if (k.indexOf(lahoud.resourceName) !== -1) {
      hasLahoud = true;
      Logger.log('Lahoud already in map with key: ' + k);
      Logger.log('Existing data: ' + JSON.stringify(existingMap[k].data));
    }
  });
  if (!hasLahoud) Logger.log('Lahoud NOT in existingMap — should be inserted');
  
  var stats = { inserted: 0, updated: 0, inactive: 0, folders: 0, errors: 0, total: 1 };
  Logger.log('Calling processContacts...');
  processContacts([lahoud], existingMap, stats);
  Logger.log('Stats after: ' + JSON.stringify(stats));
}
// ============================================================
// FRONTEND INSTANT SYNC — search and sync separate functions
// Used by Employers Manager "⚡ Sync Contact" button
// ============================================================

function searchContactsByName(params) {
  try {
    var searchName = params && params.name ? String(params.name).trim() : '';
    if (!searchName) return { success: false, error: 'Name required' };

    var token = getValidAccessToken();
    if (!token) return { success: false, error: 'Not authorized — run getAuthorizationUrl() first' };

    var found = [];
    var pageToken = null;
    var pageCount = 0;
    var maxPages = 10;

    do {
      var url = 'https://people.googleapis.com/v1/people/me/connections'
        + '?personFields=names,phoneNumbers,emailAddresses,organizations'
        + '&pageSize=1000'
        + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

      var resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });

      if (resp.getResponseCode() !== 200) {
        return { success: false, error: 'People API error: ' + resp.getContentText().substring(0, 200) };
      }

      var data = JSON.parse(resp.getContentText());

      (data.connections || []).forEach(function(p) {
        var name = (p.names && p.names[0]) ? p.names[0].displayName : '';
        if (name && name.toLowerCase().indexOf(searchName.toLowerCase()) !== -1) {
          var phone = (p.phoneNumbers && p.phoneNumbers[0]) ? p.phoneNumbers[0].value : '';
          found.push({
            name: name,
            phone: phone,
            resourceName: p.resourceName
          });
        }
      });

      pageToken = data.nextPageToken || null;
      pageCount++;
    } while (pageToken && pageCount < maxPages);

    return { success: true, contacts: found };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function syncContactByResourceName(params) {
  try {
    var resourceName = params && params.resourceName ? String(params.resourceName) : '';
    if (!resourceName) return { success: false, error: 'resourceName required' };

    var token = getValidAccessToken();
    if (!token) return { success: false, error: 'Not authorized' };

    // Fetch the single contact with full details
    var url = 'https://people.googleapis.com/v1/' + resourceName
      + '?personFields=names,phoneNumbers,emailAddresses,addresses,organizations,biographies,birthdays,userDefined';

    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      return { success: false, error: 'Failed to fetch contact (HTTP ' + resp.getResponseCode() + ')' };
    }

    var contact = JSON.parse(resp.getContentText());

    // Use existing processContacts() to write to Sheet + Firestore
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SYNC_SHEET) || ss.insertSheet(SYNC_SHEET);
    var existingMap = buildExistingEmployerMap(sheet);
    var stats = { inserted: 0, updated: 0, inactive: 0, folders: 0, errors: 0, total: 0 };

    processContacts([contact], existingMap, stats);

    return {
      success:  true,
      inserted: stats.inserted,
      updated:  stats.updated,
      folders:  stats.folders,
      errors:   stats.errors
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
// ============================================================
// actionCreateFolder — creates folder in Firebase Storage
// Replaces old Drive-based folder creation
// Folder is created by uploading a tiny .placeholder file
// at: {parentType}/{folderName}/.placeholder
// ============================================================
function actionCreateFolder(params) {
  try {
    var folderName = (params.folderName || '').trim();
    var parentType = (params.parentType || 'employers').trim();
    if (!folderName) return { success: false, error: 'Missing folderName' };

    // Get a fresh OAuth token using service account JWT (already cached for 50 min)
    var token = getFirestoreAccessToken();
    if (!token) return { success: false, error: 'Could not get service account token' };

    // Firebase Storage REST API endpoint
    var bucket = 'mgsd2026.firebasestorage.app';
    var storagePath = parentType + '/' + folderName + '/.placeholder';
    var encodedPath = encodeURIComponent(storagePath);

    var url = 'https://firebasestorage.googleapis.com/v0/b/' + bucket
            + '/o?uploadType=media&name=' + encodedPath;

    // Upload a tiny placeholder file
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'text/plain'
      },
      payload: 'placeholder',
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code !== 200) {
      Logger.log('Storage upload error: ' + code + ' — ' + resp.getContentText());
      return { success: false, error: 'HTTP ' + code };
    }

    // Build a public viewer URL (Firebase Console URL)
    var consoleUrl = 'https://console.firebase.google.com/u/0/project/' + FIREBASE_SECRETS.projectId
                   + '/storage/' + bucket
                   + '/files/~2F' + parentType + '~2F' + encodeURIComponent(folderName).replace(/%20/g, '%20');

    return {
      success: true,
      folderId: storagePath,
      folderName: folderName,
      folderUrl: consoleUrl,
      storagePath: parentType + '/' + folderName
    };
  } catch(e) {
    Logger.log('actionCreateFolder error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// getFirestoreAccessToken — reuses the JWT flow from FirestoreClient.gs
// If your project already has this exposed, this wrapper is harmless.
// ============================================================
function getFirestoreAccessToken() {
  // FirestoreClient.gs typically defines a private _getAccessToken().
  // We try a couple of common names; one of these should work.
  if (typeof _firestoreGetAccessToken === 'function') return _firestoreGetAccessToken();
  if (typeof firestoreGetAccessToken === 'function')  return firestoreGetAccessToken();
  if (typeof getFirestoreToken === 'function')        return getFirestoreToken();

  // Fallback: build it ourselves using the existing FIREBASE_SECRETS
  return _buildServiceAccountToken();
}

function _buildServiceAccountToken() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('FB_STORAGE_TOKEN');
  if (cached) return cached;

  var iat = Math.floor(Date.now() / 1000);
  var exp = iat + 3000; // 50 min
  var jwtHeader  = Utilities.base64EncodeWebSafe(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).replace(/=+$/, '');
  var jwtClaim   = Utilities.base64EncodeWebSafe(JSON.stringify({
    iss:   FIREBASE_SECRETS.clientEmail,
    scope: 'https://www.googleapis.com/auth/devstorage.read_write https://www.googleapis.com/auth/cloud-platform',
    aud:   'https://oauth2.googleapis.com/token',
    iat:   iat,
    exp:   exp
  })).replace(/=+$/, '');
  var jwtUnsigned = jwtHeader + '.' + jwtClaim;
  var signature   = Utilities.computeRsaSha256Signature(jwtUnsigned, FIREBASE_SECRETS.privateKey);
  var jwt         = jwtUnsigned + '.' + Utilities.base64EncodeWebSafe(signature).replace(/=+$/, '');

  var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions: true
  });

  var data = JSON.parse(resp.getContentText());
  if (!data.access_token) {
    Logger.log('Token error: ' + JSON.stringify(data));
    return null;
  }
  cache.put('FB_STORAGE_TOKEN', data.access_token, 3000);
  return data.access_token;
}

// ============================================================
// testActionCreateFolder — run this once to verify everything works
// ============================================================
function testActionCreateFolder() {
  var result = actionCreateFolder({
    folderName: 'TEST(99999) - Test Employer - 2026-05-13',
    parentType: 'employers'
  });
  Logger.log(JSON.stringify(result, null, 2));
}
// ============================================================
// backfillEmployerFolders — creates Firebase Storage folders
// for all employers that don't have one yet
// Safe to re-run: skips employers that already have a folder
// ============================================================
function backfillEmployerFolders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Employers');
  if (!sheet) { Logger.log('No Employers sheet'); return; }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol     = headers.indexOf('EmployerID');
  var efnCol    = headers.indexOf('EFN');
  var nameCol   = headers.indexOf('EmployerName');
  var agCol     = headers.indexOf('AgencyCode');
  var fnCol     = headers.indexOf('EmployerFolderName');
  var fidCol    = headers.indexOf('EmployerFolderID');
  var furlCol   = headers.indexOf('EmployerFolderUrl');
  var furl2Col  = headers.indexOf('FolderUrl');

  var created = 0, skipped = 0, errors = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var employerId = row[idCol];
    if (!employerId) continue;

    var existingFolderName = (fnCol >= 0 ? String(row[fnCol] || '') : '').trim();
    var existingFolderId   = (fidCol >= 0 ? String(row[fidCol] || '') : '').trim();
    
    // Skip if already has a folder
    if (existingFolderName && existingFolderId) {
      skipped++;
      continue;
    }

    // Build row object for buildEmployerFolderName
    var employerRow = {
      EmployerID:   employerId,
      AgencyCode:   row[agCol] || '',
      EFN:          row[efnCol] || row[nameCol] || ''
    };

    var folderName = buildEmployerFolderName(employerRow);
    
    try {
      var result = actionCreateFolder({ folderName: folderName, parentType: 'employers' });
      if (result && result.success) {
        // Update the row in Sheets
        if (fnCol >= 0)   sheet.getRange(i + 1, fnCol + 1).setValue(folderName);
        if (fidCol >= 0)  sheet.getRange(i + 1, fidCol + 1).setValue(result.folderId);
        if (furlCol >= 0) sheet.getRange(i + 1, furlCol + 1).setValue(result.folderUrl);
        if (furl2Col >= 0) sheet.getRange(i + 1, furl2Col + 1).setValue(result.folderUrl);

        // Also mirror to Firestore so Employers Manager page shows folder
        try {
          var fsDoc = {
            EmployerID:         employerId,
            EmployerFolderName: folderName,
            EmployerFolderID:   result.folderId,
            EmployerFolderUrl:  result.folderUrl,
            FolderUrl:          result.folderUrl
          };
          // Patch into Firestore (merge with existing doc)
          firestorePatchDoc('employers', String(employerId), fsDoc);
        } catch(fe) {
          Logger.log('Firestore mirror failed for ID ' + employerId + ': ' + fe.message);
        }

        created++;
        if (created % 25 === 0) Logger.log('Progress: ' + created + ' created, ' + skipped + ' skipped, ' + errors + ' errors');
      } else {
        errors++;
        Logger.log('Failed ID ' + employerId + ': ' + (result ? result.error : 'unknown'));
      }
    } catch(e) {
      errors++;
      Logger.log('Exception ID ' + employerId + ': ' + e.message);
    }

    // Tiny pause to avoid rate-limit
    if (created % 50 === 49) Utilities.sleep(500);
  }

  Logger.log('=== BACKFILL COMPLETE ===');
  Logger.log('Created : ' + created);
  Logger.log('Skipped : ' + skipped);
  Logger.log('Errors  : ' + errors);
  Logger.log('Total   : ' + (created + skipped + errors));
}

// Helper: PATCH a Firestore document (merge fields)
// If firestoreSaveDoc / firestorePatchDoc already exists, this is harmless
function firestorePatchDoc(collection, docId, fields) {
  if (typeof firestoreSaveDoc === 'function') {
    // firestoreSaveDoc usually does a setDoc which overwrites — we use patch via REST
  }
  var token = getFirestoreAccessToken();
  if (!token) throw new Error('No service account token');

  var url = 'https://firestore.googleapis.com/v1/projects/' + FIREBASE_SECRETS.projectId
          + '/databases/(default)/documents/' + collection + '/' + docId
          + '?updateMask.fieldPaths=' + Object.keys(fields).map(function(k){return encodeURIComponent(k);}).join('&updateMask.fieldPaths=');

  // Build Firestore-flavored fields object
  var fsFields = {};
  Object.keys(fields).forEach(function(k) {
    var v = fields[k];
    if (typeof v === 'string') fsFields[k] = { stringValue: v };
    else if (typeof v === 'number') {
      fsFields[k] = Number.isInteger(v) ? { integerValue: String(v) } : { doubleValue: v };
    } else if (typeof v === 'boolean') fsFields[k] = { booleanValue: v };
    else fsFields[k] = { stringValue: String(v) };
  });

  var resp = UrlFetchApp.fetch(url, {
    method: 'patch',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ fields: fsFields }),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code !== 200) {
    Logger.log('Firestore patch failed: ' + code + ' — ' + resp.getContentText());
  }
}

/**
 * Push (create or update) a single contact into the script-owner's Google Contacts
 * using the People API advanced service.
 *
 * Input payload (from Firebase):
 *   {
 *     ContactID:            "0123",                       // Firestore contact id, used in Notes for traceback
 *     GoogleResourceName:   "people/c1234567890",         // optional; when present we update, otherwise we create
 *     FirstName_EN:         "John",
 *     MiddleName_EN:        "",
 *     LastName_EN:          "Doe",
 *     FirstName_AR:         "...", MiddleName_AR: "...", LastName_AR: "...",
 *     EFN:                  "John Doe",                   // fallback display name
 *     CompanyName_EN:       "Acme",                       // optional
 *     ClientType:           "Personal" | "Company",
 *     MobileNumber:         "+96170...",
 *     MobileNumber2..10:    "...",
 *     Email:                "j@x.com",
 *     EmployerEmail:        "...",
 *     EmailAddress:         "...",
 *     Address:              "...",
 *     City:                 "...",
 *     Country:              "...",
 *     BirthDate:            "1980-05-23",                 // optional
 *     EmployerNotes:        "...",                        // optional
 *     DetectedCodes:        ["LRA","MGSC"]                // we prepend these as a prefix to the family name
 *   }
 *
 * Returns: { success: true, resourceName: "people/c..." , action: "created" | "updated" }
 *      or: { success: false, error: "..." }
 *
 * REQUIRED ONE-TIME SETUP IN APPS SCRIPT:
 *   1. Services (left sidebar, + icon) → add "People API" → identifier "People"
 *   2. Console / appsscript.json: oauthScopes must include
 *        "https://www.googleapis.com/auth/contacts"
 */
/**
 * Push (create or update) a single contact into the script-owner's Google Contacts
 * using the People API advanced service.
 *
 * If the contact has a GoogleResourceName but that contact no longer exists in
 * Google (e.g. it was deleted from Contacts), this falls back to CREATE instead
 * of failing.
 */
function pushContactToGoogle_(payload) {
  try {
    if (!payload || typeof payload !== 'object') {
      return { success: false, error: 'No payload received' };
    }

    var person = _buildPersonResource_(payload);
    var existingResourceName = (payload.GoogleResourceName || '').trim();

    var result;
    var action;

    // ── STEP 1: try direct UPDATE using the stored resource name
    if (existingResourceName) {
      try {
        var current = People.People.get(existingResourceName, {
          personFields: 'metadata,names,phoneNumbers,emailAddresses,addresses,organizations,birthdays,events,biographies'
        });
        person.etag = current.etag;
        result = People.People.updateContact(person, existingResourceName, {
          updatePersonFields: 'names,phoneNumbers,emailAddresses,addresses,organizations,birthdays,events,biographies'
        });
        return { success: true, action: 'updated', resourceName: result.resourceName };
      } catch (innerErr) {
        var msg = String(innerErr && innerErr.message || innerErr);
        var notFound = (msg.indexOf('not found') !== -1 || msg.indexOf('Not Found') !== -1 ||
                        msg.indexOf('not be found') !== -1 || msg.indexOf('غير موجود') !== -1);
        if (!notFound) throw innerErr;
        // stored resource name is dead — fall through to search-by-note
      }
    }

    // ── STEP 2: stored name missing or dead → search Google Contacts
    //    for the note "MGSDB ContactID: <id>" (the vCard import has it)
    //    so we RELINK to the existing contact instead of duplicating.
    var foundResourceName = null;
    if (payload.ContactID) {
      foundResourceName = _findContactByMGSDBID_(String(payload.ContactID));
    }

    if (foundResourceName) {
      var found = People.People.get(foundResourceName, {
        personFields: 'metadata,names,phoneNumbers,emailAddresses,addresses,organizations,birthdays,events,biographies'
      });
      person.etag = found.etag;
      result = People.People.updateContact(person, foundResourceName, {
        updatePersonFields: 'names,phoneNumbers,emailAddresses,addresses,organizations,birthdays,events,biographies'
      });
      action = 'relinked';
    } else {
      // ── STEP 3: truly not in Google → create fresh
      result = People.People.createContact(person);
      action = 'created';
    }

    return {
      success: true,
      action: action,
      resourceName: result.resourceName
    };
  } catch (err) {
    return { success: false, error: String(err && err.message || err) };
  }
}

/**
 * Scan the script-owner's Google Contacts for a contact whose Notes
 * (biography) contains the exact line "MGSDB ContactID: <contactId>".
 * Returns the resourceName string, or null if not found.
 *
 * READ-ONLY — never modifies any contact.
 */
function _findContactByMGSDBID_(contactId) {
  try {
    var target = String(contactId).trim();
    if (!target) return null;

    var pageToken = null;
    do {
      var params = {
        personFields: 'biographies',
        pageSize: 1000
      };
      if (pageToken) params.pageToken = pageToken;

      var resp = People.People.Connections.list('people/me', params);
      var connections = resp.connections || [];

      for (var i = 0; i < connections.length; i++) {
        var p = connections[i];
        if (!p.biographies || !p.biographies.length) continue;
        for (var b = 0; b < p.biographies.length; b++) {
          var note = String(p.biographies[b].value || '');
          var m = note.match(/MGSDB ContactID:\s*(\S+)/);
          if (m && m[1] === target) {
            return p.resourceName;
          }
        }
      }

      pageToken = resp.nextPageToken || null;
    } while (pageToken);

    return null;
  } catch (e) {
    Logger.log('_findContactByMGSDBID_ error: ' + e.message);
    return null; // fail safe — caller will create instead
  }
}

/**
 * Build a People API "Person" resource from the Firestore contact payload.
 * Prefix codes (e.g. "LRA MGSC") go into the honorificPrefix field so they
 * display BEFORE the first name in Google Contacts — NOT mixed into the last name.
 */
function _buildPersonResource_(p) {
  var prefixCodes = Array.isArray(p.DetectedCodes) ? p.DetectedCodes.join(' ') : '';
  var givenName  = (p.FirstName_EN  || '').trim();
  var middleName = (p.MiddleName_EN || '').trim();
  var lastName   = (p.LastName_EN   || '').trim();

  var nameEntry = {
    givenName:  givenName,
    middleName: middleName,
    familyName: lastName
  };
  if (prefixCodes) {
    nameEntry.honorificPrefix = prefixCodes;
  }
  var person = { names: [nameEntry] };

  // Phones — MobileNumber, then MobileNumber2 … MobileNumber10
  var phones = [];
  if (p.MobileNumber) phones.push({ value: String(p.MobileNumber), type: 'mobile' });
  for (var i = 2; i <= 10; i++) {
    var key = 'MobileNumber' + i;
    if (p[key]) phones.push({ value: String(p[key]), type: 'mobile' });
  }
  if (phones.length) person.phoneNumbers = phones;

  // Emails — Email, EmployerEmail, EmailAddress
  var emails = [];
  if (p.Email)          emails.push({ value: String(p.Email),          type: 'home' });
  if (p.EmployerEmail)  emails.push({ value: String(p.EmployerEmail),  type: 'work' });
  if (p.EmailAddress)   emails.push({ value: String(p.EmailAddress),   type: 'other' });
  if (emails.length) person.emailAddresses = emails;

  // Address
  if (p.Address || p.City || p.Country) {
    person.addresses = [{
      streetAddress: p.Address || '',
      city:          p.City    || '',
      country:       p.Country || '',
      type: 'home'
    }];
  }

  // Organization (only if marked as Company)
  if (p.ClientType === 'Company' && p.CompanyName_EN) {
    person.organizations = [{ name: String(p.CompanyName_EN) }];
  }

  // Birthday
  if (p.BirthDate) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(p.BirthDate));
    if (m) {
      person.birthdays = [{ date: { year: +m[1], month: +m[2], day: +m[3] } }];
    }
  }

   // Anniversary + Future Appointments — Google "events" with custom labels
  var events = [];

  if (p.Anniversary) {
    var mA = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(p.Anniversary));
    if (mA) {
      events.push({
        date: { year: +mA[1], month: +mA[2], day: +mA[3] },
        type: 'IKAMA'
      });
    }
  }

  if (Array.isArray(p.NamedDates)) {
    p.NamedDates.forEach(function(nd) {
      if (!nd || !nd.label || !nd.date) return;
      var mN = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(nd.date));
      if (mN) {
        events.push({
          date: { year: +mN[1], month: +mN[2], day: +mN[3] },
          type: String(nd.label).trim().toUpperCase()
        });
      }
    });
  }

  if (events.length) person.events = events;

  // Notes — keep the Firebase ContactID for round-trip identification
  var noteLines = [];
  if (p.ContactID)     noteLines.push('MGSDB ContactID: ' + p.ContactID);
  if (p.EmployerNotes) noteLines.push(p.EmployerNotes);
  if (noteLines.length) person.biographies = [{ value: noteLines.join('\n'), contentType: 'TEXT_PLAIN' }];

  return person;
}

/**
 * Manual test runner.
 */
function testPushContact() {
  var result = pushContactToGoogle_({
    ContactID: 'TEST-001',
    FirstName_EN: 'Test',
    LastName_EN: 'FromMGSDB',
    MobileNumber: '+96170000000',
    Email: 'test@example.com',
    ClientType: 'Personal',
    DetectedCodes: ['TEST']
  });
  Logger.log(JSON.stringify(result, null, 2));
}

// ════════════════════════════════════════════════════════════════
// BACKFILL DATES MODULE — paste at the very END of ContactsSync.gs
// ════════════════════════════════════════════════════════════════
// WHAT IT DOES (run AFTER deleting all Google contacts + importing
// the vCard file):
//   1. Scans Google Contacts ONCE → maps "MGSDB ContactID" → resourceName
//   2. Stamps _googleResourceName into every Firestore address_book doc
//      (so all future pushes update instead of duplicate)
//   3. For every contact with BirthDate and/or Anniversary:
//      writes the Birthday + IKAMA event into Google Contacts via API
//
// Resumable: saves position, safe to run many times. Auto-trigger
// runs it every 5 minutes until finished, then removes itself.
//
// HOW TO USE:
//   A) Run  backfillDatesRunner()  manually — repeats until done, OR
//   B) Run  setupBackfillTrigger() once — fully automatic
//   Check progress anytime with  checkBackfillStatus()
// ════════════════════════════════════════════════════════════════

var BACKFILL_POS_KEY  = 'BACKFILL_DATES_POS';
var BACKFILL_DONE_KEY = 'BACKFILL_DATES_DONE';

// ── List ALL address_book docs straight from Firestore REST
//    (self-contained: returns docId + the few fields we need)
function _backfillListContacts_() {
  var token = getFirestoreAccessToken();
  if (!token) throw new Error('No Firestore token');

  var out = [];
  var pageToken = null;

  do {
    var url = 'https://firestore.googleapis.com/v1/projects/' + FIREBASE_SECRETS.projectId
            + '/databases/(default)/documents/address_book?pageSize=300'
            + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');

    var resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error('Firestore list error: ' + resp.getContentText().substring(0, 200));
    }

    var data = JSON.parse(resp.getContentText());
    (data.documents || []).forEach(function(d) {
      var f = d.fields || {};
      function val(k) {
        if (!f[k]) return '';
        if (f[k].stringValue  != null) return String(f[k].stringValue);
        if (f[k].integerValue != null) return String(f[k].integerValue);
        return '';
      }
      out.push({
        docId:        d.name.split('/').pop(),
        ContactID:    val('ContactID'),
        BirthDate:    val('BirthDate'),
        Anniversary:  val('Anniversary'),
        StoredRN:     val('_googleResourceName'),
        Archived:     f.Archived && f.Archived.booleanValue === true
      });
    });

    pageToken = data.nextPageToken || null;
  } while (pageToken);

  return out;
}

// ── Scan Google Contacts once → { ContactID: resourceName }
function _backfillGoogleMap_() {
  var map = {};
  var pageToken = null;
  do {
    var params = { personFields: 'biographies', pageSize: 1000 };
    if (pageToken) params.pageToken = pageToken;
    var resp = People.People.Connections.list('people/me', params);
    (resp.connections || []).forEach(function(p) {
      if (!p.biographies) return;
      for (var b = 0; b < p.biographies.length; b++) {
        var m = String(p.biographies[b].value || '').match(/MGSDB ContactID:\s*(\S+)/);
        if (m) { map[m[1]] = p.resourceName; return; }
      }
    });
    pageToken = resp.nextPageToken || null;
  } while (pageToken);
  return map;
}

// ── MAIN — processes contacts from saved position, ~4.5 min per run
function backfillDatesRunner() {
  var props = PropertiesService.getScriptProperties();

  if (props.getProperty(BACKFILL_DONE_KEY) === 'true') {
    Logger.log('✅ Backfill already complete. Run resetBackfill() to start over.');
    return { success: true, done: true };
  }

  var startMs = new Date().getTime();
  var MAX_MS  = 4.5 * 60 * 1000;

  Logger.log('Loading Firestore contacts…');
  var contacts = _backfillListContacts_();
  Logger.log('Firestore contacts: ' + contacts.length);

  Logger.log('Scanning Google Contacts…');
  var gmap = _backfillGoogleMap_();
  Logger.log('Google contacts with MGSDB ContactID: ' + Object.keys(gmap).length);

  var pos = parseInt(props.getProperty(BACKFILL_POS_KEY) || '0', 10);
  var stamped = 0, dated = 0, skippedNoMatch = 0, errors = 0;

  for (var i = pos; i < contacts.length; i++) {
    if (new Date().getTime() - startMs > MAX_MS) {
      props.setProperty(BACKFILL_POS_KEY, String(i));
      Logger.log('⏸ Time limit — paused at ' + i + '/' + contacts.length
        + ' | stamped:' + stamped + ' dated:' + dated + ' errors:' + errors);
      return { success: true, done: false, position: i, total: contacts.length };
    }

    var c = contacts[i];
    if (c.Archived || !c.ContactID) continue;

    var rn = gmap[c.ContactID];
    if (!rn) { skippedNoMatch++; continue; }

    try {
      // 1. Stamp resource name into Firestore (only if different)
      if (c.StoredRN !== rn) {
        firestorePatchDoc('address_book', c.docId, { _googleResourceName: rn });
        stamped++;
      }

      // 2. Push dates (birthday + IKAMA) if the contact has any
      if (c.BirthDate || c.Anniversary) {
        var person = { };
        var fields = [];

        var mB = /^(\d{4})-(\d{2})-(\d{2})/.exec(c.BirthDate || '');
        if (mB) {
          person.birthdays = [{ date: { year: +mB[1], month: +mB[2], day: +mB[3] } }];
          fields.push('birthdays');
        }
        var mA = /^(\d{4})-(\d{2})-(\d{2})/.exec(c.Anniversary || '');
        if (mA) {
          person.events = [{ date: { year: +mA[1], month: +mA[2], day: +mA[3] }, type: 'IKAMA' }];
          fields.push('events');
        }

        if (fields.length) {
          var current = People.People.get(rn, { personFields: 'metadata' });
          person.etag = current.etag;
          People.People.updateContact(person, rn, { updatePersonFields: fields.join(',') });
          dated++;
          Utilities.sleep(800); // respect People API write quota
        }
      }
    } catch (e) {
      errors++;
      Logger.log('❌ ContactID ' + c.ContactID + ': ' + e.message);
    }
  }

  // Finished all contacts
  props.deleteProperty(BACKFILL_POS_KEY);
  props.setProperty(BACKFILL_DONE_KEY, 'true');
  removeBackfillTrigger();
  Logger.log('🎉 BACKFILL COMPLETE | stamped:' + stamped + ' dated:' + dated
    + ' noMatch:' + skippedNoMatch + ' errors:' + errors);
  return { success: true, done: true, stamped: stamped, dated: dated, errors: errors };
}

// ── Automatic mode: runs every 5 min until done, then self-removes
function setupBackfillTrigger() {
  removeBackfillTrigger();
  PropertiesService.getScriptProperties().deleteProperty(BACKFILL_DONE_KEY);
  ScriptApp.newTrigger('backfillDatesRunner').timeBased().everyMinutes(5).create();
  Logger.log('✅ Backfill trigger installed (every 5 min). It removes itself when done.');
}

function removeBackfillTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'backfillDatesRunner') ScriptApp.deleteTrigger(t);
  });
}

function checkBackfillStatus() {
  var props = PropertiesService.getScriptProperties();
  Logger.log('Done: '     + (props.getProperty(BACKFILL_DONE_KEY) || 'false'));
  Logger.log('Position: ' + (props.getProperty(BACKFILL_POS_KEY)  || '0'));
}

function resetBackfill() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(BACKFILL_POS_KEY);
  props.deleteProperty(BACKFILL_DONE_KEY);
  Logger.log('Backfill state reset — next run starts from the beginning.');
}
