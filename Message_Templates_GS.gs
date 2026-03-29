// ══════════════════════════════════════════════════════════════════
// Message_Templates_GS.gs
// MGS MGSDB2026 — Template engine with Rules & Conditions support
// Updated: 2026-03-13
// ══════════════════════════════════════════════════════════════════

var TEMPLATES_SHEET = 'MessageTemplates';

// Full column definition — order matters for new sheet creation
var TEMPLATE_COLUMNS = [
  'TemplateID',
  'TemplateName',
  'Category',
  'MessageBody',
  'ImageURL',
  'Active',
  'CreatedDate',
  'UpdatedDate',
  // Rules & Conditions columns
  'RecipientType',    // All | Employer | Worker | Agency
  'TriggerEvent',     // Manual | Birthday | Anniversary | PassportExpiry | DocumentReady | WorkerArrival | Scheduled
  'DaysBefore',       // integer — send N days before event
  'DaysAfter',        // integer — send N days after event
  'SendOnce',         // boolean — never resend to same recipient
  'MinDaysBetween',   // integer — minimum days between sends to same recipient
  'RequiredField'     // comma-separated — only show if recipient has these fields
];

// ══════════════════════════════════════════════
// SHEET SETUP
// ══════════════════════════════════════════════

function getOrCreateTemplatesSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEMPLATES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TEMPLATES_SHEET);
    sheet.appendRow(TEMPLATE_COLUMNS);
    sheet.setFrozenRows(1);
    // Format header row
    var headerRange = sheet.getRange(1, 1, 1, TEMPLATE_COLUMNS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1c2030');
  }
  return sheet;
}

// ══════════════════════════════════════════════
// MIGRATION — adds new rule columns to existing sheet
// Call once from Admin Dashboard after deploying this update
// ══════════════════════════════════════════════

function migrateTemplatesSheet() {
  try {
    var sheet   = getOrCreateTemplatesSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var added   = [];

    // Default values for each new column
    var DEFAULTS = {
      'RecipientType':   'All',
      'TriggerEvent':    'Manual',
      'DaysBefore':      '',
      'DaysAfter':       '',
      'SendOnce':        false,
      'MinDaysBetween':  '',
      'RequiredField':   ''
    };

    var newCols = ['RecipientType','TriggerEvent','DaysBefore','DaysAfter','SendOnce','MinDaysBetween','RequiredField'];

    newCols.forEach(function(colName) {
      if (headers.indexOf(colName) === -1) {
        var nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue(colName);
        sheet.getRange(1, nextCol).setFontWeight('bold');

        // Fill default values for all existing data rows
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          var def = DEFAULTS[colName];
          var vals = [];
          for (var r = 0; r < lastRow - 1; r++) {
            vals.push([def !== undefined ? def : '']);
          }
          sheet.getRange(2, nextCol, lastRow - 1, 1).setValues(vals);
        }
        headers.push(colName);
        added.push(colName);
      }
    });

    var msg = added.length > 0
      ? 'Added ' + added.length + ' column(s): ' + added.join(', ')
      : 'Sheet already up to date — no changes needed';

    return { success: true, message: msg, added: added };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════
// GET TEMPLATES — returns all templates (no ImageURL for performance)
// ══════════════════════════════════════════════

function getTemplates() {
  try {
    var sheet = getOrCreateTemplatesSheet();
    if (sheet.getLastRow() < 2) return { success: true, data: [] };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var result  = [];

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      // Skip completely empty rows
      if (!row[0] && !row[1]) continue;

      var obj = {};
      headers.forEach(function(h, i) {
        if (h === 'ImageURL') return; // Skip large base64 data for performance
        obj[h] = row[i];
      });

      // Normalize boolean fields
      obj.Active   = _toBool(obj.Active);
      obj.SendOnce = _toBool(obj.SendOnce);

      // Normalize numeric fields
      obj.DaysBefore     = obj.DaysBefore     !== '' ? Number(obj.DaysBefore)     : '';
      obj.DaysAfter      = obj.DaysAfter      !== '' ? Number(obj.DaysAfter)      : '';
      obj.MinDaysBetween = obj.MinDaysBetween !== '' ? Number(obj.MinDaysBetween) : '';

      // Default rule values for old templates that pre-date migration
      if (!obj.RecipientType) obj.RecipientType = 'All';
      if (!obj.TriggerEvent)  obj.TriggerEvent  = 'Manual';

      result.push(obj);
    }

    result.sort(function(a, b) {
      if (a.Active !== b.Active) return b.Active ? 1 : -1;
      return (a.TemplateName||'').localeCompare(b.TemplateName||'');
    });

    return { success: true, data: result };
  } catch(e) {
    return { success: false, error: e.message, data: [] };
  }
}

// ══════════════════════════════════════════════
// GET TEMPLATES FOR CONTEXT
// Used by Messaging Manager to get context-filtered templates
// params: { recipientType, triggerEvent, employerData }
// ══════════════════════════════════════════════

function getTemplatesForContext(params) {
  try {
    var result = getTemplates();
    if (!result.success) return result;

    var templates    = result.data;
    var recipientType = (params && params.recipientType) ? params.recipientType : 'Employer';
    var triggerFilter = (params && params.triggerEvent)  ? params.triggerEvent  : null; // null = show all
    var employerData  = (params && params.employerData)  ? params.employerData  : {};

    var filtered = templates.filter(function(t) {
      // Must be active
      if (!t.Active) return false;

      // Recipient type filter: show if 'All' or matches context
      var tRecipient = t.RecipientType || 'All';
      if (tRecipient !== 'All' && tRecipient !== recipientType) return false;

      // Trigger event filter (optional — if caller passes triggerEvent, only show matching)
      if (triggerFilter && t.TriggerEvent && t.TriggerEvent !== 'Manual' && t.TriggerEvent !== triggerFilter) return false;

      // Required fields check: only show if recipient has all required fields non-empty
      if (t.RequiredField && String(t.RequiredField).trim()) {
        var required = String(t.RequiredField).split(',').map(function(f){ return f.trim(); }).filter(Boolean);
        for (var i = 0; i < required.length; i++) {
          var fieldVal = employerData[required[i]];
          if (!fieldVal || String(fieldVal).trim() === '') return false;
        }
      }

      return true;
    });

    return { success: true, data: filtered };
  } catch(e) {
    return { success: false, error: e.message, data: [] };
  }
}

// ══════════════════════════════════════════════
// GET TEMPLATES PUBLIC — called from Messaging_Manager via google.script.run
// Returns filtered templates suitable for display (no huge ImageURLs)
// ══════════════════════════════════════════════

function getTemplates() {
  try {
    var sheet = getOrCreateTemplatesSheet();
    if (sheet.getLastRow() < 2) return { success: true, data: [] };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var result  = [];

    for (var r = 1; r < data.length; r++) {
      var obj = {};
      headers.forEach(function(h, i) {
        if (h !== 'ImageURL') obj[h] = data[r][i];
      });
      obj.Active = (obj.Active === true || obj.Active === 'TRUE' || obj.Active === 'true');
      result.push(obj);
    }

    result.sort(function(a, b) {
      if (a.Active !== b.Active) return b.Active ? 1 : -1;
      return (a.TemplateName||'').localeCompare(b.TemplateName||'');
    });

    return { success: true, data: result };

  } catch(e) {
    return { success: false, error: e.message, data: [] };
  }
}

// ══════════════════════════════════════════════
// GET TEMPLATES FOR CONTEXT PUBLIC
// ══════════════════════════════════════════════

function getTemplatesForContextPublic(params) {
  return getTemplatesForContext(params);
}

// ══════════════════════════════════════════════
// SAVE TEMPLATE — create or update
// ══════════════════════════════════════════════

function saveTemplate(params) {
  try {
    var sheet   = getOrCreateTemplatesSheet();
    var now     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var col     = {};
    headers.forEach(function(h, i) { col[h] = i; });

    // Build image URL
    var imageURL = params.imageURL || '';
    if (params.imageBase64 && params.imageBase64.length > 0) {
      imageURL = params.imageBase64; // Store base64 directly — avoids Drive URL blocking in GAS iframe
    }

    // Editing existing template
    if (params.templateID) {
      for (var r = 1; r < data.length; r++) {
        if (data[r][col['TemplateID']] === params.templateID) {
          var rowNum = r + 1;
          _setIfExists(sheet, rowNum, col, 'TemplateName',   params.templateName  || '');
          _setIfExists(sheet, rowNum, col, 'Category',       params.category      || 'Custom');
          _setIfExists(sheet, rowNum, col, 'MessageBody',    params.messageBody   || '');
          if (imageURL) _setIfExists(sheet, rowNum, col, 'ImageURL', imageURL);
          _setIfExists(sheet, rowNum, col, 'Active',         params.active !== false);
          _setIfExists(sheet, rowNum, col, 'UpdatedDate',    now);
          // Rules
          _setIfExists(sheet, rowNum, col, 'RecipientType',  params.recipientType  || 'All');
          _setIfExists(sheet, rowNum, col, 'TriggerEvent',   params.triggerEvent   || 'Manual');
          _setIfExists(sheet, rowNum, col, 'DaysBefore',     params.daysBefore     !== '' && params.daysBefore !== undefined ? Number(params.daysBefore) : '');
          _setIfExists(sheet, rowNum, col, 'DaysAfter',      params.daysAfter      !== '' && params.daysAfter  !== undefined ? Number(params.daysAfter)  : '');
          _setIfExists(sheet, rowNum, col, 'SendOnce',       params.sendOnce       === true || params.sendOnce === 'true');
          _setIfExists(sheet, rowNum, col, 'MinDaysBetween', params.minDaysBetween !== '' && params.minDaysBetween !== undefined ? Number(params.minDaysBetween) : '');
          _setIfExists(sheet, rowNum, col, 'RequiredField',  params.requiredField  || '');
          return { success: true, templateID: params.templateID };
        }
      }
    }

    // New template
    var newID = generateTemplateID(sheet);
    var newRow = TEMPLATE_COLUMNS.map(function(h) {
      switch(h) {
        case 'TemplateID':      return newID;
        case 'TemplateName':    return params.templateName  || '';
        case 'Category':        return params.category      || 'Custom';
        case 'MessageBody':     return params.messageBody   || '';
        case 'ImageURL':        return imageURL;
        case 'Active':          return params.active !== false;
        case 'CreatedDate':     return now;
        case 'UpdatedDate':     return now;
        case 'RecipientType':   return params.recipientType  || 'All';
        case 'TriggerEvent':    return params.triggerEvent   || 'Manual';
        case 'DaysBefore':      return params.daysBefore     !== '' && params.daysBefore !== undefined ? Number(params.daysBefore) : '';
        case 'DaysAfter':       return params.daysAfter      !== '' && params.daysAfter  !== undefined ? Number(params.daysAfter)  : '';
        case 'SendOnce':        return params.sendOnce       === true || params.sendOnce === 'true';
        case 'MinDaysBetween':  return params.minDaysBetween !== '' && params.minDaysBetween !== undefined ? Number(params.minDaysBetween) : '';
        case 'RequiredField':   return params.requiredField  || '';
        default:                return '';
      }
    });
    sheet.appendRow(newRow);
    return { success: true, templateID: newID };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════
// TOGGLE ACTIVE
// ══════════════════════════════════════════════

function toggleTemplate(params) {
  try {
    var sheet   = getOrCreateTemplatesSheet();
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var col     = {};
    headers.forEach(function(h, i) { col[h] = i; });

    for (var r = 1; r < data.length; r++) {
      if (data[r][col['TemplateID']] === params.templateID) {
        sheet.getRange(r + 1, col['Active'] + 1).setValue(params.active);
        sheet.getRange(r + 1, col['UpdatedDate'] + 1).setValue(
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
        );
        return { success: true };
      }
    }
    return { success: false, error: 'Template not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════
// DELETE TEMPLATE
// ══════════════════════════════════════════════

function deleteTemplate(params) {
  try {
    var sheet   = getOrCreateTemplatesSheet();
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var col     = {};
    headers.forEach(function(h, i) { col[h] = i; });

    for (var r = 1; r < data.length; r++) {
      if (data[r][col['TemplateID']] === params.templateID) {
        sheet.deleteRow(r + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Template not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════
// RULES ENGINE — check if a template can be sent
// Used before auto-sending or to validate manual sends
// params: { template, recipientID, recipientType, employerData, sentHistory }
// sentHistory: array of { templateID, recipientID, sentDate }
// ══════════════════════════════════════════════

function checkTemplateRules(template, recipientID, recipientType, employerData, sentHistory) {
  var result = { allowed: true, reason: '' };

  // Recipient type check
  var tRecipient = template.RecipientType || 'All';
  if (tRecipient !== 'All' && tRecipient !== recipientType) {
    return { allowed: false, reason: 'Template is for ' + tRecipient + ' only, not ' + recipientType };
  }

  // Required field check
  if (template.RequiredField && String(template.RequiredField).trim()) {
    var required = String(template.RequiredField).split(',').map(function(f){ return f.trim(); }).filter(Boolean);
    for (var i = 0; i < required.length; i++) {
      var fieldVal = employerData ? employerData[required[i]] : null;
      if (!fieldVal || String(fieldVal).trim() === '') {
        return { allowed: false, reason: 'Recipient is missing required field: ' + required[i] };
      }
    }
  }

  // SendOnce check
  if (template.SendOnce) {
    var alreadySent = (sentHistory || []).some(function(h) {
      return h.templateID === template.TemplateID && h.recipientID === recipientID;
    });
    if (alreadySent) {
      return { allowed: false, reason: 'Template already sent to this recipient (Send Once rule)' };
    }
  }

  // MinDaysBetween check
  if (template.MinDaysBetween && Number(template.MinDaysBetween) > 0) {
    var minDays = Number(template.MinDaysBetween);
    var lastSent = null;
    (sentHistory || []).forEach(function(h) {
      if (h.templateID === template.TemplateID && h.recipientID === recipientID) {
        var d = new Date(h.sentDate);
        if (!lastSent || d > lastSent) lastSent = d;
      }
    });
    if (lastSent) {
      var daysSince = (new Date() - lastSent) / (24 * 3600 * 1000);
      if (daysSince < minDays) {
        var daysLeft = Math.ceil(minDays - daysSince);
        return {
          allowed: false,
          reason: 'Minimum ' + minDays + ' days between sends. Next send allowed in ' + daysLeft + ' day(s)'
        };
      }
    }
  }

  return result;
}

// ══════════════════════════════════════════════
// EVENT DATE CALCULATOR
// Returns the Date on which a template should fire, or null
// ══════════════════════════════════════════════

function getEventSendDate(template, emp, today) {
  var daysBefore = template.DaysBefore !== '' ? Number(template.DaysBefore) : 0;
  var daysAfter  = template.DaysAfter  !== '' ? Number(template.DaysAfter)  : 0;
  var offset     = daysAfter - daysBefore; // positive = after, negative = before

  switch (template.TriggerEvent) {

    case 'Birthday': {
      var bd = emp['BirthDate'] ? new Date(emp['BirthDate']) : null;
      if (!bd || isNaN(bd)) return null;
      var nextBd = new Date(today.getFullYear(), bd.getMonth(), bd.getDate());
      if (nextBd < today) nextBd.setFullYear(today.getFullYear() + 1);
      // Send N days before birthday
      var sendDate = new Date(nextBd);
      sendDate.setDate(sendDate.getDate() - daysBefore + daysAfter);
      return sendDate;
    }

    case 'Anniversary': {
      var ann = emp['Anniversary'] ? new Date(emp['Anniversary']) : null;
      if (!ann || isNaN(ann)) return null;
      var nextAnn = new Date(today.getFullYear(), ann.getMonth(), ann.getDate());
      if (nextAnn < today) nextAnn.setFullYear(today.getFullYear() + 1);
      var sendDate = new Date(nextAnn);
      sendDate.setDate(sendDate.getDate() - daysBefore + daysAfter);
      return sendDate;
    }

    case 'PassportExpiry': {
      var expiry = emp['PassportExpiry'] ? new Date(emp['PassportExpiry']) : null;
      if (!expiry || isNaN(expiry)) return null;
      var sendDate = new Date(expiry);
      sendDate.setDate(sendDate.getDate() - daysBefore + daysAfter);
      return sendDate;
    }

    case 'WorkerArrival': {
      var arrival = emp['ArrivalDate'] ? new Date(emp['ArrivalDate']) : null;
      if (!arrival || isNaN(arrival)) return null;
      var sendDate = new Date(arrival);
      sendDate.setDate(sendDate.getDate() + daysAfter - daysBefore);
      return sendDate;
    }

    case 'DocumentReady': {
      // Manual trigger only — not auto-schedulable
      return null;
    }

    case 'Scheduled': {
      // Recurring sends need separate logic — return today if offset matches
      // This is a simple daily check; implement custom recurrence as needed
      return today;
    }

    default:
      return null;
  }
}

// ══════════════════════════════════════════════
// VARIABLE SUBSTITUTION
// ══════════════════════════════════════════════

function substituteVariables(text, emp, template, today) {
  var name        = (emp['EFN'] || emp['EmployerName'] || '').split('(')[0].trim();
  var agencyPhone = '+961-3-823339'; // TODO: pull from config

  var daysUntil = '';
  if (template.TriggerEvent === 'Birthday' && emp['BirthDate']) {
    var bd = new Date(emp['BirthDate']);
    if (!isNaN(bd)) {
      var next = new Date(today.getFullYear(), bd.getMonth(), bd.getDate());
      if (next < today) next.setFullYear(today.getFullYear() + 1);
      daysUntil = String(Math.ceil((next - today) / (24*3600*1000)));
    }
  }
  if (template.TriggerEvent === 'PassportExpiry' && emp['PassportExpiry']) {
    var exp = new Date(emp['PassportExpiry']);
    if (!isNaN(exp)) {
      daysUntil = String(Math.ceil((exp - today) / (24*3600*1000)));
    }
  }

  return text
    .replace(/\{EmployerName\}/g,   name)
    .replace(/\{WorkerName\}/g,     emp['WorkerName']      || '')
    .replace(/\{PassportExpiry\}/g, emp['PassportExpiry']  || '')
    .replace(/\{DocumentType\}/g,   emp['DocumentType']    || '')
    .replace(/\{AgencyName\}/g,     'MGS Agency')
    .replace(/\{AgencyPhone\}/g,    agencyPhone)
    .replace(/\{Date\}/g,           Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy'))
    .replace(/\{DaysUntil\}/g,      daysUntil);
}

// ══════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════

function _toBool(val) {
  return val === true || val === 'TRUE' || val === 'true';
}

function _setIfExists(sheet, rowNum, col, colName, value) {
  if (col[colName] !== undefined) {
    sheet.getRange(rowNum, col[colName] + 1).setValue(value);
  }
}

function buildMessageRow(headers, fields) {
  var colMap = {
    'MessageID':     'messageID',
    'RecipientType': 'recipientType',
    'RecipientID':   'recipientID',
    'RecipientName': 'recipientName',
    'RecipientPhone':'recipientPhone',
    'MessageText':   'messageText',
    'MediaURL':      'mediaURL',
    'ScheduledDate': 'scheduledDate',
    'ScheduledTime': 'scheduledTime',
    'Status':        'status',
    'SentDate':      'sentDate',
    'ErrorMessage':  'errorMessage',
    'CreatedDate':   'createdDate',
    'TemplateID':    'templateID',
    'TriggerEvent':  'triggerEvent'
  };
  return headers.map(function(h) {
    var key = colMap[h];
    return key ? (fields[key] !== undefined ? fields[key] : '') : '';
  });
}

function generateTemplateID(sheet) {
  var lastRow = sheet.getLastRow();
  var num     = String(lastRow).padStart(3, '0');
  return 'TPL-' + num;
}

function saveTemplateImageToDrive(base64Data, name) {
  var parts    = base64Data.split(',');
  var mimeType = parts[0].match(/:(.*?);/)[1];
  var data     = Utilities.newBlob(Utilities.base64Decode(parts[1]), mimeType, 'template_' + name + '.jpg');
  var folderName = 'MGSDB_Template_Images';
  var folders    = DriveApp.getFoldersByName(folderName);
  var folder     = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  var file       = folder.createFile(data);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}

// ══════════════════════════════════════════════
// TEST FUNCTIONS
// ══════════════════════════════════════════════

function testGetTemplates() {
  var result = getTemplates();
  Logger.log(JSON.stringify(result));
}

function testMigration() {
  var result = migrateTemplatesSheet();
  Logger.log(JSON.stringify(result));
}

function testGetTemplatesForContext() {
  var result = getTemplatesForContext({
    recipientType: 'Employer',
    employerData: { BirthDate: '1980-05-15', MobileNumber: '+9611234567' }
  });
  Logger.log(JSON.stringify(result));
}

function testGetTemplatesPublic() {
  var result = getTemplatesPublic();
  Logger.log(JSON.stringify(result));
}
