// ============================================================
// WhatsApp_API.gs — Twilio WhatsApp Integration for MGSDB
// ============================================================
// Setup: GAS → Project Settings → Script Properties → Add:
//   TWILIO_SID    = ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
//   TWILIO_TOKEN  = your_auth_token
//   TWILIO_FROM: 'whatsapp:++14155238886'  (or your paid number)
// ============================================================

var MESSAGES_SHEET = 'Messages';

function sendWhatsAppMessage(toNumber, messageText, mediaUrl) {
  var props    = PropertiesService.getScriptProperties();
  var sid      = props.getProperty('TWILIO_SID');
  var token    = props.getProperty('TWILIO_TOKEN');
  var fromNum  = props.getProperty('TWILIO_FROM');

  if (!sid || !token || !fromNum) {
    throw new Error('Twilio credentials not configured in Script Properties.');
  }

  var to = toNumber.startsWith('whatsapp:') ? toNumber : 'whatsapp:' + toNumber;

  var payload = { From: fromNum, To: to, Body: messageText || '' };
  if (mediaUrl) payload['MediaUrl'] = mediaUrl;

  var options = {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(sid + ':' + token) },
    payload: payload,
    muteHttpExceptions: true
  };

  var url      = 'https://api.twilio.com/2010-04-01/Accounts/' + sid + '/Messages.json';
  var response = UrlFetchApp.fetch(url, options);
  var result   = JSON.parse(response.getContentText());

  if (result.error_code) throw new Error('Twilio error ' + result.error_code + ': ' + result.error_message);
  return { success: true, sid: result.sid, status: result.status };
}

function scheduleMessage(params) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MESSAGES_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(MESSAGES_SHEET);
    sheet.appendRow(['MessageID','CreatedDate','RecipientType','RecipientID',
      'RecipientName','RecipientPhone','MessageText','MediaURL',
      'MediaType','ScheduledDate','ScheduledTime','Status',
      'SentDate','ErrorMessage','SentBy']);
  }

  var now = new Date();
  var messageID = 'MSG-' + now.getFullYear() +
                  String(now.getMonth()+1).padStart(2,'0') +
                  String(now.getDate()).padStart(2,'0') + '-' +
                  String(sheet.getLastRow()).padStart(4,'0');

  sheet.appendRow([messageID,
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    params.recipientType||'Employer', params.recipientID||'',
    params.recipientName||'', params.recipientPhone||'',
    params.messageText||'', params.mediaURL||'', params.mediaType||'',
    params.scheduledDate||'', params.scheduledTime||'09:00',
    'Pending','','', params.sentBy||'Admin']);

  return { success: true, messageID: messageID };
}

function processPendingMessages() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MESSAGES_SHEET);
  if (!sheet) return;

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var now   = new Date();
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];
  var col = {};
  headers.forEach(function(h,i){ col[h]=i; });

  var sent=0, failed=0;
  for (var r=1; r<data.length; r++) {
    var row = data[r];
    if (row[col['Status']] !== 'Pending') continue;
    var scheduledDT = new Date(row[col['ScheduledDate']] + ' ' + (row[col['ScheduledTime']]||'09:00'));
    if (isNaN(scheduledDT.getTime()) || scheduledDT > now) continue;

    try {
      sendWhatsAppMessage(row[col['RecipientPhone']], row[col['MessageText']], row[col['MediaURL']]||null);
      sheet.getRange(r+1, col['Status']+1).setValue('Sent');
      sheet.getRange(r+1, col['SentDate']+1).setValue(
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
      sent++;
    } catch(e) {
      sheet.getRange(r+1, col['Status']+1).setValue('Failed');
      sheet.getRange(r+1, col['ErrorMessage']+1).setValue(e.message);
      failed++;
    }
  }
  return { sent: sent, failed: failed };
}

function deleteMessage(params) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MESSAGES_SHEET);
  if (!sheet) return { success: false, error: 'Messages sheet not found' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var col = {};
  headers.forEach(function(h,i){ col[h]=i; });

  for (var r=1; r<data.length; r++) {
    if (data[r][col['MessageID']] === params.messageID) {
      sheet.deleteRow(r+1);
      return { success: true };
    }
  }
  return { success: false, error: 'Message not found' };
}

function sendNow(params) {
  try {
    var result = sendWhatsAppMessage(params.recipientPhone, params.messageText, params.mediaURL||null);
    var now  = new Date();
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MESSAGES_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(MESSAGES_SHEET);
      sheet.appendRow(['MessageID','CreatedDate','RecipientType','RecipientID',
        'RecipientName','RecipientPhone','MessageText','MediaURL',
        'MediaType','ScheduledDate','ScheduledTime','Status',
        'SentDate','ErrorMessage','SentBy']);
    }
    var messageID = 'MSG-' + now.getFullYear() +
                    String(now.getMonth()+1).padStart(2,'0') +
                    String(now.getDate()).padStart(2,'0') + '-' +
                    String(sheet.getLastRow()).padStart(4,'0');
    sheet.appendRow([messageID,
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      params.recipientType||'Employer', params.recipientID||'',
      params.recipientName||'', params.recipientPhone||'',
      params.messageText||'', params.mediaURL||'', params.mediaType||'',
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm'),
      'Sent',
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      '', params.sentBy||'Admin']);

    return { success: true, messageID: messageID, twilioSID: result.sid };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processPendingMessages') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processPendingMessages').timeBased().everyDays(1).atHour(9).create();
  Logger.log('✅ Daily trigger set for 9:00 AM');
}

function testSendWhatsApp() {
  var props = PropertiesService.getScriptProperties();
  var testNumber = props.getProperty('TEST_PHONE') || '+96170000000';
  var result = sendWhatsAppMessage(testNumber,
    '✅ MGSDB WhatsApp Test\n\nTwilio connected successfully!\n\n— MGS Agency', null);
  Logger.log(JSON.stringify(result));
}


function testWhatsAppSend() {
  var result = sendWhatsAppMessage(
    '+9613823339',
    'Test message from MGS system',
    null
  );
  Logger.log(JSON.stringify(result));
}
