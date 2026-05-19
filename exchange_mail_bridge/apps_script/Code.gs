const DEFAULT_SHEET_NAME = 'Exchange Mail';

function doGet() {
  return jsonResponse({
    ok: true,
    service: 'exchange-mail-bridge',
    message: 'POST messages to this web app endpoint.'
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();
    const expectedToken = props.getProperty('BRIDGE_TOKEN');
    const sheetId = props.getProperty('SHEET_ID');
    const sheetName = props.getProperty('SHEET_NAME') || DEFAULT_SHEET_NAME;

    if (!expectedToken || !sheetId) {
      return jsonResponse({ ok: false, error: 'Missing SHEET_ID or BRIDGE_TOKEN in Script Properties.' }, 500);
    }

    const payload = parsePayload_(e);
    if (payload.token !== expectedToken) {
      return jsonResponse({ ok: false, error: 'Unauthorized.' }, 401);
    }

    const messages = Array.isArray(payload.messages) ? payload.messages : [];
    const sheet = getOrCreateSheet_(sheetId, sheetName);
    ensureHeader_(sheet);

    const existingIds = getExistingIds_(sheet);
    const now = new Date().toISOString();
    const rows = [];
    let skipped = 0;

    messages.forEach((message) => {
      const messageId = clean_(message.id);
      if (!messageId || existingIds.has(messageId)) {
        skipped += 1;
        return;
      }

      existingIds.add(messageId);
      rows.push([
        messageId,
        clean_(message.receivedTime),
        clean_(message.fromName),
        clean_(message.fromEmail),
        clean_(message.subject),
        clean_(message.bodyPreview),
        Boolean(message.hasAttachments),
        clean_(message.conversationId),
        clean_(message.webLink),
        now
      ]);
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return jsonResponse({ ok: true, appended: rows.length, skipped });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) }, 500);
  } finally {
    lock.releaseLock();
  }
}

function setupConfigOnce() {
  PropertiesService.getScriptProperties().setProperties({
    SHEET_ID: 'PASTE_GOOGLE_SHEET_ID_HERE',
    BRIDGE_TOKEN: 'PASTE_LONG_RANDOM_TOKEN_HERE',
    SHEET_NAME: DEFAULT_SHEET_NAME
  }, true);
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing JSON body.');
  }
  return JSON.parse(e.postData.contents);
}

function getOrCreateSheet_(sheetId, sheetName) {
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function ensureHeader_(sheet) {
  const headers = [
    'MessageId',
    'ReceivedTime',
    'FromName',
    'FromEmail',
    'Subject',
    'BodyPreview',
    'HasAttachments',
    'ConversationId',
    'WebLink',
    'ImportedAt'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return;
  }

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isHeaderPresent = headers.every((header, index) => current[index] === header);
  if (!isHeaderPresent) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function getExistingIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return new Set();
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return new Set(values.map((row) => String(row[0] || '')).filter(Boolean));
}

function clean_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).replace(/\u0000/g, '').slice(0, 50000);
}

function jsonResponse(payload, statusCode) {
  if (statusCode) {
    payload.statusCode = statusCode;
  }

  const output = ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);

  // Apps Script web apps do not let us reliably set HTTP status codes from ContentService.
  return output;
}
