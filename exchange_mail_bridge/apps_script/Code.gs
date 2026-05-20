const DEFAULT_SHEET_NAME = 'Exchange Mail';
const DEFAULT_ATTACHMENT_FOLDER_NAME = 'Exchange Mail Attachments';
const DEFAULT_BODY_FOLDER_NAME = 'Exchange Mail Bodies';

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
    const headers = ensureHeader_(sheet);
    const columns = getColumnMap_(headers);

    const existingRows = getExistingIdRows_(sheet);
    const now = new Date().toISOString();
    const rows = [];
    const updates = [];
    let skipped = 0;
    let updated = 0;

    messages.forEach((message) => {
      const messageId = clean_(message.id);
      if (!messageId) {
        skipped += 1;
        return;
      }

      const rowNumber = existingRows.get(messageId);
      const existingAttachmentLinks = rowNumber && columns.AttachmentLinks
        ? clean_(sheet.getRange(rowNumber, columns.AttachmentLinks).getValue())
        : '';
      const hasStaleAttachmentLinks = Boolean(existingAttachmentLinks) && !validAttachmentLinks_(existingAttachmentLinks);
      const savedAttachmentLinks = hasStaleAttachmentLinks ? '' : existingAttachmentLinks;
      const attachmentLinks = savedAttachmentLinks || saveAttachments_(message, props);
      const existingBodyHtmlLink = rowNumber && columns.BodyHtmlLink
        ? clean_(sheet.getRange(rowNumber, columns.BodyHtmlLink).getValue())
        : '';
      const hasStaleBodyHtmlLink = Boolean(existingBodyHtmlLink) && !validAttachmentLinks_(existingBodyHtmlLink);
      const savedBodyHtmlLink = hasStaleBodyHtmlLink ? '' : existingBodyHtmlLink;
      const bodyHtmlLink = savedBodyHtmlLink || saveBodyHtml_(message, props);
      const row = buildRow_(message, attachmentLinks, bodyHtmlLink, now);

      if (rowNumber) {
        if (hasEnrichment_(message) || attachmentLinks || bodyHtmlLink || hasStaleAttachmentLinks || hasStaleBodyHtmlLink) {
          updates.push({ rowNumber, row });
          updated += 1;
        } else {
          skipped += 1;
        }
        return;
      }

      existingRows.set(messageId, sheet.getLastRow() + rows.length + 1);
      rows.push(row);
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    updates.forEach((update) => {
      sheet.getRange(update.rowNumber, 1, 1, update.row.length).setValues([update.row]);
    });

    return jsonResponse({ ok: true, appended: rows.length, updated, skipped });
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
    SHEET_NAME: DEFAULT_SHEET_NAME,
    ATTACHMENT_FOLDER_NAME: DEFAULT_ATTACHMENT_FOLDER_NAME,
    BODY_FOLDER_NAME: DEFAULT_BODY_FOLDER_NAME
  }, true);
}

function authorizeDriveOnce() {
  getOrCreateDriveFolder_(DEFAULT_ATTACHMENT_FOLDER_NAME);
  getOrCreateDriveFolder_(DEFAULT_BODY_FOLDER_NAME);
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
    'AttachmentCount',
    'AttachmentNames',
    'AttachmentLinks',
    'BodyImageUrls',
    'BodyHtmlLink',
    'BodyLinks',
    'BodyTextFull',
    'ConversationId',
    'WebLink',
    'ImportedAt'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return headers;
  }

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (current[0] === 'MessageId') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return headers;
  }

  const isHeaderPresent = headers.every((header, index) => current[index] === header);
  if (!isHeaderPresent) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return headers;
}

function getExistingIdRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return new Map();
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const rows = new Map();
  values.forEach((row, index) => {
    const messageId = String(row[0] || '');
    if (messageId) {
      rows.set(messageId, index + 2);
    }
  });
  return rows;
}

function getColumnMap_(headers) {
  return headers.reduce((columns, header, index) => {
    columns[header] = index + 1;
    return columns;
  }, {});
}

function buildRow_(message, attachmentLinks, bodyHtmlLink, importedAt) {
  const attachments = Array.isArray(message.attachments) ? message.attachments : [];
  return [
    clean_(message.id),
    clean_(message.receivedTime),
    clean_(message.fromName),
    clean_(message.fromEmail),
    clean_(message.subject),
    clean_(message.bodyPreview),
    Boolean(message.hasAttachments || attachments.length > 0),
    Number(message.attachmentCount || attachments.length || 0),
    clean_(attachmentNames_(attachments)),
    clean_(attachmentLinks),
    clean_(bodyImageUrls_(message)),
    clean_(bodyHtmlLink),
    clean_(bodyLinks_(message)),
    clean_(message.bodyTextFull),
    clean_(message.conversationId),
    clean_(message.webLink),
    importedAt
  ];
}

function hasEnrichment_(message) {
  const attachments = Array.isArray(message.attachments) ? message.attachments : [];
  const bodyImageUrls = Array.isArray(message.bodyImageUrls) ? message.bodyImageUrls : [];
  const bodyLinks = Array.isArray(message.bodyLinks) ? message.bodyLinks : [];
  return attachments.length > 0
    || bodyImageUrls.length > 0
    || bodyLinks.length > 0
    || Boolean(message.bodyHtml)
    || Boolean(message.bodyTextFull)
    || Number(message.attachmentCount || 0) > 0;
}

function saveAttachments_(message, props) {
  const attachments = Array.isArray(message.attachments) ? message.attachments : [];
  if (attachments.length === 0) {
    return '';
  }

  const folderName = props.getProperty('ATTACHMENT_FOLDER_NAME') || DEFAULT_ATTACHMENT_FOLDER_NAME;
  const folder = getOrCreateDriveFolder_(folderName);
  const links = [];

  attachments.forEach((attachment, index) => {
    if (!attachment.base64) {
      return;
    }

    const safeMessageId = clean_(message.id).replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 24) || 'message';
    const safeName = clean_(attachment.name || `attachment-${index + 1}`).replace(/[\\/:*?"<>|]/g, '_').slice(0, 120);
    const fileName = `${safeMessageId}_${index + 1}_${safeName}`;
    const bytes = Utilities.base64Decode(attachment.base64);
    const blob = Utilities.newBlob(bytes, clean_(attachment.contentType) || MimeType.BINARY, fileName);
    const file = folder.createFile(blob);
    links.push(file.getUrl());
  });

  return links.join('\n');
}

function saveBodyHtml_(message, props) {
  const html = cleanLarge_(message.bodyHtml, 200000);
  if (!html) {
    return '';
  }

  const folderName = props.getProperty('BODY_FOLDER_NAME') || DEFAULT_BODY_FOLDER_NAME;
  const folder = getOrCreateDriveFolder_(folderName);
  const safeMessageId = clean_(message.id).replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 24) || 'message';
  const safeSubject = clean_(message.subject || 'mail-body').replace(/[\\/:*?"<>|]/g, '_').slice(0, 80) || 'mail-body';
  const fileName = `${safeMessageId}_${safeSubject}.html`;
  const file = folder.createFile(fileName, html, MimeType.HTML);
  return file.getUrl();
}

function validAttachmentLinks_(value) {
  const text = clean_(value);
  if (!text) {
    return false;
  }
  return text.split(/\s+/).some((part) => /^https?:\/\//i.test(part));
}

function getOrCreateDriveFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function attachmentNames_(attachments) {
  return attachments
    .map((attachment, index) => [
      `attachment ${index + 1}`,
      attachment.name ? `name=${attachment.name}` : '',
      attachment.contentType ? `contentType=${attachment.contentType}` : '',
      attachment.isInline ? 'inline=true' : '',
      attachment.contentId ? `contentId=${attachment.contentId}` : '',
      attachment.skippedReason ? `skipped=${attachment.skippedReason}` : ''
    ].filter(Boolean).join(' | '))
    .join('\n');
}

function bodyImageUrls_(message) {
  const urls = Array.isArray(message.bodyImageUrls) ? message.bodyImageUrls : [];
  return urls.join('\n');
}

function bodyLinks_(message) {
  const links = Array.isArray(message.bodyLinks) ? message.bodyLinks : [];
  return links.join('\n');
}

function clean_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).replace(/\u0000/g, '').slice(0, 50000);
}

function cleanLarge_(value, limit) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).replace(/\u0000/g, '').slice(0, limit);
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
