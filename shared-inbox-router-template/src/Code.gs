const CONFIG = {
  ROUTE_A_DESTINATION: 'route-a@example.com',
  ROUTE_B_DESTINATION: 'route-b@example.com',
  ERROR_REPORT_EMAIL: 'alerts@example.com',

  LOG_SPREADSHEET_ID: 'REPLACE_WITH_SPREADSHEET_ID',
  RUN_LOG_SHEET: 'Run Log',
  FORWARD_LOG_SHEET: 'Forward Log',
  ERROR_LOG_SHEET: 'Error Log',

  // Blockers (add/remove words here)
  ROUTE_A_SUBJECT_BLOCKERS: [],
  ROUTE_A_BODY_BLOCKERS: [],
  ROUTE_B_SUBJECT_BLOCKERS: [],
  ROUTE_B_BODY_BLOCKERS: [],

  MAX_FORWARDS_PER_RUN: 10,

  ALLOWED_MAILBOXES: [
    'shared-mailbox-1@example.com',
    'shared-mailbox-2@example.com'
  ],

  SEARCH_QUERY: 'in:inbox newer_than:5d',
  MAX_THREADS_PER_RUN: 250,

  PROCESSED_KEY_PREFIX: 'processed_',
  ERROR_LOG_KEY: 'error_log',
  PROCESSED_RETENTION_DAYS: 45,

  ROUTE_B_KEYWORDS: [
    'lease',
    'agreement',
    'signature'
  ],

  ROUTE_A_KEYWORDS: [
    'overdue payment',
    'balance due',
    'final statement'
  ]
};

const GLOBAL_ROUTER_LOG_SS_ID = 'REPLACE_WITH_GLOBAL_LOG_SPREADSHEET_ID';
const GLOBAL_ROUTER_LOG_SHEET_NAME = 'Inbox Router Executions';

function hasRouteABodyBlocker_(msg) {
  const body = getMessageBodyForBlocker_(msg);
  return matchesAnyBlocker_(body, CONFIG.ROUTE_A_BODY_BLOCKERS);
}

function hasRouteASubjectBlocker_(subject) {
  return matchesAnyBlocker_(subject, CONFIG.ROUTE_A_SUBJECT_BLOCKERS);
}

function hasRouteBBodyBlocker_(msg) {
  const body = getMessageBodyForBlocker_(msg);
  return matchesAnyBlocker_(body, CONFIG.ROUTE_B_BODY_BLOCKERS);
}

function hasRouteBSubjectBlocker_(subject) {
  return matchesAnyBlocker_(subject, CONFIG.ROUTE_B_SUBJECT_BLOCKERS);
}

function shouldBlockByDestination_(destination, subject, msg) {
  const dest = String(destination || '').toLowerCase();

  if (dest === CONFIG.ROUTE_A_DESTINATION.toLowerCase()) {
    return hasRouteASubjectBlocker_(subject) || hasRouteABodyBlocker_(msg);
  }

  if (dest === CONFIG.ROUTE_B_DESTINATION.toLowerCase()) {
    return hasRouteBSubjectBlocker_(subject) || hasRouteBBodyBlocker_(msg);
  }

  return false;
}

function matchesAnyBlocker_(text, blockers) {
  const list = Array.isArray(blockers) ? blockers : [];
  return list.some(word => containsIgnoreCase_(String(text || ''), String(word || '')));
}

function getMessageBodyForBlocker_(msg) {
  const plain = safeGetPlainBody_(msg);
  const htmlAsText = safeGetHtmlBody_(msg).replace(/<[^>]+>/g, ' ');
  return (plain + ' ' + htmlAsText).trim();
}

function runEmailRouter() {
  const functionName = 'runEmailRouter';
  const startedAt = new Date();

  const lock = LockService.getUserLock();
  let lockAcquired = false;

  try {
    lock.waitLock(30000);
    lockAcquired = true;
  } catch (lockError) {
    logGlobalRouterExecution_(
      functionName,
      startedAt,
      'ERROR',
      'Could not acquire user lock: ' + String(lockError)
    );
    throw lockError;
  }

  let mailbox = '';
  let threadsScanned = 0;
  let messagesScanned = 0;
  let forwardedToRouteA = 0;
  let forwardedToRouteB = 0;
  let skippedNoMatch = 0;
  let skippedAlreadyProcessed = 0;
  let blockedByBlocker = 0;
  let errors = 0;
  let runStatus = 'OK';
  let forwardsThisRun = 0;
  let capReached = false;

  try {
    mailbox = Session.getActiveUser().getEmail().toLowerCase().trim();

    if (!CONFIG.ALLOWED_MAILBOXES.includes(mailbox)) {
      throw new Error('Router not allowed in mailbox: ' + mailbox);
    }

    ensureLogSheets_();

    let pageToken = null;

    do {
      if (capReached) break;

      const listResponse = Gmail.Users.Messages.list('me', {
        q: CONFIG.SEARCH_QUERY,
        maxResults: CONFIG.MAX_THREADS_PER_RUN,
        pageToken: pageToken
      });

      const messageRefs = listResponse.messages || [];
      pageToken = listResponse.nextPageToken || null;

      threadsScanned += messageRefs.length;

      for (const ref of messageRefs) {
        if (capReached) break;

        try {
          const messageId = ref.id;
          messagesScanned++;

          if (wasProcessed_(messageId)) {
            skippedAlreadyProcessed++;
            continue;
          }

          const meta = Gmail.Users.Messages.get('me', messageId, {
            format: 'metadata',
            metadataHeaders: ['Subject', 'From', 'To']
          });

          const headers = headersToMap_(meta.payload && meta.payload.headers ? meta.payload.headers : []);
          const subject = headers.subject || '';
          const from = extractEmail_(headers.from || '');
          const toHeader = headers.to || '';
          const threadId = meta.threadId || '';

          let destination = null;
          let matchedRule = '';

          if (
            matchesKeyword_(subject, CONFIG.ROUTE_A_KEYWORDS) &&
            from !== CONFIG.ROUTE_A_DESTINATION.toLowerCase()
          ) {
            destination = CONFIG.ROUTE_A_DESTINATION;
            matchedRule = 'ROUTE_A_KEYWORD';
          }

          if (
            !destination &&
            hasCityPoEndingA_(toHeader) &&
            matchesKeyword_(subject, CONFIG.ROUTE_B_KEYWORDS) &&
            from !== CONFIG.ROUTE_B_DESTINATION.toLowerCase()
          ) {
            destination = CONFIG.ROUTE_B_DESTINATION;
            matchedRule = 'CITY_PO_A + ROUTE_B_KEYWORD';
          }

          let msg = null;

          if (
            destination &&
            (
              destination.toLowerCase() === CONFIG.ROUTE_A_DESTINATION.toLowerCase() ||
              destination.toLowerCase() === CONFIG.ROUTE_B_DESTINATION.toLowerCase()
            )
          ) {
            msg = GmailApp.getMessageById(messageId);

            if (shouldBlockByDestination_(destination, subject, msg)) {
              blockedByBlocker++;
              continue;
            }
          }

          if (!destination) {
            skippedNoMatch++;
            continue;
          }

          if (forwardsThisRun >= CONFIG.MAX_FORWARDS_PER_RUN) {
            capReached = true;
            break;
          }

          const cityPos = extractAllCityPosFromHeader_(toHeader);
          const forwardSubject = buildForwardSubject_(cityPos, subject);

          msg = msg || GmailApp.getMessageById(messageId);
          const forwardPlainText = buildForwardPlainText_(toHeader, msg);
          const forwardHtml = buildForwardHtml_(toHeader, msg);

          msg.forward(destination, {
            subject: forwardSubject,
            body: forwardPlainText,
            htmlBody: forwardHtml
          });

          markProcessed_(messageId);

          if (msg.isUnread()) {
            msg.markRead();
          }

          appendForwardLog_({
            timestamp: new Date(),
            mailbox: mailbox,
            messageId: messageId,
            threadId: threadId,
            from: from,
            to: toHeader,
            subject: forwardSubject,
            destination: destination,
            matchedRule: matchedRule
          });

          if (destination.toLowerCase() === CONFIG.ROUTE_B_DESTINATION.toLowerCase()) {
            forwardedToRouteB++;
          } else if (destination.toLowerCase() === CONFIG.ROUTE_A_DESTINATION.toLowerCase()) {
            forwardedToRouteA++;
          }

          forwardsThisRun++;

        } catch (error) {
          errors++;
          runStatus = 'ERROR';

          logError_(error, {
            mailbox: mailbox,
            threadId: '',
            messageId: ref && ref.id ? ref.id : '',
            subject: 'MESSAGE-LEVEL ERROR'
          });
        }
      }

    } while (pageToken && !capReached);

    cleanupProcessedStore_();

  } catch (error) {
    errors++;
    runStatus = 'ERROR';

    try {
      logError_(error, {
        mailbox: mailbox || 'UNKNOWN',
        threadId: '',
        messageId: '',
        subject: 'RUN-LEVEL ERROR'
      });
    } catch (innerError) {
      console.error('Failed to persist run-level error: ' + innerError);
    }

    throw error;

  } finally {
    try {
      appendRunLog_({
        timestamp: new Date(),
        mailbox: mailbox || 'UNKNOWN',
        threadsScanned: threadsScanned,
        messagesScanned: messagesScanned,
        forwardedToRouteB: forwardedToRouteB,
        forwardedToRouteA: forwardedToRouteA,
        skippedNoMatch: skippedNoMatch,
        skippedAlreadyProcessed: skippedAlreadyProcessed,
        blockedByBlocker: blockedByBlocker,
        errors: errors,
        status: runStatus,
        startedAt: startedAt
      });
    } catch (logError) {
      console.error('Failed to append run log: ' + logError);
    }

    try {
      logGlobalRouterExecution_(
        functionName,
        startedAt,
        runStatus,
        'Mailbox: ' + (mailbox || 'UNKNOWN') +
        ' | Threads scanned: ' + threadsScanned +
        ' | Messages scanned: ' + messagesScanned +
        ' | Forwarded to Route B: ' + forwardedToRouteB +
        ' | Forwarded to Route A: ' + forwardedToRouteA +
        ' | Skipped no match: ' + skippedNoMatch +
        ' | Skipped already processed: ' + skippedAlreadyProcessed +
        ' | Blocker: ' + blockedByBlocker +
        ' | Errors: ' + errors +
        ' | Cap reached: ' + capReached
      );
    } catch (centralLogError) {
      console.error('Failed to append global router log: ' + centralLogError);
    }

    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function sendDailyErrorReport() {
  const props = PropertiesService.getUserProperties();
  const log = props.getProperty(CONFIG.ERROR_LOG_KEY);

  if (!log) return;

  const mailbox = Session.getActiveUser().getEmail();

  MailApp.sendEmail({
    to: CONFIG.ERROR_REPORT_EMAIL,
    subject: 'Email Router Errors - ' + mailbox,
    body:
      'Errors detected in router execution.\n\n' +
      'Mailbox: ' + mailbox + '\n' +
      'Date: ' + new Date().toLocaleString() + '\n\n' +
      log
  });

  props.deleteProperty(CONFIG.ERROR_LOG_KEY);
}

function appendRunLog_(data) {
  const sheet = getLogSheet_(CONFIG.RUN_LOG_SHEET);

  sheet.appendRow([
    data.timestamp,
    data.mailbox,
    data.threadsScanned,
    data.messagesScanned,
    data.forwardedToRouteB,
    data.forwardedToRouteA,
    data.skippedNoMatch,
    data.skippedAlreadyProcessed,
    data.blockedByBlocker,
    data.errors,
    data.status,
    data.startedAt
  ]);
}

function appendForwardLog_(data) {
  const sheet = getLogSheet_(CONFIG.FORWARD_LOG_SHEET);

  sheet.appendRow([
    data.timestamp,
    data.mailbox,
    data.messageId,
    data.threadId,
    data.from,
    data.to,
    data.subject,
    data.destination,
    data.matchedRule
  ]);
}

function appendErrorLog_(data) {
  const sheet = getLogSheet_(CONFIG.ERROR_LOG_SHEET);

  sheet.appendRow([
    data.timestamp,
    data.mailbox,
    data.threadId,
    data.messageId,
    data.subject,
    data.errorText
  ]);
}

function ensureLogSheets_() {
  const ss = SpreadsheetApp.openById(CONFIG.LOG_SPREADSHEET_ID);

  ensureSheetWithHeaders_(ss, CONFIG.RUN_LOG_SHEET, [
    'Timestamp',
    'Mailbox',
    'Threads Scanned',
    'Messages Scanned',
    'Forwarded to Route B',
    'Forwarded to Route A',
    'Skipped No Match',
    'Skipped Already Processed',
    'Blocker',
    'Errors',
    'Status',
    'Run Started At'
  ]);

  ensureSheetWithHeaders_(ss, CONFIG.FORWARD_LOG_SHEET, [
    'Timestamp',
    'Mailbox',
    'Message ID',
    'Thread ID',
    'From',
    'To',
    'Subject',
    'Destination',
    'Rule'
  ]);

  ensureSheetWithHeaders_(ss, CONFIG.ERROR_LOG_SHEET, [
    'Timestamp',
    'Mailbox',
    'Thread ID',
    'Message ID',
    'Subject',
    'Error'
  ]);
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const currentHeaders = lastColumn > 0
    ? sheet.getRange(1, 1, 1, Math.max(lastColumn, headers.length)).getValues()[0]
    : [];

  const normalizedCurrent = currentHeaders.slice(0, headers.length).map(String);
  const normalizedExpected = headers.map(String);

  if (JSON.stringify(normalizedCurrent) !== JSON.stringify(normalizedExpected)) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function getLogSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.LOG_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Log sheet not found: ' + sheetName);
  }

  return sheet;
}

function logError_(error, meta) {
  const errorText = error && error.stack ? error.stack : String(error);

  appendErrorLog_({
    timestamp: new Date(),
    mailbox: meta.mailbox || '',
    threadId: meta.threadId || '',
    messageId: meta.messageId || '',
    subject: meta.subject || '',
    errorText: errorText
  });

  const props = PropertiesService.getUserProperties();
  const existing = props.getProperty(CONFIG.ERROR_LOG_KEY) || '';

  const entry =
    new Date().toISOString() +
    ' | mailbox:' + (meta.mailbox || '') +
    ' | thread:' + (meta.threadId || '') +
    ' | message:' + (meta.messageId || '') +
    ' | subject:' + (meta.subject || '') +
    ' | ' + errorText;

  props.setProperty(CONFIG.ERROR_LOG_KEY, existing ? existing + '\n' + entry : entry);
}

function hasCityPoEndingA_(toHeader) {
  const cityPos = extractAllCityPosFromHeader_(toHeader);
  return cityPos.some(cityPo => /A$/i.test(cityPo));
}

function extractAllCityPosFromHeader_(headerValue) {
  const emails = extractAllEmails_(headerValue);
  const cityPos = [];
  const seen = {};

  for (const email of emails) {
    const localPart = (email.split('@')[0] || '').trim();

    if (/^[A-Z]{3}-\d+[A-Z]?$/i.test(localPart)) {
      const normalized = localPart.toUpperCase();
      if (!seen[normalized]) {
        seen[normalized] = true;
        cityPos.push(normalized);
      }
    }
  }

  return cityPos;
}

function buildForwardSubject_(cityPos, originalSubject) {
  const cleanSubject = String(originalSubject || '').trim();
  const safeCityPos = Array.isArray(cityPos) ? cityPos.filter(Boolean) : [];

  if (safeCityPos.length > 0) {
    return '[' + safeCityPos.join(', ') + '] ' + cleanSubject;
  }

  return cleanSubject;
}

function buildForwardPlainText_(toHeader, msg) {
  const lines = [];

  lines.push('Original to: ' + String(toHeader || '').trim());

  const from = safeGetFrom_(msg);
  if (from) {
    lines.push('From: ' + from);
  }

  const cc = safeGetCc_(msg);
  if (cc) {
    lines.push('Cc: ' + cc);
  }

  const bcc = safeGetBcc_(msg);
  if (bcc) {
    lines.push('Bcc: ' + bcc);
  }

  const headerBlock = lines.join('\n');
  const originalPlainBody = safeGetPlainBody_(msg);

  return headerBlock + '\n\n' + originalPlainBody;
}

function buildForwardHtml_(toHeader, msg) {
  const rows = [];

  rows.push(
    '<div><strong>Original to:</strong> ' +
    escapeHtml_(String(toHeader || '').trim()) +
    '</div>'
  );

  const from = safeGetFrom_(msg);
  if (from) {
    rows.push(
      '<div><strong>From:</strong> ' +
      escapeHtml_(from) +
      '</div>'
    );
  }

  const cc = safeGetCc_(msg);
  if (cc) {
    rows.push(
      '<div><strong>Cc:</strong> ' +
      escapeHtml_(cc) +
      '</div>'
    );
  }

  const bcc = safeGetBcc_(msg);
  if (bcc) {
    rows.push(
      '<div><strong>Bcc:</strong> ' +
      escapeHtml_(bcc) +
      '</div>'
    );
  }

  const originalHtmlBody = safeGetHtmlBody_(msg);

  return (
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5;margin-bottom:16px;">' +
      rows.join('') +
    '</div>' +
    originalHtmlBody
  );
}

function safeGetPlainBody_(msg) {
  try {
    return msg.getPlainBody() || '';
  } catch (e) {
    return '';
  }
}

function safeGetHtmlBody_(msg) {
  try {
    return msg.getBody() || '';
  } catch (e) {
    return '<div></div>';
  }
}

function safeGetFrom_(msg) {
  try {
    return msg.getFrom() || '';
  } catch (e) {
    return '';
  }
}

function safeGetCc_(msg) {
  try {
    return msg.getCc() || '';
  } catch (e) {
    return '';
  }
}

function safeGetBcc_(msg) {
  try {
    return msg.getBcc() || '';
  } catch (e) {
    return '';
  }
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function matchesKeyword_(text, keywords) {
  const value = normalizeForKeywordMatch_(text);

  return keywords.some(function(keyword) {
    const normalizedKeyword = normalizeForKeywordMatch_(keyword);
    const pattern = new RegExp('(^|[^a-z0-9])' + escapeRegex_(normalizedKeyword) + '([^a-z0-9]|$)', 'i');
    return pattern.test(value);
  });
}

function containsIgnoreCase_(text, term) {
  const value = normalizeForKeywordMatch_(text);
  const normalizedTerm = normalizeForKeywordMatch_(term);
  const pattern = new RegExp('(^|[^a-z0-9])' + escapeRegex_(normalizedTerm) + '([^a-z0-9]|$)', 'i');
  return pattern.test(value);
}

function normalizeForKeywordMatch_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function escapeRegex_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function extractEmail_(headerValue) {
  const value = String(headerValue || '');
  const match = value.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? match[0].toLowerCase() : value.toLowerCase().trim();
}

function extractAllEmails_(headerValue) {
  const value = String(headerValue || '');
  const matches = value.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig);
  return (matches || []).map(email => email.toLowerCase());
}

function headersToMap_(headers) {
  const map = {};
  headers.forEach(function(header) {
    map[String(header.name || '').toLowerCase()] = header.value || '';
  });
  return map;
}

function wasProcessed_(messageId) {
  const props = PropertiesService.getUserProperties();
  return !!props.getProperty(CONFIG.PROCESSED_KEY_PREFIX + messageId);
}

function markProcessed_(messageId) {
  const props = PropertiesService.getUserProperties();
  props.setProperty(
    CONFIG.PROCESSED_KEY_PREFIX + messageId,
    String(Date.now())
  );
}

function cleanupProcessedStore_() {
  const props = PropertiesService.getUserProperties();
  const allProps = props.getProperties();
  const now = Date.now();
  const maxAgeMs = CONFIG.PROCESSED_RETENTION_DAYS * 24 * 60 * 60 * 1000;

  Object.keys(allProps).forEach(key => {
    if (!key.startsWith(CONFIG.PROCESSED_KEY_PREFIX)) return;

    const timestamp = Number(allProps[key]);
    if (!timestamp || (now - timestamp) > maxAgeMs) {
      props.deleteProperty(key);
    }
  });
}

function logGlobalRouterExecution_(functionName, executionStart, status, comment) {
  try {
    const executionEnd = new Date();
    const durationSec = Math.round(((executionEnd.getTime() - executionStart.getTime()) / 1000) * 100) / 100;

    const ss = SpreadsheetApp.openById(GLOBAL_ROUTER_LOG_SS_ID);
    let sheet = ss.getSheetByName(GLOBAL_ROUTER_LOG_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(GLOBAL_ROUTER_LOG_SHEET_NAME);
      sheet.appendRow(['Function', 'Timestamp', 'Status', 'Duration (sec)', 'Comment']);
    }

    sheet.appendRow([
      functionName,
      executionEnd,
      status,
      durationSec,
      comment || ''
    ]);

  } catch (logErr) {
    Logger.log('Could not write global router log: ' + String(logErr));
  }
}

function createTriggers() {
  const handlers = ['runEmailRouter', 'sendDailyErrorReport'];

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (handlers.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runEmailRouter')
    .timeBased()
    .everyMinutes(10)
    .create();

  ScriptApp.newTrigger('sendDailyErrorReport')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}
