var CONFIG = {
  SHEET_ID: '1Bs_37_XbkIZ0e7CsUFcbdZDlVjb_ggre4ptxho5tHjE',
  OFFICE_EMAIL: 'bill@billlayneinsurance.com',
  AGENCY_NAME: 'Bill Layne Insurance Agency',
  AGENCY_PHONE: '(336) 835-1993',
  AGENCY_WEBSITE: 'https://www.billlayneinsurance.com',
  AGENCY_ADDRESS: '127 CC Camp Rd, Elkin, NC 28621',
  LOGO_URL: 'https://i.imgur.com/lxu9nfT.png',
  TIMEZONE: 'America/New_York'
};

function doPost(e) {
  try {
    var p = e.parameter;
    var confirmNum = p.confirmationNumber || generateConfirmation_();

    logToSheet_(p, confirmNum);
    sendOfficeEmail_(p, confirmNum);

    if (p.insuredEmail) {
      sendCustomerEmail_(p, confirmNum);
    }

    return ContentService.createTextOutput('Success').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    console.error('Cancellation backend error: ' + (err.stack || err));
    return ContentService.createTextOutput('Error: ' + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    ok: true, app: 'BLI Cancellation Backend', version: '2026-04-01'
  })).setMimeType(ContentService.MimeType.JSON);
}

function generateConfirmation_() {
  var now = new Date();
  var d = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyMMdd');
  var r = Math.random().toString(36).slice(2, 6).toUpperCase();
  return 'CANCEL-' + d + '-' + r;
}

// ── SHEET LOGGING ──

function logToSheet_(p, confirmNum) {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName('Cancellations');

  if (!sheet) {
    sheet = ss.insertSheet('Cancellations');
    sheet.appendRow([
      'Timestamp', 'Confirmation #', 'Policy Number', 'Insured Name', 'Email', 'Phone',
      'Address', 'Company', 'Policy Type', 'Cancel Date', 'Cancel Time',
      'Reason', 'Notes', 'Entry Mode', 'Signature', 'Signature Date/Time'
    ]);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#003f87').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  var localTime = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "M/d/yyyy h:mm a");

  sheet.appendRow([
    localTime, confirmNum, p.policyNumber || '', p.insuredName || '', p.insuredEmail || '',
    p.insuredPhone || '', p.insuredAddress || '', p.company || '', p.policyType || '',
    p.cancelDate || '', p.cancelTime || '', p.reason || '', p.notes || '',
    p.entryMode || '', p.typedSignature || '', p.signatureDateTime || ''
  ]);
}

// ── OFFICE EMAIL ──

function sendOfficeEmail_(p, confirmNum) {
  var subject = '[CANCELLATION] ' + (p.insuredName || 'Unknown') + ' - Policy ' + (p.policyNumber || 'N/A') + ' (' + confirmNum + ')';

  var rows = '';
  rows += row_('Confirmation #', confirmNum);
  rows += row_('Insured Name', p.insuredName);
  rows += row_('Policy Number', p.policyNumber);
  rows += row_('Email', p.insuredEmail);
  rows += row_('Phone', p.insuredPhone);
  rows += row_('Address', p.insuredAddress);
  rows += row_('Company', p.company);
  rows += row_('Policy Type', p.policyType);
  rows += row_('Cancel Date', p.cancelDate);
  rows += row_('Cancel Time', p.cancelTime);
  rows += row_('Reason', p.reason);
  rows += row_('Notes', p.notes);
  rows += row_('Entry Mode', p.entryMode);
  rows += row_('Signature', p.typedSignature);
  rows += row_('Signed At', p.signatureDateTime);

  var html = '<div style="font-family:Arial,sans-serif;max-width:600px;"><h2 style="color:#dc2626;margin:0 0 16px;">Cancellation Request Received</h2><table style="width:100%;border-collapse:collapse;">' + rows + '</table></div>';

  MailApp.sendEmail({
    to: CONFIG.OFFICE_EMAIL,
    subject: subject,
    htmlBody: html,
    body: 'Cancellation request from ' + (p.insuredName || 'Unknown') + ' - Policy ' + (p.policyNumber || 'N/A') + ' - Confirmation: ' + confirmNum,
    replyTo: p.insuredEmail || CONFIG.OFFICE_EMAIL,
    name: CONFIG.AGENCY_NAME
  });
}

function row_(label, value) {
  if (!value && value !== 0) return '';
  return '<tr><td style="padding:6px 10px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#64748b;width:35%;vertical-align:top;">' + esc_(label) + '</td><td style="padding:6px 10px;border-bottom:1px solid #e2e8f0;font-size:14px;color:#0f2744;font-weight:500;">' + esc_(String(value)) + '</td></tr>';
}

// ── CUSTOMER CONFIRMATION EMAIL ──

function sendCustomerEmail_(p, confirmNum) {
  var firstName = (p.insuredName || 'there').split(' ')[0];
  var subject = '\u2705 Cancellation Request Received \u2014 ' + confirmNum + ' | Bill Layne Insurance';
  var localTime = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "MMMM d, yyyy 'at' h:mm a");

  var html = [
    '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>',
    '<body style="margin:0;padding:0;background-color:#f1f5f9;-webkit-text-size-adjust:100%;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:600px;margin:0 auto;">',

    '<tr><td style="padding:0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#003f87;background:linear-gradient(135deg,#003f87 0%,#0076d3 100%);border-radius:0 0 16px 16px;">',
    '<tr><td style="padding:36px 30px 28px;text-align:center;">',
    '<img src="' + CONFIG.LOGO_URL + '" alt="Bill Layne Insurance" width="180" height="45" style="display:block;margin:0 auto 16px;max-width:180px;height:auto;border:0;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr><td style="background-color:#1a5296;border-radius:20px;padding:6px 16px;"><span style="font-family:Arial,sans-serif;font-size:13px;color:#ffffff;">&#10003; Cancellation Request Received</span></td></tr></table>',
    '</td></tr></table></td></tr>',

    '<tr><td style="padding:20px 16px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#ffffff;border-radius:16px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">',

    '<tr><td style="padding:28px 28px 0;">',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:22px;font-weight:700;color:#0f2744;">Thank you, ' + esc_(firstName) + '.</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:15px;color:#64748b;line-height:1.6;">We have received your cancellation request and appreciate you letting us know. Our team will process this promptly and reach out if we need any additional information.</p>',
    '</td></tr>',

    '<tr><td style="padding:20px 28px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0f9ff;border-radius:12px;border:1px solid #bae6fd;">',
    '<tr><td style="padding:16px 20px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">',
    '<tr><td style="font-family:Arial,sans-serif;font-size:11px;font-weight:700;color:#0369a1;text-transform:uppercase;letter-spacing:0.5px;padding-bottom:8px;">Request Details</td></tr>',
    confRow_('Confirmation #', confirmNum, true),
    confRow_('Policy Number', p.policyNumber, false),
    confRow_('Cancel Date', p.cancelDate, false),
    confRow_('Submitted', localTime, false),
    '</table></td></tr></table></td></tr>',

    '<tr><td style="padding:24px 28px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0fdf4;border-radius:12px;border:1px solid #bbf7d0;">',
    '<tr><td style="padding:16px 20px;">',
    '<p style="margin:0 0 8px;font-family:Arial,sans-serif;font-size:13px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:0.5px;">What Happens Next</p>',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0">',
    '<tr><td style="padding:4px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128203; We review your cancellation request</td></tr>',
    '<tr><td style="padding:4px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128222; An agent will contact you to confirm the details</td></tr>',
    '<tr><td style="padding:4px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128196; You will receive written confirmation once processed</td></tr>',
    '<tr><td style="padding:4px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128274; Your request is securely stored for our records</td></tr>',
    '</table></td></tr></table></td></tr>',

    '<tr><td style="padding:24px 28px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#fff7ed;border-radius:12px;border:1px solid #fed7aa;">',
    '<tr><td style="padding:16px 20px;text-align:center;">',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:14px;font-weight:700;color:#9a3412;">Changed your mind or have questions?</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:14px;color:#78350f;">Call us at <a href="tel:3368351993" style="color:#0076d3;text-decoration:none;font-weight:700;">' + CONFIG.AGENCY_PHONE + '</a> before the cancellation date.</p>',
    '</td></tr></table></td></tr>',

    '<tr><td style="padding:24px 28px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="text-align:center;">',
    '<a href="' + CONFIG.AGENCY_WEBSITE + '" target="_blank" style="display:inline-block;background-color:#0076d3;color:#ffffff;font-family:Arial,sans-serif;font-size:15px;font-weight:700;text-decoration:none;padding:14px 36px;border-radius:12px;">Visit Our Website</a>',
    '</td></tr></table></td></tr>',

    '</table></td></tr>',

    '<tr><td style="padding:20px 16px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#0f172a;border-radius:16px;">',
    '<tr><td style="padding:28px 28px 20px;text-align:center;">',
    '<img src="' + CONFIG.LOGO_URL + '" alt="Bill Layne Insurance" width="140" height="35" style="display:block;margin:0 auto 12px;max-width:140px;height:auto;border:0;">',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:14px;color:#e2e8f0;">' + CONFIG.AGENCY_NAME + '</p>',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:13px;color:#94a3b8;">' + CONFIG.AGENCY_ADDRESS + '</p>',
    '<p style="margin:0 0 12px;font-family:Arial,sans-serif;font-size:13px;color:#94a3b8;"><a href="tel:3368351993" style="color:#60a5fa;text-decoration:none;">' + CONFIG.AGENCY_PHONE + '</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="mailto:bill@billlayneinsurance.com" style="color:#60a5fa;text-decoration:none;">bill@billlayneinsurance.com</a></p>',
    '</td></tr>',
    '<tr><td style="padding:0 28px 20px;text-align:center;"><p style="margin:0;font-family:Arial,sans-serif;font-size:11px;color:#475569;">&copy; 2026 Bill Layne Insurance Agency. All rights reserved.</p></td></tr>',
    '</table></td></tr>',

    '<tr><td style="padding:20px 0;">&nbsp;</td></tr>',
    '</table></body></html>'
  ].join('');

  var plain = [
    'Thank you, ' + firstName + '.',
    '',
    'We received your cancellation request and will process it promptly.',
    '',
    '--- REQUEST DETAILS ---',
    'Confirmation #: ' + confirmNum,
    'Policy Number: ' + (p.policyNumber || 'N/A'),
    'Cancel Date: ' + (p.cancelDate || 'N/A'),
    'Submitted: ' + localTime,
    '',
    '--- WHAT HAPPENS NEXT ---',
    '  1. We review your cancellation request',
    '  2. An agent will contact you to confirm details',
    '  3. You will receive written confirmation once processed',
    '',
    'Changed your mind? Call us at ' + CONFIG.AGENCY_PHONE,
    '',
    CONFIG.AGENCY_NAME,
    CONFIG.AGENCY_ADDRESS,
    CONFIG.AGENCY_PHONE
  ].join('\n');

  MailApp.sendEmail({
    to: p.insuredEmail,
    subject: subject,
    htmlBody: html,
    body: plain,
    replyTo: CONFIG.OFFICE_EMAIL,
    name: CONFIG.AGENCY_NAME
  });
}

function confRow_(label, value, bold) {
  var weight = bold ? 'font-weight:700;' : '';
  return '<tr><td style="padding-bottom:6px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-family:Arial,sans-serif;font-size:14px;color:#64748b;">' + esc_(label) + '</td><td style="font-family:Arial,sans-serif;font-size:14px;color:#0f2744;text-align:right;' + weight + '">' + esc_(String(value || '')) + '</td></tr></table></td></tr>';
}

function esc_(value) {
  return String(value || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
