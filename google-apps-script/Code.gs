var CONFIG = {
  SHEET_ID: '1Bs_37_XbkIZ0e7CsUFcbdZDlVjb_ggre4ptxho5tHjE',
  OFFICE_EMAIL: 'docs@billlayneinsurance.com',
  AGENCY_NAME: 'Bill Layne Insurance Agency',
  AGENCY_PHONE: '(336) 835-1993',
  AGENCY_WEBSITE: 'https://www.billlayneinsurance.com',
  AGENCY_ADDRESS: '1283 N Bridge St, Elkin, NC 28621',
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
  var carrier = p.company || 'Your Insurance Company';
  var policyType = p.policyType || 'Policy';
  var localTime = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "MMMM d, yyyy 'at' h:mm a");
  var subject = '\u2705 Cancellation Confirmed \u2014 ' + esc_(p.insuredName || '') + ' \u2014 ' + esc_(carrier);
  var logoUrl = CONFIG.LOGO_URL;

  var html = [
    '<!DOCTYPE html><html lang="en" xmlns="http://www.w3.org/1999/xhtml"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><meta name="x-apple-disable-message-reformatting"><title>Cancellation Confirmed</title>',
    '<style>body,table,td,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}table,td{mso-table-lspace:0;mso-table-rspace:0}img{-ms-interpolation-mode:bicubic;border:0;height:auto;line-height:100%;outline:none;text-decoration:none}body{margin:0;padding:0;width:100%!important;background-color:#f1f5f9}@media only screen and (max-width:620px){.email-container{width:100%!important;padding:0 12px!important}.hero-pad{padding:28px 20px 32px!important}.card-pad{padding:20px 16px!important}.btn-td{padding:14px 24px!important}.hero-date{font-size:28px!important}}</style>',
    '</head><body style="margin:0;padding:0;background-color:#f1f5f9;font-family:Arial,\'Helvetica Neue\',Helvetica,sans-serif;">',

    '<div style="display:none;font-size:1px;color:#f1f5f9;line-height:1px;max-height:0;max-width:0;opacity:0;overflow:hidden;">cancellation confirmed effective ' + esc_(p.cancelDate || '') + ' &#8212; your signed request is on file</div>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#f1f5f9" style="background-color:#f1f5f9;"><tr><td align="center" style="padding:24px 16px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="600" class="email-container" style="width:600px;max-width:600px;margin:0 auto;">',

    // ── CARD 1: HEADER ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#fafafa;border-radius:16px 16px 0 0;border:1px solid #e2e8f0;border-bottom:none;">',
    '<tr><td style="height:4px;background-color:#003f87;font-size:0;line-height:0;border-radius:16px 16px 0 0;">&nbsp;</td></tr>',
    '<tr><td style="padding:20px 24px;" class="card-pad">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>',
    '<td align="center" valign="middle">',
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;border-radius:8px;padding:8px 12px;">',
    '<img src="' + logoUrl + '" width="160" alt="Bill Layne Insurance Agency" style="display:block;width:160px;max-width:160px;height:auto;">',
    '</td></tr></table></td>',
    '</tr></table></td></tr>',
    '<tr><td style="padding:0 24px 14px;text-align:center;">',
    '<p style="margin:0;font-size:11px;color:#64748b;font-family:Arial,sans-serif;letter-spacing:0.3px;">Cancellation Confirmation &bull; ' + esc_(carrier) + ' &bull; Bill Layne Insurance Agency</p>',
    '</td></tr></table></td></tr>',

    // ── CARD 2: HERO BLUE ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border:1px solid #e2e8f0;border-top:none;border-bottom:none;">',
    '<tr><td class="hero-pad" style="padding:36px 32px 40px;background-color:#003f87;text-align:center;">',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 16px;"><tr><td style="background-color:#0058b9;border-radius:20px;padding:5px 16px;">',
    '<span style="font-size:11px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;">&#128203; Cancellation Confirmed</span>',
    '</td></tr></table>',

    '<p style="margin:0 0 6px;font-size:13px;font-weight:600;color:rgba(255,255,255,0.80);font-family:Arial,sans-serif;letter-spacing:0.5px;text-transform:uppercase;">Cancelled at Insured\'s Request</p>',
    '<p style="margin:0 0 20px;font-size:28px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;line-height:1.2;">' + esc_(p.insuredName || '') + '</p>',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr>',
    '<td style="background-color:#0058b9;border-radius:16px;padding:20px 40px;border:1px solid rgba(255,255,255,0.30);text-align:center;">',
    '<p style="margin:0 0 4px;font-size:11px;font-weight:700;color:#C8A84E;font-family:Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;">Cancellation Effective</p>',
    '<p class="hero-date" style="margin:0;font-size:32px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;line-height:1.1;">' + esc_(p.cancelDate || '') + '</p>',
    '<p style="margin:6px 0 0;font-size:13px;color:rgba(255,255,255,0.75);font-family:Arial,sans-serif;">' + esc_(carrier) + ' &bull; ' + esc_(policyType) + ' &bull; Policy ' + esc_(p.policyNumber || '') + '</p>',
    '</td></tr></table>',

    '</td></tr></table></td></tr>',

    // ── CARD 3: BODY + DETAILS ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;border:1px solid #e2e8f0;border-top:none;border-bottom:none;">',
    '<tr><td style="padding:28px 28px 0;" class="card-pad">',
    '<p style="margin:0 0 16px;font-size:15px;color:#334155;font-family:Arial,sans-serif;line-height:1.65;">' + esc_(firstName) + ', your ' + esc_(carrier) + ' ' + esc_(policyType) + ' policy has been cancelled at your request. Your signed cancellation form is on file and your coverage terminates as of ' + esc_(p.cancelDate || '') + ' at ' + esc_(p.cancelTime || '12:01 AM') + '.</p>',
    '</td></tr>',

    '<tr><td style="padding:0 28px 24px;" class="card-pad">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0f9ff;border-radius:12px;border:1px solid #bae6fd;">',
    '<tr><td style="padding:18px 20px;">',
    '<p style="margin:0 0 12px;font-size:10px;font-weight:700;color:#0369a1;font-family:Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;">Cancellation Details</p>',

    detailRow_('Confirmation #', confirmNum, true),
    detailRow_('Insured', p.insuredName, false),
    detailRow_('Policy Number', p.policyNumber, false),
    detailRow_('Insurance Company', carrier, false),
    detailRow_('Policy Type', policyType, false),

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:8px;"><tr><td style="height:1px;background-color:#e2e8f0;font-size:0;line-height:0;">&nbsp;</td></tr></table>',

    detailRow_('Cancel Date', p.cancelDate, true),
    detailRow_('Cancel Time', p.cancelTime, false),
    detailRow_('Reason', p.reason, false),
    detailRowLast_('Signed At', p.signatureDateTime || localTime),

    '</td></tr></table></td></tr></table></td></tr>',

    // ── CARD 4: IMPORTANT DISCLOSURES ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#fafafa" style="background-color:#fafafa;border:1px solid #e2e8f0;border-top:none;border-bottom:none;">',
    '<tr><td style="padding:24px 28px;" class="card-pad">',
    '<p style="margin:0 0 4px;font-size:10px;font-weight:700;color:#64748b;font-family:Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;">Important Information</p>',
    '<p style="margin:0 0 16px;font-size:20px;font-weight:700;color:#0f172a;font-family:Arial,sans-serif;">What You Should Know</p>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#fffbeb;border-radius:12px;border:1px solid #fde68a;margin-bottom:12px;">',
    '<tr><td style="padding:16px 20px;">',
    '<p style="margin:0 0 6px;font-size:13px;font-weight:700;color:#92400e;font-family:Arial,sans-serif;">&#9888;&#65039; Coverage Gap Warning</p>',
    '<p style="margin:0;font-size:13px;color:#78350f;font-family:Arial,sans-serif;line-height:1.55;">Any period without insurance coverage may result in higher premiums when you obtain new coverage. Most carriers apply a surcharge for lapses of 30 days or more.</p>',
    '</td></tr></table>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0fdf4;border-radius:12px;border:1px solid #bbf7d0;">',
    '<tr><td style="padding:16px 20px;">',
    '<p style="margin:0 0 6px;font-size:13px;font-weight:700;color:#166534;font-family:Arial,sans-serif;">&#128176; Premium Refund</p>',
    '<p style="margin:0;font-size:13px;color:#14532d;font-family:Arial,sans-serif;line-height:1.55;">If any unearned premium is owed, ' + esc_(carrier) + ' will process your refund within 15 business days. Contact us at (336) 835-1993 if you haven\'t received it within that timeframe.</p>',
    '</td></tr></table>',

    '</td></tr></table></td></tr>',

    // ── CARD 5: SIGN-OFF + CTA ──
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;border:1px solid #e2e8f0;border-top:none;border-bottom:none;">',
    '<tr><td style="padding:24px 28px;" class="card-pad">',

    '<p style="margin:0 0 4px;font-size:15px;color:#334155;font-family:Arial,sans-serif;line-height:1.6;">Thanks in advance,</p>',
    '<p style="margin:0 0 16px;font-size:15px;font-weight:700;color:#0f172a;font-family:Arial,sans-serif;">&mdash; Bill Layne</p>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:20px;"><tr><td style="padding:0 4px;text-align:center;">',
    '<p style="margin:0;font-size:13px;color:#64748b;font-family:Arial,sans-serif;line-height:1.6;font-style:italic;">If you change your mind or need coverage again in the future, this email comes straight to me &#8212; just reply.</p>',
    '</td></tr></table>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:12px;"><tr><td align="center">',
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td style="border:2px solid #003f87;border-radius:12px;padding:14px 36px;">',
    '<a href="mailto:Save@BillLayneInsurance.com?subject=Cancellation%20Received%20-%20' + encodeURIComponent(p.insuredName || '') + '" style="display:block;font-size:15px;font-weight:700;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;text-align:center;">Reply to Confirm Receipt</a>',
    '</td></tr></table></td></tr></table>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td align="center">',
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td class="btn-td" style="background-color:#003f87;border-radius:12px;padding:14px 36px;">',
    '<a href="tel:3368351993" style="display:block;font-size:15px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;text-decoration:none;text-align:center;">Call (336) 835-1993</a>',
    '</td></tr></table></td></tr></table>',

    '</td></tr></table></td></tr>',

    // ── CARD 6: DISCLAIMER ──
    '<tr><td style="padding-bottom:0;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#fafafa" style="background-color:#fafafa;border:1px solid #e2e8f0;border-top:none;border-bottom:none;">',
    '<tr><td style="padding:20px 28px;" class="card-pad">',
    '<p style="margin:0;font-size:11px;color:#94a3b8;font-family:Arial,sans-serif;line-height:1.6;">This confirmation is issued on behalf of ' + esc_(carrier) + ' and documents the cancellation of the above-referenced policy at the insured\'s request. Coverage terminates ' + esc_(p.cancelDate || '') + ' at ' + esc_(p.cancelTime || '12:01 AM') + ' per the terms of the policy and applicable North Carolina law. This notice has been sent to the insured at the email address of record.<br><br>Bill Layne Insurance Agency &mdash; Licensed NC Property &amp; Casualty Agent<br>1283 N Bridge St, Elkin, NC 28621 &bull; (336) 835-1993</p>',
    '</td></tr></table></td></tr>',

    // ── FOOTER ──
    '<tr><td>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#fafafa" style="background-color:#fafafa;border-radius:0 0 16px 16px;border:1px solid #e2e8f0;border-top:none;">',
    '<tr><td style="padding:28px 24px;text-align:center;" class="card-pad">',

    '<table cellpadding="0" cellspacing="0" border="0" width="60" style="margin:0 auto 20px auto;"><tr><td style="height:3px;background-color:#003f87;font-size:0;line-height:0;">&nbsp;</td></tr></table>',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 12px auto;"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;border-radius:8px;padding:8px 12px;">',
    '<img src="' + logoUrl + '" width="140" alt="Bill Layne Insurance Agency" style="display:block;width:140px;max-width:140px;height:auto;">',
    '</td></tr></table>',

    '<p style="margin:0 0 4px;font-size:14px;font-weight:700;color:#0f172a;font-family:Arial,sans-serif;">Bill Layne Insurance Agency</p>',
    '<p style="margin:0 0 2px;font-size:12px;color:#64748b;font-family:Arial,sans-serif;">1283 N Bridge St &bull; Elkin, NC 28621</p>',
    '<p style="margin:0 0 2px;font-size:12px;color:#64748b;font-family:Arial,sans-serif;"><a href="tel:3368351993" style="color:#64748b;text-decoration:none;">(336) 835-1993</a> &bull; <a href="mailto:Save@BillLayneInsurance.com" style="color:#64748b;text-decoration:none;">Save@BillLayneInsurance.com</a></p>',
    '<p style="margin:0 0 14px;font-size:12px;color:#64748b;font-family:Arial,sans-serif;"><a href="https://www.BillLayneInsurance.com" style="color:#64748b;text-decoration:none;">www.BillLayneInsurance.com</a> &bull; Est. 2005</p>',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 14px auto;"><tr>',
    '<td style="padding:0 6px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#e2e8f0;border-radius:6px;padding:6px 12px;"><a href="https://www.facebook.com/dollarbillagency" style="font-size:11px;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;font-weight:700;">Facebook</a></td></tr></table></td>',
    '<td style="padding:0 6px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#e2e8f0;border-radius:6px;padding:6px 12px;"><a href="https://www.youtube.com/@ncautoandhome" style="font-size:11px;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;font-weight:700;">YouTube</a></td></tr></table></td>',
    '<td style="padding:0 6px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#e2e8f0;border-radius:6px;padding:6px 12px;"><a href="https://www.instagram.com/ncautoandhome" style="font-size:11px;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;font-weight:700;">Instagram</a></td></tr></table></td>',
    '<td style="padding:0 6px;"><table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#e2e8f0;border-radius:6px;padding:6px 12px;"><a href="https://twitter.com/shopsavecompare" style="font-size:11px;color:#003f87;font-family:Arial,sans-serif;text-decoration:none;font-weight:700;">X</a></td></tr></table></td>',
    '</tr></table>',

    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 14px auto;"><tr><td style="background-color:#f8fafc;border-radius:8px;padding:8px 14px;border:1px solid #e2e8f0;">',
    '<p style="margin:0;font-size:12px;font-weight:700;color:#0f172a;font-family:Arial,sans-serif;">4.9 &#11088;&#11088;&#11088;&#11088;&#11088; <span style="font-weight:400;color:#64748b;">100+ Google Reviews</span></p>',
    '</td></tr></table>',

    '<p style="margin:0;font-size:11px;color:#94a3b8;font-family:Arial,sans-serif;text-align:center;">You\'re receiving this because you requested cancellation of your ' + esc_(carrier) + ' policy.<br>&copy; 2026 Bill Layne Insurance Agency. All rights reserved.</p>',

    '</td></tr></table></td></tr>',

    '</table></td></tr></table>',
    '</body></html>'
  ].join('');

  var plain = [
    'Cancellation Confirmed - ' + (p.insuredName || ''),
    '',
    firstName + ', your ' + carrier + ' ' + policyType + ' policy has been cancelled at your request.',
    'Coverage terminates ' + (p.cancelDate || '') + ' at ' + (p.cancelTime || '12:01 AM') + '.',
    '',
    '--- CANCELLATION DETAILS ---',
    'Confirmation #: ' + confirmNum,
    'Insured: ' + (p.insuredName || ''),
    'Policy Number: ' + (p.policyNumber || ''),
    'Insurance Company: ' + carrier,
    'Policy Type: ' + policyType,
    'Cancel Date: ' + (p.cancelDate || ''),
    'Cancel Time: ' + (p.cancelTime || ''),
    'Reason: ' + (p.reason || ''),
    '',
    'WARNING: Any period without coverage may result in higher premiums.',
    'If a refund is owed, ' + carrier + ' will process within 15 business days.',
    '',
    'Questions? Call (336) 835-1993 or reply to this email.',
    '',
    '-- Bill Layne',
    'Bill Layne Insurance Agency',
    '1283 N Bridge St, Elkin, NC 28621',
    '(336) 835-1993 | Save@BillLayneInsurance.com'
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

function detailRow_(label, value, bold) {
  var w = bold ? 'font-weight:700;' : '';
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:8px;"><tr><td style="font-size:13px;color:#64748b;font-family:Arial,sans-serif;">' + esc_(label) + '</td><td align="right" style="font-size:13px;' + w + 'color:#0f172a;font-family:Arial,sans-serif;">' + esc_(String(value || '')) + '</td></tr></table>';
}

function detailRowLast_(label, value) {
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-size:13px;color:#64748b;font-family:Arial,sans-serif;">' + esc_(label) + '</td><td align="right" style="font-size:13px;color:#0f172a;font-family:Arial,sans-serif;">' + esc_(String(value || '')) + '</td></tr></table>';
}

function esc_(value) {
  return String(value || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
