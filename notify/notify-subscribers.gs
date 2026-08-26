/**
 * Google Apps Script: sender mails til gæster, der har sagt ja til opdateringer.
 *
 * OPSÆTNING (én gang):
 * 1. Åbn Google Sheets-arket med formularsvarene.
 * 2. Udvidelser -> Apps Script. Indsæt dette script.
 * 3. Ret SHEET_NAME, EMAIL_HEADER og UPDATES_HEADER, så de matcher kolonnerne i arket.
 * 4. Vælg et hemmeligt token og indsæt det i SECRET herunder.
 * 5. Udrul -> Ny udrulning -> Webapp:
 *    - Udfør som: Mig
 *    - Hvem har adgang: Alle
 *    Kopiér webapp-URL'en.
 * 6. Læg i GitHub-repoets Settings -> Secrets and variables -> Actions:
 *    - NOTIFY_URL    = webapp-URL'en
 *    - NOTIFY_SECRET = samme værdi som SECRET
 */

const SECRET = 'SKIFT-MIG-TIL-NOGET-HEMMELIGT';
const SHEET_NAME = 'Formularsvar 1';
const EMAIL_HEADER = 'email';
const UPDATES_HEADER = 'Tilmelding af opdateringer';
const SITE_URL = 'https://xn--schneflensted-dnb.dk';
const FROM_NAME = 'Ann-Katrine & Simon';

function doPost(e) {
  const params = JSON.parse(e.postData.contents || '{}');
  if (params.secret !== SECRET) {
    return ContentService.createTextOutput('Forbidden').setMimeType(ContentService.MimeType.TEXT);
  }

  const message = (params.message || '').trim();
  if (!message) {
    return ContentService.createTextOutput('Missing message').setMimeType(ContentService.MimeType.TEXT);
  }

  const recipients = getSubscriberEmails_();
  const subject = params.subject || 'Nyt på vores bryllupshjemmeside 💐';
  const body =
    'Hej!\n\n' +
    'Der er kommet opdateringer på vores bryllupshjemmeside:\n\n' +
    message + '\n\n' +
    'Se det hele her: ' + SITE_URL + '\n\n' +
    'Kærlig hilsen\n' + FROM_NAME;

  recipients.forEach(function (email) {
    MailApp.sendEmail({ to: email, subject: subject, body: body, name: FROM_NAME });
  });

  return ContentService
    .createTextOutput('Sent to ' + recipients.length + ' subscriber(s)')
    .setMimeType(ContentService.MimeType.TEXT);
}

function getSubscriberEmails_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const emailCol = headers.indexOf(EMAIL_HEADER);
  const updatesCol = headers.indexOf(UPDATES_HEADER);
  if (emailCol === -1 || updatesCol === -1) {
    throw new Error('Kolonner ikke fundet – tjek EMAIL_HEADER og UPDATES_HEADER.');
  }

  const seen = {};
  const emails = [];
  for (var i = 1; i < rows.length; i++) {
    const email = String(rows[i][emailCol]).trim().toLowerCase();
    const wantsUpdates = String(rows[i][updatesCol]).trim() === 'Ja';
    if (wantsUpdates && email && !seen[email]) {
      seen[email] = true;
      emails.push(email);
    }
  }
  return emails;
}

/** Kør denne manuelt i editoren for at teste uden at sende mails. */
function testListSubscribers() {
  Logger.log(getSubscriberEmails_());
}
