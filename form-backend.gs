/**
 * Le Yéti de Villard — Backend du formulaire de contact (Google Apps Script)
 *
 * Ce script reçoit les soumissions du formulaire du site, les enregistre
 * dans le Google Sheet ET envoie une notification par email.
 *
 * --- Installation ---
 * 1. Ouvre le Google Sheet :
 *    https://docs.google.com/spreadsheets/d/1bb5rUoXyH-AzTo-Tj5B-Er2FdCOvjbaKYO-OLdgd6uQ/edit
 * 2. Menu "Extensions" > "Apps Script".
 * 3. Remplace TOUT le code existant par celui-ci, puis enregistre (Ctrl/Cmd + S).
 * 4. Clique sur "Déployer" > "Gérer les déploiements" > (icône crayon) >
 *    "Version" = "Nouvelle version" > "Déployer".
 *    -> L'URL /exec reste la même, donc rien à changer sur le site.
 * 5. À la première exécution, Google demande l'autorisation d'envoyer des
 *    emails : accepte (compte du propriétaire du script).
 */

const SHEET_ID = '1bb5rUoXyH-AzTo-Tj5B-Er2FdCOvjbaKYO-OLdgd6uQ';
const NOTIFY_EMAIL = 'thibdebec@gmail.com';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 1) Enregistrer la demande dans le Google Sheet (première feuille)
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    sheet.appendRow([
      new Date(),
      data.prenom || '',
      data.nom || '',
      data.email || '',
      data.telephone || '',
      data.arrivee || '',
      data.depart || '',
      data.message || ''
    ]);

    // 2) Envoyer une notification email à chaque soumission
    const nomComplet = ((data.prenom || '') + ' ' + (data.nom || '')).trim();
    const subject = 'Nouvelle demande - Le Yéti de Villard' +
      (nomComplet ? ' (' + nomComplet + ')' : '');

    const body =
      'Nouvelle demande de réservation reçue via le site :\n\n' +
      'Prénom        : ' + (data.prenom || '') + '\n' +
      'Nom           : ' + (data.nom || '') + '\n' +
      'Email         : ' + (data.email || '') + '\n' +
      'Téléphone     : ' + (data.telephone || '') + '\n' +
      'Date arrivée  : ' + (data.arrivee || '') + '\n' +
      'Date départ   : ' + (data.depart || '') + '\n\n' +
      'Message :\n' + (data.message || '') + '\n';

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: subject,
      body: body,
      replyTo: data.email || NOTIFY_EMAIL  // répondre directement au prospect
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
