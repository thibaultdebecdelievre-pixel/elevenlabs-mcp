/**
 * ===========================================================
 *  Le Yéti de Villard — SCRIPT COMPLET (formulaire + email)
 * ===========================================================
 *
 *  COMMENT L'UTILISER :
 *  1) Ouvre ton Google Sheet "Demandes" > menu Extensions > Apps Script.
 *  2) Sélectionne TOUT le code existant et supprime-le.
 *  3) Colle CE code en entier, puis clique sur 💾 (Enregistrer).
 *
 *  POUR ACTIVER L'EMAIL (pas besoin du bouton "Déployer") :
 *  4) Dans la barre de gauche, clique sur ⏰ (Déclencheurs).
 *  5) En bas à droite : "+ Ajouter un déclencheur".
 *  6) Règle :
 *        - Fonction à exécuter ........ notifyNewRow
 *        - Source de l'événement ...... Depuis une feuille de calcul
 *        - Type d'événement ........... Lors d'une modification
 *  7) Enregistrer, puis Autoriser (envoi d'emails).
 *
 *  -> Le site n'a PAS besoin d'être modifié, et le bouton "Déployer"
 *     n'est PAS nécessaire.
 * ===========================================================
 */


// 1) Reçoit les réponses du formulaire et les enregistre dans le tableau.
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Demandes');
    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),
      data.nom || '',
      data.email || '',
      data.arrivee || '',
      data.depart || '',
      data.message || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// 2) Envoie un email à thibdebec@gmail.com à chaque nouvelle demande.
//    (déclenché automatiquement par le Sheet — voir étapes 4 à 7 ci-dessus)
function notifyNewRow(e) {
  if (e && e.changeType && e.changeType !== 'INSERT_ROW') return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Demandes');
  const row = sheet.getLastRow();
  if (row < 1) return;

  // Colonnes : Date | Nom | Email | Arrivée | Départ | Message
  const v = sheet.getRange(row, 1, 1, 6).getValues()[0];

  MailApp.sendEmail({
    to: 'thibdebec@gmail.com',
    subject: '🏔️ Nouvelle demande - Le Yéti de Villard',
    body: 'Nouvelle demande reçue via le site :\n\n'
      + 'Nom      : ' + (v[1] || '') + '\n'
      + 'Email    : ' + (v[2] || '') + '\n'
      + 'Arrivée  : ' + (v[3] || '') + '\n'
      + 'Départ   : ' + (v[4] || '') + '\n\n'
      + 'Message :\n' + (v[5] || '')
  });
}
