/**
 * ALARME EMAIL — Le Yéti de Villard
 *
 * Cette fonction envoie un email à thibdebec@gmail.com
 * dès qu'une nouvelle demande arrive dans le Google Sheet.
 *
 * --> À COLLER en plus de ton code doPost (ne supprime rien d'autre).
 * --> Pas besoin de "Déployer". On l'active avec un "déclencheur" (voir étapes).
 */
function notifyNewRow(e) {
  // On n'agit que lorsqu'une nouvelle ligne est ajoutée
  if (e && e.changeType && e.changeType !== 'INSERT_ROW') return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Demandes');
  const row = sheet.getLastRow();
  if (row < 1) return;

  // Colonnes du tableau : Date | Nom | Email | Arrivée | Départ | Message
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
