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

    // Notification email envoyée à CHAQUE demande
    MailApp.sendEmail({
      to: 'thibdebec@gmail.com',
      subject: '🏔️ Nouvelle demande - Le Yéti de Villard',
      replyTo: data.email || 'thibdebec@gmail.com',
      body: 'Nouvelle demande reçue via le site :\n\n'
        + 'Prénom : ' + (data.prenom || '') + '\n'
        + 'Nom : ' + (data.nom || '') + '\n'
        + 'Email : ' + (data.email || '') + '\n'
        + 'Téléphone : ' + (data.telephone || '') + '\n'
        + 'Arrivée : ' + (data.arrivee || '') + '\n'
        + 'Départ : ' + (data.depart || '') + '\n\n'
        + 'Message :\n' + (data.message || '')
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
