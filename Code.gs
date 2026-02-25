function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📞 Dupliquer la ligne')
    .addItem('Dupliquer', 'nouvelAppel')
    .addToUi();
  ui.createMenu('🎙️ Générer CR')
    .addItem('Générer', 'openCompteRendu')
    .addToUi();
}

function nouvelAppel() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'TRAVAIL') {
    SpreadsheetApp.getUi().alert("Fonctionne uniquement sur l'onglet TRAVAIL.");
    return;
  }
  var row = SpreadsheetApp.getActiveRange().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("Sélectionnez une ligne de données.");
    return;
  }
  var lastCol = sheet.getLastColumn();
  sheet.insertRowAfter(row);
  sheet.getRange(row, 1, 1, lastCol).copyTo(sheet.getRange(row + 1, 1, 1, lastCol));
  sheet.getRange(row + 1, 1, 1, lastCol).setBackground('#FFF2CC');
  SpreadsheetApp.flush();
}

function openCompteRendu() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'TRAVAIL') {
    SpreadsheetApp.getUi().alert("Fonctionne uniquement sur l'onglet TRAVAIL.");
    return;
  }
  var row = SpreadsheetApp.getActiveRange().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("Sélectionnez une ligne de données.");
    return;
  }

  // Récupérer les headers + valeurs de la ligne
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  // Construire l'objet de données
  var rowData = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) rowData[headers[i]] = values[i];
  }
  rowData['_rowNumber'] = row;

  // URL de la page externe hébergée + webhook encodé en base64
  var EXTERNAL_PAGE = 'https://sofianekorbi.github.io/business-on/audio-recorder.html';
  var WEBHOOK_URL = 'https://n8n.srv1353111.hstgr.cloud/webhook-test/26a15343-e911-4400-a918-b3cf06074f15';
  var webhookB64 = Utilities.base64Encode(WEBHOOK_URL);
  var rowDataB64 = Utilities.base64Encode(JSON.stringify(rowData));

  var nom = ((rowData['PRENOM'] || '') + ' ' + (rowData['Nom'] || '')).toString().trim();
  var email = (rowData['E-mail'] || '').toString().trim();

  var url = EXTERNAL_PAGE
    + '?webhook=' + encodeURIComponent(webhookB64)
    + '&rowData=' + encodeURIComponent(rowDataB64)
    + '&row=' + row
    + '&name=' + encodeURIComponent(nom)
    + '&email=' + encodeURIComponent(email);

  // Ouvrir automatiquement dans un nouvel onglet + fallback lien cliquable
  var html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:Google Sans,Arial,sans-serif;padding:20px;text-align:center;">' +
    '<p id="msg" style="margin-bottom:16px;color:#374151;font-size:14px;">Ouverture en cours...</p>' +
    '<a id="link" href="' + url.replace(/"/g, '&quot;') + '" target="_blank" ' +
    'style="display:none;padding:12px 32px;background:#111827;color:white;text-decoration:none;border-radius:8px;font-weight:600;font-size:14px;" ' +
    'onclick="google.script.host.close()">' +
    '🎙️ Ouvrir l\'enregistreur</a>' +
    '<p style="margin-top:12px;color:#9ca3af;font-size:11px;">Ligne ' + row + ' — ' + nom + '</p>' +
    '<script>' +
    'var w = window.open("' + url.replace(/"/g, '\\"') + '", "_blank");' +
    'if (w) { google.script.host.close(); }' +
    'else { document.getElementById("msg").textContent = "Le popup a été bloqué. Cliquez ci-dessous :"; document.getElementById("link").style.display = "inline-block"; }' +
    '</script>' +
    '</body></html>'
  ).setWidth(400).setHeight(160);

  SpreadsheetApp.getUi().showModalDialog(html, 'Générer le compte rendu');
}

function sendToWebhook(payload) {
  var response = UrlFetchApp.fetch(
    'https://n8n.srv1353111.hstgr.cloud/webhook-test/26a15343-e911-4400-a918-b3cf06074f15',
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );
  var code = response.getResponseCode();
  if (code >= 200 && code < 300) {
    return 'OK';
  } else {
    throw new Error('Erreur ' + code + ': ' + response.getContentText());
  }
}
