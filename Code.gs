function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Business On')
    .addItem('Ouvrir le panneau', 'openSidebar')
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Business On');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Retourne les infos de la ligne sélectionnée
function getSelectedRowData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'TRAVAIL') {
    return { error: 'wrong_sheet' };
  }
  var row = SpreadsheetApp.getActiveRange().getRow();
  if (row <= 1) {
    return { error: 'header_row' };
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  var rowData = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) rowData[headers[i]] = values[i];
  }
  rowData['_rowNumber'] = row;

  var nom = ((rowData['PRENOM'] || '') + ' ' + (rowData['Nom'] || '')).toString().trim();
  var email = (rowData['E-mail'] || '').toString().trim();
  var societe = (rowData['Société'] || rowData['Societe'] || '').toString().trim();
  var statut = (rowData['STATUTS'] || '').toString().trim();

  return {
    row: row,
    nom: nom,
    email: email,
    societe: societe,
    statut: statut,
    rowData: rowData
  };
}

// Duplique la ligne et met la date du jour
function dupliquerLigne() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'TRAVAIL') {
    return { error: "Fonctionne uniquement sur l'onglet TRAVAIL." };
  }
  var row = SpreadsheetApp.getActiveRange().getRow();
  if (row <= 1) {
    return { error: "Sélectionnez une ligne de données." };
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Dupliquer
  sheet.insertRowAfter(row);
  sheet.getRange(row, 1, 1, lastCol).copyTo(sheet.getRange(row + 1, 1, 1, lastCol));
  sheet.getRange(row + 1, 1, 1, lastCol).setBackground('#FFF2CC');

  // Trouver la colonne "Date d'appel" et mettre la date du jour
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === "Date d'appel") {
      sheet.getRange(row + 1, i + 1).setValue(new Date());
      sheet.getRange(row + 1, i + 1).setNumberFormat('dd/MM/yyyy');
      break;
    }
  }

  // Vider le commentaire de la nouvelle ligne
  for (var j = 0; j < headers.length; j++) {
    if (headers[j] === 'Commentaire') {
      sheet.getRange(row + 1, j + 1).setValue('');
      break;
    }
  }

  SpreadsheetApp.flush();

  // Sélectionner la nouvelle ligne
  sheet.getRange(row + 1, 1).activate();

  return { success: true, newRow: row + 1 };
}

// Génère l'URL de la page d'enregistrement
function getRecordingUrl() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'TRAVAIL') {
    return { error: "Fonctionne uniquement sur l'onglet TRAVAIL." };
  }
  var row = SpreadsheetApp.getActiveRange().getRow();
  if (row <= 1) {
    return { error: "Sélectionnez une ligne de données." };
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  var rowData = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) rowData[headers[i]] = values[i];
  }
  rowData['_rowNumber'] = row;

  var EXTERNAL_PAGE = 'https://sofianekorbi.github.io/business-on/audio-recorder.html';
  var rowDataB64 = Utilities.base64Encode(JSON.stringify(rowData));

  var nom = ((rowData['PRENOM'] || '') + ' ' + (rowData['Nom'] || '')).toString().trim();
  var email = (rowData['E-mail'] || '').toString().trim();

  var url = EXTERNAL_PAGE
    + '?rowData=' + encodeURIComponent(rowDataB64)
    + '&row=' + row
    + '&name=' + encodeURIComponent(nom)
    + '&email=' + encodeURIComponent(email);

  return { url: url, row: row, nom: nom };
}

// Proxy doPost : reçoit les données de la page externe et forward à n8n
function doPost(e) {
  var WEBHOOK_URL = 'https://business-on.bkbx.io/webhook/26a15343-e911-4400-a918-b3cf06074f15';

  var response = UrlFetchApp.fetch(WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: e.postData.contents,
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var output = ContentService.createTextOutput(
    JSON.stringify({ status: code >= 200 && code < 300 ? 'ok' : 'error', code: code })
  ).setMimeType(ContentService.MimeType.JSON);

  return output;
}

function sendToWebhook(payload) {
  var response = UrlFetchApp.fetch(
    'https://business-on.bkbx.io/webhook/26a15343-e911-4400-a918-b3cf06074f15',
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
