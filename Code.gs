function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Business On')
    .addItem('Ouvrir le panneau', 'openSidebar')
    .addToUi();
}

function openSidebar() {
  // Stocker la sélection actuelle avant d'ouvrir
  storeCurrentSelection_();
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Business On');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Trigger automatique : se déclenche à chaque changement de sélection
// Optimisé : ne lit que les colonnes nécessaires pour la sidebar (rapide)
function onSelectionChange(e) {
  try {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== 'TRAVAIL') {
      PropertiesService.getScriptProperties().setProperty('sidebarData',
        JSON.stringify({ error: 'wrong_sheet' }));
      return;
    }
    var row = e.range.getRow();
    if (row <= 1) {
      PropertiesService.getScriptProperties().setProperty('sidebarData',
        JSON.stringify({ error: 'header_row' }));
      return;
    }

    var data = buildRowDataLight_(sheet, row);
    PropertiesService.getScriptProperties().setProperty('sidebarData', JSON.stringify(data));
  } catch (err) {
    // Silencieux pour le trigger
  }
}

// Version légère : lit seulement les colonnes nécessaires pour la sidebar
function buildRowDataLight_(sheet, row) {
  var headers = getCachedHeaders_(sheet);
  var lastCol = headers.length;
  var values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  var nom = '';
  var email = '';
  var societe = '';
  var statut = '';

  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h === 'PRENOM') nom = (values[i] || '').toString();
    else if (h === 'Nom') nom = (nom ? nom + ' ' : '') + (values[i] || '').toString();
    else if (h === 'E-mail') email = (values[i] || '').toString().trim();
    else if (h === 'Société' || h === 'Societe') societe = (values[i] || '').toString().trim();
    else if (h === 'STATUTS') statut = (values[i] || '').toString().trim();
  }

  return { row: row, nom: nom.trim(), email: email, societe: societe, statut: statut };
}

// Cache les headers pour éviter de relire la ligne 1 à chaque sélection
var headersCache_ = null;
function getCachedHeaders_(sheet) {
  if (!headersCache_) {
    var lastCol = sheet.getLastColumn();
    headersCache_ = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  return headersCache_;
}

// Fonction interne : construit les données d'une ligne
function buildRowData_(sheet, row) {
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

// Stocke la sélection actuelle dans PropertiesService
function storeCurrentSelection_() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() !== 'TRAVAIL') {
      PropertiesService.getScriptProperties().setProperty('sidebarData',
        JSON.stringify({ error: 'wrong_sheet' }));
      return;
    }
    var row = SpreadsheetApp.getActiveRange().getRow();
    if (row <= 1) {
      PropertiesService.getScriptProperties().setProperty('sidebarData',
        JSON.stringify({ error: 'header_row' }));
      return;
    }
    var data = buildRowData_(sheet, row);
    PropertiesService.getScriptProperties().setProperty('sidebarData', JSON.stringify(data));
  } catch (err) {}
}

// Appelé par la sidebar : lit les données stockées (rapide, pas de getActiveRange)
function getStoredRowData() {
  var stored = PropertiesService.getScriptProperties().getProperty('sidebarData');
  if (!stored) return null;
  try {
    return JSON.parse(stored);
  } catch (e) {
    return null;
  }
}

// Appelé par le bouton Rafraîchir : force la lecture + stockage
function getSelectedRowData() {
  storeCurrentSelection_();
  return getStoredRowData();
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
