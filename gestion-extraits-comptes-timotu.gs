// Configuration initiale
const DEST_SHEET_ID = "1phKrlZI6WPfU3UKMqBOUAYfm7AGq9NU9Mg30zNvUzvI"; // Google Sheet destination
const CSV_FILE_NAME = "CA20250404_075756.csv";     // Fichier CSV pour test
const CSV_FOLDER_ID = "16FYmwWvKy2icIUgZWQ_jfOMnPycZWV2M"; // ID du dossier Google Drive contenant les CSV

function processCsvAndConsolidate() {
  var ss = SpreadsheetApp.openById(DEST_SHEET_ID);
  
  // Récupérer ou créer l'onglet temporaire "Sheet0"
  var sheet0 = ss.getSheetByName("Sheet0");
  if (!sheet0) {
    sheet0 = ss.insertSheet("Sheet0");
  } else {
    sheet0.clearContents();
  }
  
  var csvFile = getCsvFileByName(CSV_FILE_NAME);
  if (!csvFile) {
    Logger.log("Fichier CSV introuvable : " + CSV_FILE_NAME);
    return;
  }
  
  // Lecture du CSV avec encodage ISO-8859-1
  var csvContent = csvFile.getBlob().getDataAsString("ISO-8859-1");
  
  // Utiliser le parseur personnalisé pour reconstituer les lignes multi-lignes
  var allLines = parseCsvCustom(csvContent, ";");
  if (allLines.length === 0) {
    Logger.log("Aucune donnée parsée.");
    return;
  }
  
  // Rechercher la ligne d'en-tête attendue (premier champ "Date" et deuxième champ contenant "libellé")
  var headerIndex = -1;
  for (var i = 0; i < allLines.length; i++) {
    var row = allLines[i];
    if (row.length >= 4 && row[0].trim().toLowerCase() === "date" && row[1].toLowerCase().indexOf("libell") !== -1) {
      headerIndex = i;
      break;
    }
  }
  if (headerIndex === -1) {
    Logger.log("L'en-tête attendu n'a pas été trouvé.");
    return;
  }
  
  // Conserver les lignes à partir de l'en-tête
  var csvData = allLines.slice(headerIndex);
  if (csvData[0].length < 4) {
    Logger.log("L'en-tête ne comporte pas 4 colonnes.");
    return;
  }
  
  // Placer les données dans l'onglet temporaire "Sheet0"
  sheet0.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  
  // Ne conserver que les lignes de crédit : supprimer les lignes dont la colonne C (Débit euros) est renseignée
  supprimerLignesDebits(sheet0);
  
  // Étape 1 : Nettoyer la ligne vérifiée existante dans l'onglet destination
  clearVerifiedLine(ss);
  
  // Étape 2 : Marquer la ligne 11 en plaçant la valeur 1 dans la colonne Check (E) sans colorer
  markCheckRow11(ss);
  
  // Étape 3 : Copier le bloc de lignes du CSV situées avant la correspondance
  copierBlocDeBase(ss, sheet0);
  
  // Étape 4 : Rechercher la ligne qui contient 1 dans la colonne Check et appliquer la mise en forme (fond vert sur les 4 premières cellules)
  applyVerifiedFormatting(ss);
  
  // Supprimer l'onglet temporaire
  ss.deleteSheet(sheet0);
  
  Logger.log("Traitement terminé.");
}

function getCsvFileByName(fileName) {
  var folder = DriveApp.getFolderById(CSV_FOLDER_ID);
  var files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    return files.next();
  }
  return null;
}

/**
 * Parseur CSV personnalisé :
 * Scinde le texte en lignes brutes, puis reconstitue les lignes
 * en tenant compte du nombre de guillemets pour gérer les enregistrements multi-lignes.
 */
function parseCsvCustom(text, delimiter) {
  var rawLines = text.split(/\r?\n/);
  var fixedLines = [];
  var buffer = "";
  var inMultiline = false;
  
  for (var i = 0; i < rawLines.length; i++) {
    var line = rawLines[i];
    if (inMultiline) {
      buffer += "\n" + line;
    } else {
      buffer = line;
    }
    var quoteCount = (buffer.match(/"/g) || []).length;
    if (quoteCount % 2 === 0) {
      fixedLines.push(buffer);
      inMultiline = false;
      buffer = "";
    } else {
      inMultiline = true;
    }
  }
  
  var result = [];
  for (var j = 0; j < fixedLines.length; j++) {
    var parsed = Utilities.parseCsv(fixedLines[j], delimiter);
    if (parsed && parsed.length > 0) {
      result.push(parsed[0]);
    }
  }
  return result;
}

// Supprime à partir de la ligne 11 de Sheet0 les lignes dont la colonne C (Débit euros) n'est pas vide
function supprimerLignesDebits(sheet0) {
  var lastRow = sheet0.getLastRow();
  for (var i = lastRow; i >= 11; i--) {
    var cellValue = sheet0.getRange(i, 3).getValue();
    if (cellValue !== "" && cellValue !== null) {
      sheet0.deleteRow(i);
    }
  }
}

/**
 * Supprime toute mise en forme vérifiée (fond vert et valeur 1 dans la colonne Check) de l'onglet "justecredits-sansdeb".
 */
function clearVerifiedLine(ss) {
  var destSheet = ss.getSheetByName("justecredits-sansdeb");
  if (!destSheet) return;
  var lastRow = destSheet.getLastRow();
  // Effacer le fond de toutes les cellules et vider la colonne E (Check)
  destSheet.getRange(1, 1, lastRow, destSheet.getLastColumn()).setBackground(null);
  destSheet.getRange(1, 5, lastRow, 1).clearContent();
}

/**
 * Marque la ligne 11 en plaçant la valeur 1 dans la colonne Check (colonne E) de l'onglet "justecredits-sansdeb".
 * Cette fonction ne modifie pas la couleur.
 */
function markCheckRow11(ss) {
  var destSheet = ss.getSheetByName("justecredits-sansdeb");
  if (!destSheet) {
    Logger.log("L'onglet 'justecredits-sansdeb' est introuvable.");
    return;
  }
  destSheet.getRange(11, 5).setValue(1);
}

/**
 * Rechercher dans l'onglet "justecredits-sansdeb" la ligne contenant la valeur 1 en colonne Check (E)
 * et appliquer le fond vert (#00B050) aux 4 premières cellules de cette ligne.
 */
function applyVerifiedFormatting(ss) {
  var destSheet = ss.getSheetByName("justecredits-sansdeb");
  if (!destSheet) return;
  var lastRow = destSheet.getLastRow();
  for (var i = 1; i <= lastRow; i++) {
    var cell = destSheet.getRange(i, 5);
    if (cell.getValue() == 1) {
      destSheet.getRange(i, 1, 1, 4).setBackground("#00B050");
      break;
    }
  }
}

/**
 * Copie le bloc de base du CSV (stocké dans Sheet0) vers l'onglet "justecredits-sansdeb".
 * La procédure consiste à rechercher, à partir de la ligne 11 du CSV, la première ligne dont le libellé (colonne B),
 * après normalisation (suppression des espaces, des guillemets et conversion en minuscules), correspond à celui de la ligne vérifiée
 * (la ligne contenant 1 dans la colonne Check) dans l'onglet destination.
 * Ensuite, on copie toutes les lignes du CSV situées avant cette correspondance en insérant ces lignes au-dessus de la ligne vérifiée.
 */
function copierBlocDeBase(ss, sheet0) {
  var destSheet = ss.getSheetByName("justecredits-sansdeb");
  if (!destSheet) {
    Logger.log("L'onglet 'justecredits-sansdeb' est introuvable.");
    return;
  }
  
  var dataSource = sheet0.getDataRange().getValues();
  // Récupérer la ligne vérifiée dans l'onglet destination (celle contenant 1 en colonne E)
  var destData = destSheet.getDataRange().getValues();
  var verifiedRow = -1;
  for (var i = 1; i <= destData.length; i++) {
    if (String(destSheet.getRange(i, 5).getValue()).trim() === "1") {
      verifiedRow = i;
      break;
    }
  }
  if (verifiedRow === -1) {
    Logger.log("Aucune ligne vérifiée trouvée dans la destination.");
    return;
  }
  
  var destRowData = destSheet.getRange(verifiedRow, 1, 1, 4).getValues()[0];
  
  // Chercher dans le CSV (à partir de la ligne 11) la première ligne dont le libellé correspond à celui de la ligne vérifiée
  var matchIndex = -1;
  for (var i = 11; i <= dataSource.length; i++) {
    var rowData = dataSource[i - 1].slice(0, 4);
    if (normalizeString(rowData[1]) === normalizeString(destRowData[1])) {
      matchIndex = i;
      break;
    }
  }
  
  if (matchIndex === -1) {
    Logger.log("Aucune correspondance trouvée dans le CSV pour la ligne vérifiée.");
    return;
  }
  
  // Copier les lignes du CSV situées avant la ligne correspondante (de la ligne 11 jusqu'à matchIndex - 1)
  var rowsToCopy = matchIndex - 11;
  if (rowsToCopy > 0) {
    destSheet.insertRows(11, rowsToCopy);
    var blockData = sheet0.getRange(11, 1, rowsToCopy, 4).getValues();
    destSheet.getRange(11, 1, rowsToCopy, 4).setValues(blockData);
  }
  Logger.log("Bloc de base copié.");
}

/**
 * Normalise une chaîne en convertissant en minuscules, en supprimant les guillemets et tous les espaces.
 */
function normalizeString(str) {
  return String(str).toLowerCase().replace(/"/g, "").replace(/\s+/g, "");
}
