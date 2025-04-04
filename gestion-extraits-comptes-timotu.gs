// Configuration générale (CONTROL_COLUMN supprimée)
const CONFIG = {
  DEST_SHEET_ID: "1phKrlZI6WPfU3UKMqBOUAYfm7AGq9NU9Mg30zNvUzvI",
  CSV_FILE_NAME: "CA20250404_075756.csv",
  CSV_FOLDER_ID: "16FYmwWvKy2icIUgZWQ_jfOMnPycZWV2M",
  DEST_SHEET_NAME: "conso-23111741301",
  START_ROW: 2 // Première ligne de données (la ligne 1 reste l'en-tête)
};

function processCsvAndConsolidate() {
  const ss = SpreadsheetApp.openById(CONFIG.DEST_SHEET_ID);
  const sheet0 = createOrClearTempSheet(ss);
  
  const csvFile = getCsvFileByName(CONFIG.CSV_FILE_NAME);
  if (!csvFile) {
    Logger.log("Fichier CSV introuvable : " + CONFIG.CSV_FILE_NAME);
    return;
  }
  
  const csvContent = csvFile.getBlob().getDataAsString("ISO-8859-1");
  const allLines = parseCsvCustom(csvContent, ";");
  if (allLines.length === 0) {
    Logger.log("Aucune donnée parsée.");
    return;
  }
  
  const headerIndex = findCsvHeader(allLines);
  if (headerIndex === -1) {
    Logger.log("En-tête CSV introuvable.");
    return;
  }
  
  // On saute l'en-tête pour ne prendre que les vraies données
  const csvData = allLines.slice(headerIndex + 1);
  
  // Vérification solide avant de continuer
  if (csvData.length === 0 || !csvData[0] || csvData[0].length < 4) {
    Logger.log("Erreur : pas de données trouvées après l'en-tête !");
    return;
  }
  if (!isProbablyDate(csvData[0][0])) {
    Logger.log("Erreur : première cellule des données n'est pas une date valide : " + csvData[0][0]);
    return;
  }
  
  // Copie le CSV dans la feuille temporaire
  sheet0.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  
  supprimerLignesDebits(sheet0);
  copierBlocDeBase(sheet0);
  
  ss.deleteSheet(sheet0);
  Logger.log("Traitement terminé.");
}

function createOrClearTempSheet(ss) {
  let sheet0 = ss.getSheetByName("Sheet0");
  if (!sheet0) {
    sheet0 = ss.insertSheet("Sheet0");
  } else {
    sheet0.clearContents();
  }
  return sheet0;
}

function getDestSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.DEST_SHEET_ID);
  return ss.getSheetByName(CONFIG.DEST_SHEET_NAME);
}

function getCsvFileByName(fileName) {
  const folder = DriveApp.getFolderById(CONFIG.CSV_FOLDER_ID);
  const files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

/**
 * Parse le contenu CSV en gérant les lignes multi-lignes et en ignorant les lignes vides.
 */
function parseCsvCustom(text, delimiter) {
  const rawLines = text.split(/\r?\n/);
  const fixedLines = [];
  let buffer = "";
  let inMultiline = false;
  
  for (const line of rawLines) {
    if (line.trim() === "") continue; // ignorer les lignes vides
    buffer = inMultiline ? buffer + "\n" + line : line;
    const quoteCount = (buffer.match(/"/g) || []).length;
    if (quoteCount % 2 === 0) {
      fixedLines.push(buffer);
      inMultiline = false;
      buffer = "";
    } else {
      inMultiline = true;
    }
  }
  
  const parsedLines = fixedLines
    .map(l => Utilities.parseCsv(l, delimiter)[0])
    .filter(row => row && row.length > 0);
    
  return parsedLines;
}

function findCsvHeader(lines) {
  for (let i = 0; i < lines.length; i++) {
    const row = lines[i];
    if (row.length >= 4 && row[0].trim().toLowerCase() === "date" && row[1].toLowerCase().includes("libell")) {
      return i;
    }
  }
  return -1;
}

/**
 * Vérifie si une chaîne ressemble à une date au format DD/MM/YYYY.
 */
function isProbablyDate(str) {
  if (!str) return false;
  str = String(str).trim();
  return /^\d{2}\/\d{2}\/\d{4}$/.test(str);
}

/**
 * Supprime, dans la feuille temporaire, les lignes à partir de START_ROW dont la colonne 3 (Débit euros)
 * est renseignée.
 */
function supprimerLignesDebits(sheet0) {
  for (let i = sheet0.getLastRow(); i >= CONFIG.START_ROW; i--) {
    const cellValue = sheet0.getRange(i, 3).getValue();
    if (cellValue !== "" && cellValue !== null) {
      sheet0.deleteRow(i);
    }
  }
}

/**
 * Copie le bloc de données du CSV (stocké dans la feuille temporaire) vers la feuille destination.
 * Ici, on utilise la première ligne de données de la feuille destination (définie par CONFIG.START_ROW)
 * comme référence pour le libellé, puis on cherche dans le CSV la ligne dont le libellé correspond.
 * On copie ensuite les lignes du CSV situées avant cette correspondance dans la destination,
 * en les insérant dans les colonnes 2 à 5 (en conservant l'en-tête en ligne 1).
 */
function copierBlocDeBase(sheet0) {
  const destSheet = getDestSheet();
  if (!destSheet) {
    Logger.log("Onglet destination introuvable.");
    return;
  }
  
  const dataSource = sheet0.getDataRange().getValues();
  
  // Utiliser la première ligne de données de la destination comme référence
  const verifiedRow = CONFIG.START_ROW;
  const destRowData = destSheet.getRange(verifiedRow, 2, 1, 4).getValues()[0];
  
  let matchIndex = -1;
  for (let i = CONFIG.START_ROW; i <= dataSource.length; i++) {
    const rowData = dataSource[i - 1].slice(0, 4); // [date, libellé, débit, crédit] du CSV
    if (normalizeString(rowData[1]) === normalizeString(destRowData[1])) {
      matchIndex = i;
      break;
    }
  }
  
  if (matchIndex === -1) {
    Logger.log("Aucune correspondance trouvée pour déterminer le bloc à copier.");
    return;
  }
  
  const rowsToCopy = matchIndex - CONFIG.START_ROW;
  if (rowsToCopy > 0) {
    destSheet.insertRows(CONFIG.START_ROW, rowsToCopy);
    const blockData = sheet0.getRange(CONFIG.START_ROW, 1, rowsToCopy, 4).getValues();
    destSheet.getRange(CONFIG.START_ROW, 2, rowsToCopy, 4).setValues(blockData);
  }
  Logger.log("Bloc de base copié.");
}

/**
 * Normalise une chaîne en convertissant en minuscules, en supprimant les guillemets et les espaces.
 */
function normalizeString(str) {
  return String(str).toLowerCase().replace(/"/g, "").replace(/\s+/g, "");
}


// Nouvelle configuration pour le master
const CONFIG_MASTER = {
  DEST_SHEET_ID: "1phKrlZI6WPfU3UKMqBOUAYfm7AGq9NU9Mg30zNvUzvI",
  CSV_FOLDER_ID: "16FYmwWvKy2icIUgZWQ_jfOMnPycZWV2M",
  MASTER_SHEET_NAME: "master"  // Nom de l'onglet master dans le Google Sheet
};

/**
 * Lancement de la mise à jour pour tous les fichiers CSV du dossier.
 * Pour chaque fichier dont le nom correspond au pattern CAYYYYMMDD_random.csv,
 * le script extrait le numéro de compte (suite de 11 chiffres après "Compte courant n°")
 * et met à jour l'onglet master en inscrivant la date et l'heure du passage du script
 * en colonne C sur la même ligne où se trouve le numéro de compte dans la plage B6:B11.
 */
function lancementMajAllCsv() {
  const folder = DriveApp.getFolderById(CONFIG_MASTER.CSV_FOLDER_ID);
  const files = folder.getFiles();
  
  // Pattern : CA suivi de 8 chiffres, underscore, puis au moins un chiffre, et l'extension .csv
  const filePattern = /^CA\d{8}_\d+\.csv$/i;
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    if (!filePattern.test(fileName)) continue;
    
    // Récupération du contenu du fichier en ISO-8859-1
    const fileContent = file.getBlob().getDataAsString("ISO-8859-1");
    
    // Recherche du numéro de compte (11 chiffres) après "Compte courant n°"
    const accountRegex = /Compte courant n°\s*(\d{11})/;
    const match = fileContent.match(accountRegex);
    
    if (match && match[1]) {
      const accountNumber = match[1];
      updateMasterSheet(accountNumber);
    } else {
      Logger.log("Numéro de compte non trouvé dans le fichier : " + fileName);
    }
  }
}

/**
 * Recherche dans l'onglet master (plage B6:B11) le numéro de compte fourni.
 * Si trouvé, inscrit la date et l'heure du passage du script dans la colonne C de la même ligne.
 */
function updateMasterSheet(accountNumber) {
  const ss = SpreadsheetApp.openById(CONFIG_MASTER.DEST_SHEET_ID);
  const masterSheet = ss.getSheetByName(CONFIG_MASTER.MASTER_SHEET_NAME);
  if (!masterSheet) {
    Logger.log("Onglet '" + CONFIG_MASTER.MASTER_SHEET_NAME + "' introuvable.");
    return;
  }
  
  const range = masterSheet.getRange("B6:B11");
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === accountNumber) {
      // La ligne dans le sheet est i+6 (car on part de la ligne 6)
      masterSheet.getRange(i + 6, 3).setValue(new Date());
      Logger.log("Compte " + accountNumber + " mis à jour à " + new Date());
      break;
    }
  }
}

