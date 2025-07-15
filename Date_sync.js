
// =================================================================================
// SECTION 1 : CONFIGURATION
// =================================================================================

// À MODIFIER : Mettez votre clé d'API privée HubSpot
const API_KEY = "HB_API_KEY"; // Remplacez par votre vraie clé

// ID de l'objet personnalisé "Device" à mettre à jour
const DEVICE_OBJECT_TYPE_ID = '2-37633992';


// =================================================================================
// SECTION 2 : FONCTION DU MENU
// =================================================================================

/**
 * Crée un menu personnalisé dans l'interface Google Sheets au démarrage.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Actions HubSpot')
    .addItem('Mettre à jour les dates des Devices', 'updateDeviceDatesInHubSpot')
    .addToUi();
}


// =================================================================================
// SECTION 3 : SCRIPT PRINCIPAL DE MISE À JOUR DES DATES
// =================================================================================

/**
 * Trouve l'index (position) d'une colonne en fonction de son nom d'en-tête.
 * @param {string[]} headers La liste des noms d'en-tête.
 * @param {string} columnName Le nom de la colonne à trouver.
 * @return {number} L'index de la colonne (commence à 0) ou -1 si non trouvée.
 */
function getColumnIndex(headers, columnName) {
  return headers.indexOf(columnName);
}

/**
 * Met à jour les propriétés de date d'un objet "Device" dans HubSpot,
 * en ignorant les lignes déjà traitées avec succès.
 */
function updateDeviceDatesInHubSpot() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const ui = SpreadsheetApp.getUi();

  // Si la feuille a moins de 2 lignes (juste l'en-tête ou vide), on arrête.
  if (lastRow < 2) {
    // On garde une alerte pour ce cas car aucune action n'est possible.
    ui.alert("Aucune donnée à traiter.");
    return;
  }

  // --- NOUVELLE CONFIGURATION DES COLONNES PAR NOM D'EN-TÊTE ---
  // Modifiez ces noms si les en-têtes de votre feuille sont différents.
  const DATE_INSTALL_HEADER = 'Date installatin';
  const FIN_LEASING_HEADER = 'Fin leasing';
  const DEVICE_ID_HEADER = 'DeviceID';   // Nom d'en-tête pour l'ID
  const SYNC_STATUS_HEADER = 'Sync';     // Nom d'en-tête pour le statut

  const SUCCESS_STATUS_TEXT = "Traité";

  // Lecture de la première ligne pour obtenir les en-têtes
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Recherche des index de chaque colonne requise
  const dateInstallColIndex = getColumnIndex(headers, DATE_INSTALL_HEADER);
  const finLeasingColIndex = getColumnIndex(headers, FIN_LEASING_HEADER);
  const deviceIdColIndex = getColumnIndex(headers, DEVICE_ID_HEADER);
  const syncStatusColIndex = getColumnIndex(headers, SYNC_STATUS_HEADER);

  // Vérification que toutes les colonnes nécessaires existent
  const requiredCols = [
    { name: DEVICE_ID_HEADER, index: deviceIdColIndex },
    { name: SYNC_STATUS_HEADER, index: syncStatusColIndex },
    { name: DATE_INSTALL_HEADER, index: dateInstallColIndex },
    { name: FIN_LEASING_HEADER, index: finLeasingColIndex }
  ];

  const missingCols = requiredCols.filter(col => col.index === -1);
  if (missingCols.length > 0) {
    const missingNames = missingCols.map(col => col.name).join(', ');
    ui.alert(`Action annulée. Colonne(s) manquante(s) : ${missingNames}. Veuillez vérifier les en-têtes de la feuille.`);
    return;
  }
  
  // Lecture de toutes les données et des statuts
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const statuses = sheet.getRange(2, syncStatusColIndex + 1, lastRow - 1, 1).getValues();
  
  let successfulUpdates = 0;
  let failedUpdates = 0;
  let skippedCount = 0;

  data.forEach((row, index) => {
    const rowNumber = index + 2;
    const currentStatus = statuses[index][0];

    // Si la ligne est déjà traitée, on l'ignore
    if (currentStatus === SUCCESS_STATUS_TEXT) {
      skippedCount++;
      return;
    }
    
    // On récupère les données via leur index
    const deviceId = row[deviceIdColIndex];

    if (!deviceId) {
      console.log(`Ligne ${rowNumber}: DeviceID manquant dans la colonne "${DEVICE_ID_HEADER}". Ligne ignorée.`);
      statuses[index][0] = "Erreur: DeviceID manquant";
      return;
    }

    const propertiesToUpdate = {};
    const installDate = row[dateInstallColIndex];
    const leasingEndDate = row[finLeasingColIndex];

    if (installDate instanceof Date) {
      propertiesToUpdate.date_de_livraison_effective__d_ = formatDateForHubSpot(installDate);
    }

    if (leasingEndDate instanceof Date) {
      propertiesToUpdate.date_de_fin_de_contrat_temporaire__d_ = formatDateForHubSpot(leasingEndDate);
    }

    if (Object.keys(propertiesToUpdate).length === 0) {
      console.log(`Ligne ${rowNumber}: Aucune date valide trouvée. Ligne ignorée.`);
      statuses[index][0] = "Aucune date à traiter"; // Ajout d'un statut plus clair
      return;
    }

    try {
      const url = `https://api.hubapi.com/crm/v3/objects/${DEVICE_OBJECT_TYPE_ID}/${deviceId}`;
      const payload = { properties: propertiesToUpdate };
      const options = {
        'method': 'patch',
        'contentType': 'application/json',
        'headers': { 'Authorization': `Bearer ${API_KEY}` },
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        console.log(`Ligne ${rowNumber}: Mise à jour réussie pour le DeviceID ${deviceId}.`);
        successfulUpdates++;
        statuses[index][0] = SUCCESS_STATUS_TEXT;
      } else {
        console.error(`Ligne ${rowNumber}: Échec de la mise à jour pour ${deviceId}. Code: ${responseCode}, Réponse: ${response.getContentText()}`);
        failedUpdates++;
        statuses[index][0] = `Erreur (${responseCode})`;
      }
    } catch (e) {
      console.error(`Ligne ${rowNumber}: Erreur API pour ${deviceId}. Erreur: ${e.toString()}`);
      failedUpdates++;
      statuses[index][0] = "Erreur (Exception)";
    }
  });
  
  // Écrit tous les statuts dans la feuille en une seule opération
  if (lastRow > 1) {
    sheet.getRange(2, syncStatusColIndex + 1, statuses.length, 1).setValues(statuses);
  }

  // L'alerte UI a été retirée. On peut ajouter un log pour marquer la fin.
  console.log(
    `Traitement terminé. Réussies: ${successfulUpdates}, Échecs: ${failedUpdates}, Ignorées: ${skippedCount}.`
  );
}

/**
 * Formate un objet Date JavaScript en une chaîne "YYYY-MM-DD" pour l'API HubSpot.
 * @param {Date} date L'objet Date à formater.
 * @return {string} La date formatée.
 */
function formatDateForHubSpot(date) {
  return date.toISOString().split('T')[0];
}