

// --- CONFIGURATION ---

const COMPANY_ID_HEADER = "Company ID (HS)"; // <<-- NOM DE LA COLONNE AVEC L'ID DE L'ENTREPRISE
const DEVICE_ID_COLUMN_HEADER = "DeviceID"; // <<-- NOM DE LA COLONNE OÙ ÉCRIRE L'ID DU DEVICE
const SHEET_NAME = 'Leasing'
const  HUBSPOT_PRIVATE_APP_ACCESS_TOKEN = 'HB_API_KEY'

// --- CONFIGURATION HUBSPOT ---
// Remplacez par votre jeton d'accès privé

// ID du type d'objet personnalisé "Device". À vérifier dans votre portail HubSpot.
const HUBSPOT_DEVICE_OBJECT_TYPE_ID = "2-37633992";
const HUBSPOT_COMPANY_OBJECT_TYPE_ID = "companies";

// Taille des lots pour les appels API
const BATCH_SIZE = 10; // Traitement par lot de 10 entreprises

/**
 * Fonction principale optimisée qui traite toutes les entreprises en lots.
 * Les alertes UI sont remplacées par des logs.
 */
function fetchAndWriteDeviceIdsOptimized() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`Erreur: La feuille "${SHEET_NAME}" est introuvable.`);
    return;
  }
  
  if (!HUBSPOT_PRIVATE_APP_ACCESS_TOKEN || HUBSPOT_PRIVATE_APP_ACCESS_TOKEN.includes("xxxx")) {
    Logger.log("Erreur: Le jeton d'accès HubSpot n'est pas configuré. Veuillez le modifier dans le script.");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const companyIdCol = headers.indexOf(COMPANY_ID_HEADER) + 1;
  const deviceIdCol = headers.indexOf(DEVICE_ID_COLUMN_HEADER) + 1;

  if (companyIdCol === 0) {
    Logger.log(`Erreur: Colonne "${COMPANY_ID_HEADER}" introuvable.`);
    return;
  }
  if (deviceIdCol === 0) {
    Logger.log(`Erreur: Colonne "${DEVICE_ID_COLUMN_HEADER}" introuvable.`);
    return;
  }

  // Étape 1: Lire tous les IDs de device déjà présents dans la feuille
  const deviceIdColumnValues = sheet.getRange(2, deviceIdCol, sheet.getLastRow() - 1, 1).getValues();
  const existingDeviceIdsInSheet = new Set();
  deviceIdColumnValues.forEach(row => {
    if (row[0]) {
      String(row[0]).split(',').forEach(id => {
        if(id.trim()) existingDeviceIdsInSheet.add(id.trim());
      });
    }
  });
  Logger.log(`🔍 Found ${existingDeviceIdsInSheet.size} existing device IDs in the sheet.`);

  // Étape 2: Identifier toutes les lignes à traiter
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const data = dataRange.getValues();
  const rowsToProcess = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const companyId = row[companyIdCol - 1];
    const existingDeviceIdInRow = row[deviceIdCol - 1];

    // On traite la ligne si un ID d'entreprise est présent et que la colonne DeviceID est vide
    if (companyId && !existingDeviceIdInRow) {
      rowsToProcess.push({
        rowIndex: i + 2, // +2 car on commence à la ligne 2 dans la feuille
        companyId: String(companyId).trim()
      });
    }
  }

  Logger.log(`📋 Found ${rowsToProcess.length} rows to process.`);

  if (rowsToProcess.length === 0) {
    Logger.log("Information: Aucune ligne à traiter. Toutes les entreprises ont déjà un Device ID assigné.");
    return;
  }

  Logger.log(`Lancement du script optimisé. Traitement de ${rowsToProcess.length} entreprises en cours...`);

  // Étape 3: Traitement par lots
  let processedCount = 0;
  let errorCount = 0;
  const companiesWithoutDevices = [];

  for (let i = 0; i < rowsToProcess.length; i += BATCH_SIZE) {
    const batch = rowsToProcess.slice(i, i + BATCH_SIZE);
    Logger.log(`🔄 Processing batch ${Math.floor(i/BATCH_SIZE) + 1}/${Math.ceil(rowsToProcess.length/BATCH_SIZE)} (${batch.length} companies)`);
    
    try {
      const batchResults = processBatch(batch, existingDeviceIdsInSheet);
      
      // Écriture des résultats dans la feuille
      for (const result of batchResults) {
        if (result.success && result.deviceId) {
          sheet.getRange(result.rowIndex, deviceIdCol).setValue(result.deviceId);
          existingDeviceIdsInSheet.add(result.deviceId);
          processedCount++;
          Logger.log(`✅ Company ${result.companyId} (row ${result.rowIndex}): Assigned device ${result.deviceId}`);
        } else if (result.success && !result.deviceId) {
          companiesWithoutDevices.push(result.companyId);
          Logger.log(`⚠️ Company ${result.companyId} (row ${result.rowIndex}): No devices found or all already assigned`);
        } else {
          Logger.log(`❌ Error for company ${result.companyId} at row ${result.rowIndex}: ${result.error}`);
          errorCount++;
        }
      }
      
      // Pause courte entre les lots
      Utilities.sleep(200);
      
    } catch (e) {
      Logger.log(`❌ Batch processing error: ${e.message}`);
      errorCount += batch.length;
    }
  }

  // Logging des entreprises sans devices
  if (companiesWithoutDevices.length > 0) {
    Logger.log(`⚠️ Companies without associated devices (${companiesWithoutDevices.length}): ${companiesWithoutDevices.join(', ')}`);
  }

  // Résultats finaux dans les logs
  Logger.log(`--- Opération terminée ---`);
  Logger.log(`✅ ${processedCount} ligne(s) mise(s) à jour`);
  Logger.log(`❌ ${errorCount} erreur(s)`);
  Logger.log(`⚠️ ${companiesWithoutDevices.length} entreprise(s) sans devices`);
  Logger.log(`--- Fin du rapport ---`);
}

/**
 * Traite un lot d'entreprises pour récupérer leurs devices associés.
 * @param {Array} batch Un tableau d'objets {rowIndex, companyId}
 * @param {Set} existingDeviceIdsInSheet Set des IDs déjà dans la feuille
 * @returns {Array} Tableau des résultats
 */
function processBatch(batch, existingDeviceIdsInSheet) {
  const results = [];
  
  // Étape 1: Récupération des associations pour toutes les entreprises du lot
  const associationsResults = batch.map(item => getCompanyAssociations(item.companyId));
  
  // Étape 2: Collecte de tous les device IDs uniques
  const allDeviceIds = new Set();
  associationsResults.forEach(result => {
    if (result.success && result.deviceIds) {
      result.deviceIds.forEach(id => allDeviceIds.add(id));
    }
  });
  
  // Étape 3: Récupération des détails des devices en une seule requête
  let devicesDetails = {};
  if (allDeviceIds.size > 0) {
    const devicesDetailsResult = getDevicesDetails(Array.from(allDeviceIds));
    if (devicesDetailsResult.success) {
      devicesDetails = devicesDetailsResult.devices;
    }
  }
  
  // Étape 4: Attribution des devices aux entreprises
  for (let i = 0; i < batch.length; i++) {
    const item = batch[i];
    const associationResult = associationsResults[i];
    
    if (!associationResult.success) {
      results.push({
        rowIndex: item.rowIndex,
        companyId: item.companyId,
        success: false,
        error: associationResult.error
      });
      continue;
    }
    
    if (!associationResult.deviceIds || associationResult.deviceIds.length === 0) {
      results.push({
        rowIndex: item.rowIndex,
        companyId: item.companyId,
        success: true,
        deviceId: null // Pas de device associé
      });
      continue;
    }
    
    // Trouve le plus ancien device disponible pour cette entreprise
    const availableDevices = associationResult.deviceIds
      .filter(deviceId => !existingDeviceIdsInSheet.has(deviceId))
      .map(deviceId => ({
        id: deviceId,
        createdate: devicesDetails[deviceId]?.createdate || new Date(0)
      }))
      .sort((a, b) => new Date(a.createdate) - new Date(b.createdate));
    
    if (availableDevices.length > 0) {
      results.push({
        rowIndex: item.rowIndex,
        companyId: item.companyId,
        success: true,
        deviceId: availableDevices[0].id
      });
    } else {
      results.push({
        rowIndex: item.rowIndex,
        companyId: item.companyId,
        success: true,
        deviceId: null // Tous les devices sont déjà assignés
      });
    }
  }
  
  return results;
}

/**
 * Récupère les associations d'une entreprise avec les devices.
 * @param {string} companyId L'ID de l'entreprise
 * @returns {Object} Résultat avec success et deviceIds ou error
 */
function getCompanyAssociations(companyId) {
  const url = `https://api.hubapi.com/crm/v4/objects/companies/${companyId}/associations/${HUBSPOT_DEVICE_OBJECT_TYPE_ID}`;
  
  const options = {
    method: "GET",
    headers: { "Authorization": `Bearer ${HUBSPOT_PRIVATE_APP_ACCESS_TOKEN}` },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code !== 200) {
      return { success: false, error: `API Associations [${code}]: ${body}` };
    }

    const json = JSON.parse(body);
    const deviceIds = json.results.map(result => result.toObjectId);
    
    return { success: true, deviceIds: deviceIds };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Récupère les détails de plusieurs devices en une seule requête.
 * @param {Array} deviceIds Tableau des IDs de devices
 * @returns {Object} Résultat avec success et devices ou error
 */
function getDevicesDetails(deviceIds) {
  if (deviceIds.length === 0) {
    return { success: true, devices: {} };
  }
  
  const url = `https://api.hubapi.com/crm/v3/objects/${HUBSPOT_DEVICE_OBJECT_TYPE_ID}/batch/read`;
  
  const payload = {
    "inputs": deviceIds.map(id => ({ "id": id })),
    "properties": ["createdate"]
  };

  const options = {
    method: "POST",
    contentType: "application/json",
    headers: { "Authorization": `Bearer ${HUBSPOT_PRIVATE_APP_ACCESS_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code !== 200) {
      return { success: false, error: `API Devices Batch [${code}]: ${body}` };
    }

    const json = JSON.parse(body);
    const devices = {};
    
    if (json.results) {
      json.results.forEach(device => {
        devices[device.id] = {
          createdate: device.properties.createdate
        };
      });
    }
    
    return { success: true, devices: devices };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
