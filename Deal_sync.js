
// ===============================================================
// CONFIGURATION REQUISE
// ===============================================================
const HUBSPOT_API_KEY = 'HB_API_KEY'; // Remplacez par votre clé si nécessaire



/**
 * Fonction principale qui parcourt la feuille pour traiter chaque transaction.
 */
function processTransactionsFromSheet() {
  const ui = SpreadsheetApp.getUi();
  if (!HUBSPOT_API_KEY || HUBSPOT_API_KEY === 'VOTRE_JETON_HUBSPOT_ICI') {
    Logger.log("Configuration requise: Veuillez renseigner votre jeton d'accès HubSpot.");
    ui.alert("Clé API HubSpot manquante. Veuillez la configurer dans le script.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const allData = dataRange.getValues();
  
  const headers = allData.shift(); 
  const data = allData;
  const startRow = 2;

  const HEADER_NAMES = {
    leaser: 'Leaser',
    statut: 'Statut',
    finLeasing: 'Fin leasing', // Colonne cible pour l'écriture de la date
    dureeMois: 'Durée (Mois)',
    idTransac: 'Id Transac ADV',
    companyIdOutput: 'Company ID (HS)',
    sync: 'Sync',
    client : 'Clients',
    dateInstallation: 'Date installatin'
  };

  const colIndices = {};
  for (const key in HEADER_NAMES) {
    colIndices[key] = headers.indexOf(HEADER_NAMES[key]);
  }
  
  const missingHeaders = Object.keys(colIndices).filter(key => colIndices[key] === -1);
  if (missingHeaders.length > 0) {
    const missingNames = missingHeaders.map(key => HEADER_NAMES[key]).join(', ');
    ui.alert(`Action annulée. Colonne(s) manquante(s) : ${missingNames}. Veuillez vérifier les en-têtes.`);
    return;
  }

  data.forEach((row, index) => {
    const currentRow = startRow + index;

    const syncStatus = row[colIndices.sync];
    if (syncStatus === 'Traité') {
      Logger.log(`Ligne ${currentRow}: Statut "Traité". Ligne ignorée.`);
      return;
    }

    const hubspotInfo = row[colIndices.idTransac];
    sheet.getRange(currentRow, colIndices.companyIdOutput + 1).clearContent();

    if (!hubspotInfo) {
      Logger.log(`Ligne ${currentRow}: Pas d'ID de transaction. Ligne ignorée.`);
      return;
    }

    try {
      const dealId = extractDealIdFromUrl(hubspotInfo.toString()); 
      if (!dealId) {
        throw new Error(`Impossible d'extraire l'ID depuis "${hubspotInfo}"`);
      }

      Logger.log(`Ligne ${currentRow}: Traitement de la transaction ID ${dealId}`);

      // --- NOUVELLE LOGIQUE DE DATE ---
      const dateInstallation = row[colIndices.dateInstallation];
      const dureeMois = row[colIndices.dureeMois];
      let finalDateForHubspot = null;

      // Si une date d'installation et une durée existent, on calcule la nouvelle date
      if (dateInstallation instanceof Date && dureeMois && !isNaN(parseInt(dureeMois, 10))) {
        let calculatedDate = new Date(dateInstallation.getTime());
        calculatedDate.setMonth(calculatedDate.getMonth() + parseInt(dureeMois, 10));
        calculatedDate.setMonth(calculatedDate.getMonth() + 1, 1); // Arrondi au 1er du mois suivant
        
        // **ACTION : Inscrire la nouvelle date dans la colonne 'Fin leasing'**
        sheet.getRange(currentRow, colIndices.finLeasing + 1).setValue(calculatedDate);
        Logger.log(`Ligne ${currentRow}: Nouvelle date "${calculatedDate.toLocaleDateString()}" inscrite dans la colonne '${HEADER_NAMES.finLeasing}'.`);
        
        finalDateForHubspot = calculatedDate; // On garde l'objet Date pour la fonction suivante
      }

      // 1. Met à jour les propriétés HubSpot
      const propertyUpdateErrors = updateDealPropertiesIndividually(dealId, row, colIndices, finalDateForHubspot);

      // 2. Récupère l'ID de la société associée
      try {
        const companyId = getAssociatedCompanyId(dealId);
        const companyIdCell = sheet.getRange(currentRow, colIndices.companyIdOutput + 1);
        if (companyId) {
          companyIdCell.setValue(companyId);
          Logger.log(`Ligne ${currentRow}: ID Société ${companyId} inscrit.`);
        } else {
          companyIdCell.setValue("Non trouvée");
          Logger.log(`Ligne ${currentRow}: Aucune société associée trouvée.`);
        }
      } catch (e) {
        sheet.getRange(currentRow, colIndices.companyIdOutput + 1).setValue(`Erreur récup. ID`);
        Logger.log(`Ligne ${currentRow}: ERREUR - ${e.message}`);
      }
      
      if (propertyUpdateErrors.length > 0) {
        const errorString = propertyUpdateErrors.join('; ');
        Logger.log(`Ligne ${currentRow}: Erreurs de MàJ: ${errorString}`);
      }

    } catch (e) {
      Logger.log(`Ligne ${currentRow}: ERREUR FATALE - ${e.message}`);
      sheet.getRange(currentRow, colIndices.companyIdOutput + 1).setValue(`ERREUR FATALE`);
    }
  });

  Logger.log("Mise à jour HubSpot terminée. Vérifiez les résultats dans la feuille et les journaux d'exécution.");
}


/**
 * Extrait l'ID de la transaction (dealId) à partir d'une URL HubSpot ou d'une chaîne.
 */
function extractDealIdFromUrl(input) {
  const urlMatch = input.match(/record\/0-3\/(\d+)/);
  if (urlMatch) return urlMatch[1];
  const idMatch = input.match(/^\d+$/);
  if (idMatch) return idMatch[0];
  return null;
}

/**
 * Met à jour les propriétés d'une transaction via l'API HubSpot.
 * @param {string} dealId L'ID de la transaction.
 * @param {Array} row La ligne de données.
 * @param {Object} colIndices L'objet contenant les index des colonnes.
 * @param {Date|null} calculatedEndDate La date de fin calculée (peut être null).
 * @returns {Array<string>} Une liste des messages d'erreur.
 */
function updateDealPropertiesIndividually(dealId, row, colIndices, calculatedEndDate) {
  const url = `https://api.hubapi.com/crm/v3/objects/deals/${dealId}`;
  let updateErrors = [];
  let dateFinContratStr = null;

  // On utilise la date calculée si elle existe, sinon on se rabat sur la valeur de la colonne
  if (calculatedEndDate instanceof Date) {
    dateFinContratStr = calculatedEndDate.toISOString().split('T')[0]; // Format YYYY-MM-DD
  } else {
    const dateFinLeasing = row[colIndices.finLeasing];
    if (dateFinLeasing instanceof Date) {
      dateFinContratStr = dateFinLeasing.toISOString().split('T')[0];
    }
  }

  let leaserValue = row[colIndices.leaser];
  if (leaserValue === 'Achat Direct') leaserValue = 'Achat direct';

  let statutContratValue = row[colIndices.statut];
  if (statutContratValue === 'Installé') statutContratValue = 'livré';

  const propertiesToUpdate = {
    "leaser": leaserValue,
    "statut_du_contrat": statutContratValue,
    "duree_maintenance_garantie": row[colIndices.dureeMois],
    "date_de_fin_de_contrat": dateFinContratStr,
    "nomenclature_adv_modifie" : row[colIndices.client]
  };

  for (const propertyName in propertiesToUpdate) {
    const propertyValue = propertiesToUpdate[propertyName];
    if (propertyValue === null || propertyValue === undefined || propertyValue === '') {
      continue;
    }

    const payload = { properties: { [propertyName]: propertyValue } };
    const options = {
      method: 'patch',
      contentType: 'application/json',
      headers: { 'Authorization': `Bearer ${HUBSPOT_API_KEY}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      const responseBody = response.getContentText();
      let errorMessage = `Propriété '${propertyName}'`;
      try {
        errorMessage += `: ${JSON.parse(responseBody).message || responseBody}`;
      } catch (e) {
        errorMessage += `: Erreur de réponse du serveur.`;
      }
      updateErrors.push(errorMessage);
    }
  }
  return updateErrors;
}

/**
 * Récupère l'ID de la première société associée à une transaction.
 */
function getAssociatedCompanyId(dealId) {
  const url = `https://api.hubapi.com/crm/v3/objects/deals/${dealId}?associations=company`;
  const options = {
    method: 'get',
    headers: { 'Authorization': `Bearer ${HUBSPOT_API_KEY}` },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`Échec récupération associations. [${responseCode}]: ${responseBody}`);
  }

  const dealData = JSON.parse(responseBody);
  const associations = dealData.associations;

  if (associations && associations.companies && associations.companies.results.length > 0) {
    return associations.companies.results[0].id;
  }

  return null;
}