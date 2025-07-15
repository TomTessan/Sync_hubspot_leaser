# Documentation des Scripts de Synchro Google Sheets vers HubSpot

Ce document contient la documentation pour trois scripts Google Apps Script conçus pour interagir avec HubSpot.

---

## Script 1 : Deals_sync.js

Ce script automatise la mise à jour des transactions (Deals) HubSpot à partir des données d'une feuille de calcul Google Sheets.

### ✨ Fonctionnalités

* **Calcul de date** : Calcule automatiquement la date de fin de contrat à partir d'une date d'installation et d'une durée.
* **Mise à jour HubSpot** : Met à jour plusieurs propriétés d'une transaction (leaser, statut, date de fin, etc.).
* **Récupération de données** : Récupère l'ID de l'entreprise associée à la transaction et l'inscrit dans la feuille.
* **Flexibilité** : Identifie les colonnes par leur nom d'en-tête, l'ordre des colonnes n'a donc pas d'importance.
* **Traitement conditionnel** : Ignore les lignes déjà traitées pour éviter les doublons.

### 🚀 Installation et Configuration

#### 1. Clé d'API HubSpot

1.  Ouvrez l'éditeur de script (`Extensions > Apps Script`).
2.  Modifiez la constante `HUBSPOT_API_KEY` avec votre jeton d'accès privé.

```javascript
// CONFIGURATION REQUISE
const HUBSPOT_API_KEY = 'pat-na1-xxxx-xxxx-xxxx-xxxxxxxxxxxx'; // <--- REMPLACEZ PAR VOTRE VRAIE CLÉ
```

#### 2. Structure de la Feuille Google Sheets

| En-tête Requis      | Description                                                    |
| :------------------ | :------------------------------------------------------------- |
| `Leaser`            | Le nom du bailleur.                                            |
| `Statut`            | Le statut du contrat (ex: `Installé`).                         |
| `Fin leasing`       | **[Sortie]** La date de fin de contrat, calculée par le script. |
| `Durée (Mois)`      | La durée du contrat en mois (ex: `36`).                        |
| `Id Transac ADV`    | L'ID de la transaction HubSpot (ID seul ou URL complète).      |
| `Company ID (HS)`   | **[Sortie]** L'ID de l'entreprise associée, récupéré par le script. |
| `Sync`              | Colonne de statut. Mettre `Traité` pour ignorer la ligne.      |
| `Clients`           | Le nom du client.                                              |
| `Date installatin`  | La date de début/installation du contrat.                      |

### 🛠️ Utilisation

1.  Ouvrez votre feuille de calcul.
2.  Allez dans `Extensions > Apps Script`.
3.  Sélectionnez la fonction `processTransactionsFromSheet` et cliquez sur **▶ Exécuter**.

---

## Script 2 : Device_sync.js

Ce script récupère les IDs des objets personnalisés "Device" associés à une entreprise et les inscrit dans la feuille. Il est optimisé pour traiter un grand volume de données et les ordonner par date de création du device depuis hubspot.

### ✨ Fonctionnalités

* **Récupération d'associations** : Trouve tous les objets "Device" liés à un ID d'entreprise.
* **Logique d'attribution "en cascade"** : Pour chaque entreprise, il trouve le plus ancien "Device" associé qui n'a pas encore été assigné dans la feuille.
* **Traitement par lots (Batch)** : Traite les entreprises par lots pour éviter les erreurs de timeout.
* **Prévention des doublons** : Analyse les IDs déjà présents pour ne pas les réassigner.

### 🚀 Installation et Configuration

1.  Ouvrez l'éditeur de script et ajustez les constantes suivantes :

```javascript
const COMPANY_ID_HEADER = "Company ID (HS)"; // Colonne avec l'ID de l'entreprise
const DEVICE_ID_COLUMN_HEADER = "DeviceID"; // Colonne où écrire l'ID du device
const SHEET_NAME = 'Leasing'; // Nom de l'onglet à traiter
const HUBSPOT_PRIVATE_APP_ACCESS_TOKEN = 'pat-na1-xxxx-xxxx-xxxx-xxxxxxxxxxxx'; // <--- REMPLACEZ PAR VOTRE VRAIE CLÉ
```

2.  Assurez-vous que votre feuille contient les colonnes suivantes :

| En-tête Requis    | Description                                                        |
| :---------------- | :----------------------------------------------------------------- |
| `Company ID (HS)` | **[Entrée]** L'ID de l'entreprise, rempli par le Script 1.          |
| `DeviceID`        | **[Sortie]** La colonne où le script inscrira l'ID du "Device". |

### 🛠️ Utilisation

1.  Assurez-vous que la colonne `Company ID (HS)` est remplie.
2.  Allez dans `Extensions > Apps Script`.
3.  Sélectionnez la fonction `fetchAndWriteDeviceIdsOptimized` et cliquez sur **▶ Exécuter**.

---

## Script 3 : Date_sync.js

Ce script met à jour deux propriétés de date sur un objet "Device" dans HubSpot en utilisant les données de la feuille.

### ✨ Fonctionnalités

* **Mise à jour ciblée** : Met à jour les propriétés `date_de_livraison_effective__d_` et `date_de_fin_de_contrat_temporaire__d_`.
* **Menu personnalisé** : Ajoute un menu `Actions HubSpot` dans l'interface de Google Sheets pour un accès facile.
* **Gestion de statut** : Utilise une colonne `Sync` pour marquer les lignes traitées ou en erreur et éviter de les retraiter.
* **Robuste** : Vérifie la présence des colonnes requises avant de s'exécuter.

### 🚀 Installation et Configuration

 Vérifiez que votre feuille contient les en-têtes suivants :

| En-tête Requis     | Description                                                         |
| :----------------- | :------------------------------------------------------------------ |
| `DeviceID`         | **[Entrée]** L'ID du "Device" à mettre à jour, rempli par le Script 2. |
| `Date installatin` | **[Entrée]** La date utilisée pour la mise à jour.                  |
| `Fin leasing`      | **[Entrée]** La date utilisée pour la mise à jour.                  |
| `Sync`             | **[Entrée/Sortie]** Colonne de statut pour le suivi du traitement.  |

### 🛠️ Utilisation

1.  Ouvrez ou actualisez votre feuille de calcul. Un nouveau menu "Actions HubSpot" devrait apparaître.
2.  Cliquez sur `Actions HubSpot > Mettre à jour les dates des Devices` pour lancer le script.
3.  Le statut de chaque ligne sera mis à jour dans la colonne `Sync`.
