# Script de Synchro Google Sheets vers HubSpot

Ce script Google Apps Script automatise la mise à jour des transactions (Deals) HubSpot à partir des données d'une feuille de calcul Google Sheets.

## ✨ Fonctionnalités

* **Calcul de date** : Calcule automatiquement la date de fin de contrat à partir d'une date d'installation et d'une durée.
* **Mise à jour HubSpot** : Met à jour plusieurs propriétés d'une transaction (leaser, statut, date de fin, etc.).
* **Récupération de données** : Récupère l'ID de l'entreprise associée à la transaction et l'inscrit dans la feuille.
* **Flexibilité** : Identifie les colonnes par leur nom d'en-tête, l'ordre des colonnes n'a donc pas d'importance.
* **Traitement conditionnel** : Ignore les lignes déjà traitées pour éviter les doublons.

## 🚀 Installation et Configuration

Suivez ces deux étapes pour rendre le script opérationnel.

### 1. Clé d'API HubSpot

La communication avec HubSpot nécessite un jeton d'accès.

1. Ouvrez l'éditeur de script (`Extensions > Apps Script`).
2. Modifiez la constante `HUBSPOT_API_KEY` en y ajoutant votre jeton d'accès privé HubSpot.

```javascript
// CONFIGURATION REQUISE
const HUBSPOT_API_KEY = 'pat-na1-xxxx-xxxx-xxxx-xxxxxxxxxxxx'; // <--- REMPLACEZ PAR VOTRE VRAIE CLÉ
```

### 2. Structure de la Feuille Google Sheets

Le script requiert que la première ligne de votre feuille contienne des en-têtes spécifiques. Si un en-tête est manquant, le script ne s'exécutera pas.

| En-tête Requis | Description |
| :--- | :--- |
| `Leaser` | Le nom du bailleur. |
| `Statut` | Le statut du contrat (ex: `Installé`). |
| `Fin leasing` | **[Sortie]** La date de fin de contrat, calculée par le script. |
| `Durée (Mois)`| La durée du contrat en mois (ex: `36`). |
| `Id Transac ADV`| L'ID de la transaction HubSpot (ID seul ou URL complète). |
| `Company ID (HS)`| **[Sortie]** L'ID de l'entreprise associée, récupéré par le script. |
| `Sync` | Colonne de statut. Mettre `Traité` pour ignorer la ligne. |
| `Clients` | Le nom du client. |
| `Date installatin`| La date de début/installation du contrat. |

## 🛠️ Utilisation

1. Ouvrez votre feuille de calcul Google Sheets.
2. Allez dans `Extensions > Apps Script`.
3. Dans l'éditeur, assurez-vous que la fonction `processTransactionsFromSheet` est sélectionnée dans la barre d'outils.
4. Cliquez sur **▶ Exécuter**.
5. Lors de la première exécution, autorisez le script à accéder à vos données Google Sheets et à se connecter à des services externes.

## ⚙️ Logique de Traitement

* Le script parcourt chaque ligne de la feuille.
* Une ligne est **ignorée** si la colonne `Sync` contient la valeur `Traité` ou si la colonne `Id Transac ADV` est vide.
* En cas d'erreur sur une ligne, un message est inscrit dans la colonne `Company ID (HS)` et des détails sont disponibles dans les journaux d'exécution (`Extensions > Apps Script > Exécutions`).
