# Vue d'ensemble

**CegedimToSQL.ps1** est un script PowerShell de synchronisation des données d'équipements depuis l'API Ivanti/HEAT de Cegedim vers une base de données SQL Server.

Le script interroge l'API REST OData de Cegedim pour récupérer les informations des équipements (CI - Configuration Items) et synchronise ces données avec une base SQL Server.

* Source de données provenant de l'**API REST OData Ivanti/HEAT** (Cegedim Cloud)
* Synchronisation vers une base de données **SQL Server**
* Gestion de la pagination automatique des résultats API
* Support de deux modes de transaction SQL : **AllInOne** ou **OneByOne**

## Prérequis

* PowerShell 7 ou supérieur
* Accès réseau à l'API Ivanti/HEAT de Cegedim
* Accès à une base de données SQL Server
* Token d'authentification API valide
* Modules PowerShell requis (chargés automatiquement depuis le dossier `Modules/`) :
  - **Ini.ps1** : Gestion des fichiers de configuration .ini
  - **Log.ps1** : Gestion des logs et traces
  - **Encode.ps1** : Gestion de l'encodage (UTF-8) et décodage des mots de passe
  - **SendEmail.ps1** : Envoi d'emails de notification
  - **StrConvert.ps1** : Conversion et manipulation de chaînes (dates, nombres)
  - **SQLServer - TransactionAllInOne.ps1** ou **SQLServer - TransactionOneByOne.ps1** : Gestion des transactions SQL Server (chargé dynamiquement selon configuration)

## Structure du script

Le script suit une architecture modulaire avec :
1. **Chargement des modules de base** : Import des modules PowerShell essentiels (Log, Ini, Encode, StrConvert, SendEmail)
2. **Initialisation** : Chargement du fichier de configuration `.ini`
3. **Configuration de l'encodage** : Passage de la console en UTF-8
4. **Chargement du module SQL** : Import dynamique du module de transaction SQL selon le paramètre `[start][TransacSQL]`
5. **Récupération des données** : Interrogation de l'API Cegedim et de la base SQL
6. **Synchronisation** : Mise à jour de la base SQL avec les données de l'API
7. **Finalisation** : Génération des logs et envoi optionnel d'email

## Source des données

|                   Paramètres .ini                   |                    Source                           |
| --------------------------------------------------- | --------------------------------------------------- |
| [**URL**][Workstations]                             | API Ivanti/HEAT - Postes de travail                |
| [**URL**][MobileDevices]                            | API Ivanti/HEAT - Appareils mobiles                |
| [**URL**][Routers]                                  | API Ivanti/HEAT - Routeurs                         |
| [**URL**][TechnicalApplications]                    | API Ivanti/HEAT - Applications techniques          |
| [**URL**][RSAs]                                     | API Ivanti/HEAT - RSAs                             |
| [**URL**][VOIPs]                                    | API Ivanti/HEAT - Téléphones VoIP                  |
| [**URL**][VideoConferences]                         | API Ivanti/HEAT - Équipements de visioconférence   |
| [**URL**][Servers]                                  | API Ivanti/HEAT - Serveurs                         |
| [**URL**][Services]                                 | API Ivanti/HEAT - Services                         |
| [**URL**][VirtualWorkstations]                      | API Ivanti/HEAT - Postes de travail virtuels       |
| [**URL**][EnterpriseApplications]                   | API Ivanti/HEAT - Applications d'entreprise        |

## Base de données cible

|                   Paramètres .ini                   |            Description                |
| --------------------------------------------------- | ------------------------------------- |
| [**SQL_Server**][server]                            | Serveur SQL Server                    |
| [**SQL_Server**][database]                          | Nom de la base de données             |
| [**SQL_Server**][table]                             | Table de destination (CI)             |
| [**SQL_Server**][login]                             | Login de connexion SQL                |
| [**SQL_Server**][password]                          | Mot de passe (encodé)                 |

## Principe du traitement

Le script interroge l'API REST OData de Cegedim pour récupérer les données des équipements (CI), puis synchronise ces données avec une table SQL Server.

### Flux de traitement

1. **LoadIni** : Charge le fichier de configuration `CegedimToSQL.ini`
   - Initialise les variables de log, d'erreur et de modification
   - Vérifie l'existence du fichier de configuration
   - Crée les fichiers de log nécessaires
   - Initialise la hashtable `$script:CI` pour stocker les données API

2. **SetConsoleToUFT8** : Configure l'encodage de la console en UTF-8

3. **Chargement du module SQL** : Charge dynamiquement le module de transaction SQL
   - Si `[start][TransacSQL] = "AllInOne"` : charge `SQLServer - TransactionAllInOne.ps1`
   - Si `[start][TransacSQL] = "OneByOne"` : charge `SQLServer - TransactionOneByOne.ps1`

4. **Query_BDD_CI** : Charge en mémoire le contenu de la table SQL
   - Interroge la table `[SQL_Server][table]` (CI)
   - Indexe les données par le champ `RecID`
   - Stocke le résultat dans la hashtable `$script:BDDCI`

5. **Query_type** : Récupère les données depuis l'API Ivanti/HEAT (appelé pour chaque URL)
   - Effectue des requêtes HTTP GET avec authentification par token
   - Gère la pagination automatique (25 enregistrements par page)
   - Applique un tri par `RecId` pour garantir la cohérence
   - Implémente un mécanisme de retry (3 tentatives avec 2 secondes d'attente)
   - Convertit les champs de type date au format spécifié dans `[SQL_Server][frmtdateOUT]`
   - Convertit les champs numériques décimaux (2 ou 4 décimales selon le type)
   - Filtre les champs à ignorer (liens OData, champs techniques)
   - Stocke les données dans la hashtable `$script:CI` indexée par `RecId`

6. **Update_BDD_CI** : Synchronise la base SQL avec les données API
   - Compare `$script:CI` (données API) avec `$script:BDDCI` (données SQL)
   - Identifie les enregistrements à créer, modifier ou supprimer
   - Applique les modifications selon le paramètre `[start][ApplyUpdate]`
   - Recharge les données SQL après mise à jour

7. **QUIT** : Finalise le traitement
   - Affiche le résumé des erreurs et warnings
   - Envoie un email de notification si configuré

**Nota important** : *Le script peut fonctionner en mode simulation (`[start][ApplyUpdate] = no`) pour tester sans modifier la base de données.*

## Utilisation

### Exécution du script

```powershell
# Depuis le répertoire du script
.\CegedimToSQL.ps1

# Ou avec un chemin absolu
pwsh -File "C:\chemin\vers\CegedimToSQL.ps1"
```

Le script recherche automatiquement le fichier de configuration `CegedimToSQL.ini` dans le même répertoire.

### Configuration

Avant la première exécution, configurer le fichier `CegedimToSQL.ini` :
1. Définir le token d'authentification API dans la section `[Cegedim]`
2. Vérifier les URLs des endpoints API dans la section `[URL]`
3. Configurer les paramètres de connexion SQL Server dans la section `[SQL_Server]`
4. Choisir le mode de transaction SQL avec le paramètre `[start][TransacSQL]` (AllInOne ou OneByOne)
5. Activer ou désactiver les mises à jour avec le paramètre `[start][ApplyUpdate]` (yes ou no)
6. Configurer les options de log dans la section `[start]`
7. Configurer l'envoi d'email dans la section `[email]` (optionnel)

### Variables d'environnement

Le script utilise `$rootpath$` comme variable de substitution dans le fichier `.ini` :
- Par défaut : répertoire du script (`$PSScriptRoot`)
- Peut être redéfini dans `[intf][rootpath]`

# Détail des fonctions principales

### LoadIni

Charge et initialise le fichier de configuration `CegedimToSQL.ini`.

**Actions réalisées :**
- Initialise les variables de log (`$script:pathfilelog`, `$script:pathfileerr`, `$script:pathfilemod`)
- Vérifie l'existence du fichier `.ini`
- Charge les paramètres via la fonction `Add-IniFiles`
- Résout les chemins des fichiers avec `GetFilePath`
- Supprime le fichier log One_Shot s'il existe (nouveau traitement)
- Crée les fichiers de log nécessaires
- Décode le token d'authentification API
- Initialise la hashtable `$script:CI` pour stocker les données API

### Query_type

Interroge l'API REST OData Ivanti/HEAT pour récupérer les données d'équipements.

**Paramètres :**
- `$url` : URL de l'endpoint API à interroger
- `[switch]$Debug` : Active le mode debug (optionnel)

**Traitement spécifique :**
- Utilise `System.Net.Http.HttpClient` pour les requêtes HTTP
- Ajoute l'en-tête d'authentification : `Authorization: rest_api_key=<token>`
- Gère la pagination automatique avec `$top`, `$skip` et `$orderby=RecId`
- Taille de page fixe : 25 enregistrements (imposé par l'API)
- Limite de sécurité : 1000 pages maximum
- Mécanisme de retry : 3 tentatives avec 2 secondes d'attente entre chaque
- Détection des doublons pour éviter les boucles infinies
- Conversion automatique des champs de type date
- Conversion automatique des champs numériques décimaux
- Filtrage des champs à ignorer (définis dans `$ignore`)
- Stocke les données dans `$script:CI` indexée par `RecId`

**Champs ignorés :**
- Liens OData : `A_CommandeLink`, `AssetLink`, `DefaultSLPLink`, etc.
- Champ de pagination : `pPage`

**Champs de type date :**
- `A_DateLivraison`, `A_DateInventaireManuel`, `LastModDateTime`, `BIOSDate`, etc.
- Convertis au format défini dans `[SQL_Server][frmtdateOUT]`

**Champs numériques décimaux :**
- 2 décimales : `TotalMemory`, `CPUSpeed`, `CIVersion`
- 4 décimales : `TargetAvailability`

### Query_BDD_CI

Charge en mémoire le contenu de la table SQL Server.

**Traitement spécifique :**
- Appelle la fonction utilitaire `Query_BDDTable`
- Interroge la table définie dans `[SQL_Server][table]`
- Indexe les données par le champ `RecID`
- Applique le format de date défini dans `[SQL_Server][frmtdateOUT]`
- Stocke le résultat dans la hashtable `$script:BDDCI`

### Update_BDD_CI

Synchronise la base SQL Server avec les données de l'API.

**Traitement spécifique :**
- Appelle la fonction utilitaire `Update_BDDTable`
- Compare `$script:CI` (source API) avec `$script:BDDCI` (cible SQL)
- Utilise le champ `RecID` comme clé de comparaison
- Identifie les enregistrements à créer, modifier ou supprimer
- Applique les modifications selon le paramètre `[start][ApplyUpdate]`
- Recharge les données SQL après mise à jour via `Query_BDD_CI`

### Query_BDDTable (fonction utilitaire)

Fonction générique pour interroger une table SQL Server.

**Paramètres :**
- `$tableName` : Nom de la table à interroger
- `$functionName` : Nom de la fonction appelante (pour les logs)
- `$keyColumns` : Colonnes utilisées comme clé d'indexation
- `$targetVariable` : Hashtable cible pour stocker les résultats
- `[switch]$UseFrmtDateOUT` : Active la conversion des dates

**Actions réalisées :**
- Récupère les paramètres de connexion via `Get-BDDConnectionParams`
- Vide la hashtable cible
- Appelle la fonction `QueryTable` du module SQL
- Copie les résultats dans la hashtable cible

### Update_BDDTable (fonction utilitaire)

Fonction générique pour mettre à jour une table SQL Server.

**Paramètres :**
- `$sourceData` : Hashtable source (données API)
- `$targetData` : Hashtable cible (données SQL)
- `$keyColumns` : Colonnes utilisées comme clé de comparaison
- `$tableName` : Nom de la table à mettre à jour
- `$functionName` : Nom de la fonction appelante (pour les logs)
- `$reloadFunction` : Scriptblock pour recharger les données après mise à jour

**Actions réalisées :**
- Récupère les paramètres de connexion via `Get-BDDConnectionParams`
- Appelle la fonction `UpdateTable` du module SQL
- Exécute le scriptblock de rechargement si fourni

### Get-BDDConnectionParams (fonction utilitaire)

Retourne les paramètres de connexion à la base SQL Server.

**Retourne une hashtable contenant :**
- `server` : Serveur SQL Server
- `database` : Nom de la base de données
- `login` : Login de connexion
- `password` : Mot de passe décodé
- `datefrmtout` : Format de date pour la sortie

# Fichiers de logs

Les fichiers de logs sont définis dans la section `[intf]` du fichier `.ini`.

## CegedimToSQL_One_Shot.log

Contient les logs du dernier traitement.

**Comportement :**
- Supprimé automatiquement au démarrage de chaque nouveau traitement
- Contient toutes les traces d'exécution (LOG, DBG, WARN, ERR)
- Paramètre `[start][logtoscreen]` : contrôle l'affichage dans la console
- Paramètre `[start][debug]` : contrôle l'affichage des messages de debug

## CegedimToSQL_Cumul.err

Contient le cumul des erreurs constatées dans tous les traitements.

**Comportement :**
- Conservé entre les exécutions (mode cumulatif)
- Contient uniquement les erreurs (ERR)
- Paramètre `[start][warntoerr]` : permet d'inclure aussi les warnings (WARN)

## CegedimToSQL_Cumul.mod

Contient le cumul des modifications appliquées à la base SQL.

**Comportement :**
- Conservé entre les exécutions (mode cumulatif)
- Contient les enregistrements créés, modifiés ou supprimés
- Permet de tracer l'historique des synchronisations

# Exemple de fichier .ini

```ini
# -----------------------------------------------------------------------------------------------------------------------------
#    CegedimToSQL.ini - Necessite Powershell 7 ou +
#      Ce script met à jour la base SQL CegedimToSQL avec les données Clouds de Cegedim
# -----------------------------------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------
#     Parametrage du comportement de l'interface CegedimToSQL.ps1
# -------------------------------------------------------------------

[start]
# Le parametre "ApplyUpdate" yes/no : permet de simuler sans modifier la base CegedimToSQL si ApplyUpdate = no
ApplyUpdate = yes

# TransacSQL : OneByOne ou AllInOne
TransacSQL = AllInOne

# Le parametre "logtoscreen" contrôle l'affichage de toutes les infos de log/error/warning dans la console
logtoscreen = yes

# Le parametre "debug" contrôle l'affichage des infos de debug dans la console
debug       = no

# Le parametre "warntoerr" permet d'inclure ou pas les warnings dans le fichier CegedimToSQL_Cumul.err
warntoerr   = yes

# -------------------------------------------------------------------
#     Chemin des fichiers de LOGS
# -------------------------------------------------------------------
[intf]
name        = Synchronisation Clouds Cegedim >> Base SQL CegedimToSQL

# Chemin du fichier log : 
pathfilelog = $rootpath$\logs\CegedimToSQL_One_Shot.log

# Chemin du fichier logs d'erreur
pathfileerr = $rootpath$\logs\CegedimToSQL_Cumul.err

# Chemin du fichier des logs modifications
pathfilemod = $rootpath$\logs\CegedimToSQL_Cumul.mod

[Cegedim]
Token = xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

[URL]
Workstations           = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__Workstations
MobileDevices          = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__MobileDevices
Routers                = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__Routers
TechnicalApplications  = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__TechnicalApplications
RSAs                   = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__RSAs
VOIPs                  = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__VOIPs
VideoConferences       = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__VideoConferences
Servers                = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__Servers
Services               = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__Services
VirtualWorkstations    = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__VirtualWorkstations
EnterpriseApplications = https://fortil.cegedim.com/HEAT/api/odata/businessobject/CI__EnterpriseApplications

# -------------------------------------------------------------------
#     Parametrage du serveur SQL
# -------------------------------------------------------------------

# Parametre de connection à la base Ivanti
[SQL_Server]                                                                       
frmtdateOUT  = yyyy-dd-MM HH:mm:ss
server       = WIN-09T11CB4M65\TEST
database     = ISMIvantiSM
login        = sa
password     = !Plmuvimvmhpb2
table        = CI

# -------------------------------------------------------------------
#     Parametrage des Emails
# -------------------------------------------------------------------

# Parametre pour l'envoi de mails (Protocoles possible : Microsoft.Graph ou SMTP)
# Le parametre "emailmode" permet de choisir le mode d'emission d'un mail (GRAPH ou SMTP)
# Envoi de mail si sendemail = "yes" / "no"
[email]
sendemail    = no
expediteur   = btran56@gmail.com
destinataire = btran56@gmail.com
Subject      = Synchro Cegedim vers SQL
emailmode    = SMTP
UseSSL       = false
server       = smtp.gmail.com
port         = 
password     = 
```

# Modules PowerShell

Le script charge automatiquement les modules suivants depuis le répertoire `Modules/` :

## Modules chargés au démarrage

Ces modules sont chargés systématiquement au démarrage du script :

| Module | Description | Utilisation dans le script |
|--------|-------------|----------------------------|
| **Log.ps1** | Système de logging | Fonctions LOG, DBG, WARN, ERR, QUIT |
| **Ini.ps1** | Gestion des fichiers .ini | Chargement et parsing du fichier de configuration |
| **Encode.ps1** | Gestion de l'encodage | Configuration UTF-8 de la console et décodage des mots de passe |
| **StrConvert.ps1** | Conversion de chaînes | Conversion de dates et nombres, manipulation de données |
| **SendEmail.ps1** | Envoi d'emails | Notification de fin de traitement (optionnel) |

**Chargement des modules de base :**
```powershell
. "$PSScriptRoot\Modules\Log.ps1" > $null 
. "$PSScriptRoot\Modules\Ini.ps1" > $null 
. "$PSScriptRoot\Modules\Encode.ps1" > $null 
. "$PSScriptRoot\Modules\StrConvert.ps1" > $null 
. "$PSScriptRoot\Modules\SendEmail.ps1" > $null 
```

## Modules chargés dynamiquement

Ces modules sont chargés conditionnellement selon la configuration :

| Module | Condition de chargement | Description |
|--------|------------------------|-------------|
| **SQLServer - TransactionAllInOne.ps1** | `[start][TransacSQL] = "AllInOne"` | Gestion des transactions SQL en mode "tout en une fois" (plus rapide) |
| **SQLServer - TransactionOneByOne.ps1** | `[start][TransacSQL] = "OneByOne"` | Gestion des transactions SQL en mode "une par une" (plus sûr) |

**Chargement dynamique du module SQL :**
```powershell
if ($script:cfg["start"]["TransacSQL"] -eq "AllInOne" ) {
    . "$PSScriptRoot\Modules\SQLServer - TransactionAllInOne.ps1" > $null
} else {
    . "$PSScriptRoot\Modules\SQLServer - TransactionOneByOne.ps1" > $null
}
```

**Choix du mode de transaction :**
- **AllInOne** : Toutes les modifications sont regroupées dans une seule transaction SQL. Plus rapide mais nécessite plus de mémoire.
- **OneByOne** : Chaque modification est exécutée dans une transaction séparée. Plus lent mais plus robuste en cas d'erreur.

Les modules sont chargés en mode "dot sourcing" (`.`) pour rendre leurs fonctions disponibles dans le scope du script principal.

**Note :** La documentation détaillée de chaque module sera fournie séparément.

## Modules disponibles mais non utilisés

Ces modules sont présents dans le répertoire `Modules/` mais ne sont pas utilisés par le script actuel :

| Module | Description |
|--------|-------------|
| **Csv.ps1** | Gestion des fichiers CSV |
| **FileTools.ps1** | Outils de manipulation de fichiers |
| **XLSX.ps1** | Lecture de fichiers Excel |
| **PostgreSQL - TransactionAllInOne.ps1** | Gestion des transactions PostgreSQL (mode AllInOne) |
| **PostgreSQL - TransactionOneByOne.ps1** | Gestion des transactions PostgreSQL (mode OneByOne) |

Ces modules peuvent être utilisés pour des évolutions futures du script (export CSV, support PostgreSQL, etc.).