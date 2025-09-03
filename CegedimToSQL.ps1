# CegedimToSQL.ps1

# --------------------------------------------------------
#               Fonction de traçage pour $script:CI
# --------------------------------------------------------
function Trace-CI {
    param([string]$location, [string]$action = "CHECK")
    $caller = (Get-PSCallStack)[1].Command
    $line = (Get-PSCallStack)[1].ScriptLineNumber
    DBG "Trace-CI" "TRACE-CI [$caller`:$line] $location - $action - Taille: $($script:CI.Count)" -ForegroundColor Magenta
    if ($script:CI.Count -gt 0) {
        $firstKey = $script:CI.Keys | Select-Object -First 1
        DBG "Trace-CI" "  Premier RecId: $firstKey" -ForegroundColor Magenta
    }
}

# --------------------------------------------------------
#               Traitement des LOGS
# --------------------------------------------------------
function Level {
    Param ( [string]$func )
    
    $n = (Get-PSCallStack).Count-3
    $s = ' ' * $n +  $func
    $s = "{0,-36}" -f $s
    return $s
} # code   param -func                                                        Return : [string]    -> Formate la chaine (func)
function Var_Add {
	param ( [string]$file, [string]$value)

	if ( $file -eq $script:cfg["intf"]["pathfilelog"] )     { $script:pathfilelog     += $value }
	if ( $file -eq $script:cfg["intf"]["pathfileerr"] )     { $script:pathfileerr     += $value }
	if ( $file -eq $script:cfg["intf"]["pathfilemod"] )     { $script:pathfilemod     += $value }
}
function Save_logs {
	$script:pathfilelog | Add-Content -Path $script:cfg["intf"]["pathfilelog"]
	$script:pathfileerr | Add-Content -Path $script:cfg["intf"]["pathfileerr"]
	$script:pathfilemod | Add-Content -Path $script:cfg["intf"]["pathfilemod"]
}
function OUT {
	Param ( [string]$trigramme, [string]$func, [string]$msg, [string]$color="White",[bool]$CRLF=$false, [bool]$EMAIL=$false, [bool]$NOSCREEN=$false, [switch]$DBG, [switch]$LOG, [switch]$MOD, [switch]$DLT, [switch]$INA, [switch]$ERR, [switch]$ADDERR, [switch]$ADDWRN)

	# Chaine a afficher
	$f = Level $func
	$Stamp = (Get-Date).toString("yyyy-MM-dd HH:mm:ss")
	$str = "$trigramme : $f : $msg"
	$stampstr = "$Stamp : $str"

	# Affichage a l'ecran

	if ( $script:cfg["start"]["logtoscreen"] -eq "yes" -and -not $NOSCREEN ) { 
		if ( $CRLF ) { Write-Host "" }
		try {
			Write-Host $str -ForegroundColor $color 
		} catch {
			Write-Host $str -ForegroundColor Green
		}
	}

	# Ajout dans les fichiers de logs
	if ( $CRLF ) {
		if ( $DBG -and $script:cfg["start"]["debug"] -eq "yes" ) { Var_Add $($script:cfg["intf"]["pathfilelog"]) -value $stampstr }
		if ( $LOG ) { Var_Add $($script:cfg["intf"]["pathfilelog"])     -value "" }
		if ( $ERR ) { Var_Add $($script:cfg["intf"]["pathfileerr"])     -value "" }
		if ( $MOD ) { Var_Add $($script:cfg["intf"]["pathfilemod"])     -value "" }

	}
	if ( $DBG -and $script:cfg["start"]["debug"] -eq "yes" ) { Var_Add $($script:cfg["intf"]["pathfilelog"]) -value $stampstr }
	if ( $LOG ) { Var_Add $($script:cfg["intf"]["pathfilelog"])     -value $stampstr }
	if ( $ERR ) { Var_Add $($script:cfg["intf"]["pathfileerr"])     -value $stampstr }
	if ( $MOD ) { Var_Add $($script:cfg["intf"]["pathfilemod"])     -value $stampstr }
	
	if ( $ADDERR ) { $script:ERREUR  += 1 }
	if ( $ADDWRN ) { $script:WARNING += 1 }

	# Ajoute a Email
	if ( $EMAIL ) { 
		if ( $CRLF ) { $script:emailtxt.Add("") }
		$script:emailtxt.Add($stampstr)	
	}
}
function DBG {
	Param ( [string]$func, [string]$msg, [switch]$CRLF )
	if ( $script:cfg["start"]["debug"] -eq "yes" ) { 
		OUT "DBG" $func $msg "Gray" -DBG -LOG -CRLF $CRLF
	}
} # code   param -func, -msg                                                  Return : N/A         -> Ecrit DBG (func) (msg) dans LOG si [start][debugtolog], et a l'ecran [start][debug] = yes
function LOG {
	Param ( [string]$func, [string]$msg, [string]$color = "Cyan", [switch]$CRLF, [switch]$EMAIL)
	OUT "LOG" $func $msg $color -LOG -CRLF $CRLF -EMAIL $EMAIL
} # code   param -func, -msg, -color, -CRLF, -EMAIL                           Return : N/A         -> Ecrit LOG (func) (msg) dans LOG, et a l'ecran [start][debug] = yes couleur (color)
function ERR {
	Param ( [string]$func, [string]$msg, [switch]$CRLF )

	if ( $script:ERREUR -eq 0 ) {
		OUT "ERR" $func $entete "Red" -ERR -CRLF $CRLF -NOSCREEN $true -EMAIL $true
	}
	OUT "ERR" $func $msg "Red" -ERR -LOG -CRLF $CRLF -ADDERR -EMAIL $true
} # code   param -func, -msg, -CRLF                                           Return : N/A         -> Ecrit ERR (func) (msg) dans ERR, et a l'ecran [start][debug] = yes
function WRN {
	Param ( [string]$func, [string]$msg, [switch]$CRLF )

	if ( $script:cfg["start"]["warntoerr"] -eq "yes" ) {
		OUT "WRN" $func $msg "Magenta" -LOG -CRLF $CRLF -ERR -EMAIL $true -ADDWRN
	} else {
		OUT "WRN" $func $msg "Magenta" -LOG -CRLF $CRLF -EMAIL $true -ADDWRN
	}
} # code   param -func, -msg, -CRLF                                           Return : N/A         -> Ecrit WRN (func) (msg) dans et sort du script
function QUIT {
    Param ( [string]$func, [string]$msg )

    $s = "Duree d'execution : {0:N1} secondes" -f $script:start.Elapsed.TotalSeconds

    if ( $script:ERREUR -eq 0 ) { $c = "Green" } else { $c = "Red" }
	LOG "QUIT" "$($script:ERREUR) erreur, $($script:WARNING) warning, $s" $c -EMAIL
    #$script:emailtxt.Add("$Stamp : QUIT : $f : $msg")
    if ( $script:ERREUR -ne 0 ) {
		OUT "END" $func $msg $c -LOG

        # Contexte
        Get-PSCallStack | Where-Object { $_.Command -and $_.Location } | ForEach-Object {
            if ($_.Command -ne "QUIT") { 
				OUT "END" "$($_.Command)" "$($_.Location)" "Gray" -DBG -LOG -EMAIL $true
            } 
        }
    }

	# Sujet de l'email
	if ( -not $script:MailErr ) {
		$subject = (
			"$($script:cfg['email']['Subject']) : " +
			"[$($script:ERREUR) Erreurs], " +
			"[$($script:WARNING) Warnings]"
		)
		SendEmail $subject $script:emailtxt 
	}

	Save_logs
    exit 0
} # code   param -func, -msg                                                  Return : N/A         -> Ecrit ERR (func) (msg) dans et sort du script 
function QUITEX {
    Param ( [string]$func, [string]$msg, [switch]$ADDERR )

    $lines = $msg -split "`n"
    foreach ($line in $lines) { 
        if ($ADDERR) { 
            ERR "$func" "$line" }
        else { LOG "$func" "$line" }
    }
    QUIT "$func" "Script interrompu."
} # code   param -func -msg, -ADDERR                                          Return : N/A         -> Ecrit ERR (func) (msg) dans et sort du script
function MOD {
	Param ( [string]$func, [string]$msg, [switch]$CRLF )

	if ( $script:cfg["start"]["ApplyUpdate"] -eq "no" ) { $mod = "SIM" } else { $mod = "MOD" }
	OUT $mod $func $msg "Yellow" -LOG -MOD -CRLF $CRLF -EMAIL $true
} # code   param -func, -msg, -CRLF                                           Return : N/A         -> Ecrit MOD (func) (msg) dans LOG, et a l'ecran [start][debug] = yes, saut de ligne si switch -CRLF

# --------------------------------------------------------
#               Chargement fichier .ini
# --------------------------------------------------------

function LoadIni {
	# initialisation variables liste des logs
	$script:pathfilelog = @()
	$script:pathfileerr = @()
	$script:pathfileina = @()
	$script:pathfiledlt = @()
	$script:pathfilemod = @()
	
	# sections de base du fichier .ini
	$script:cfg = @{
        "start"                   = @{}
        "intf"                    = @{}
        "email"                   = @{}
        "URL"                     = @{}
    }
    # Recuperation des parametres passes au script 
    $script:execok  = $false

    if (-not(Test-Path $($script:cfgFile) -PathType Leaf)) { Write-Host "Fichier de parametrage $script:cfgFile innexistant"; exit 1 }
    Write-Host "Fichier de parametrage $script:cfgFile"

    # Initialisation des sections parametres.
    $script:start    = [System.Diagnostics.Stopwatch]::startNew()
    $script:MailErr  = $false
    $script:WARNING  = 0
    $script:ERREUR   = 0
	
	$script:emailtxt = New-Object 'System.Collections.Generic.List[string]'

	$script:cfg = Add-IniFiles $script:cfg $script:cfgFile

	# Recherche des chemins de tous les fichiers et verification de leur existence
	if (-not ($script:cfg["intf"].ContainsKey("rootpath")) ) {
		$script:cfg["intf"]["rootpath"] = $PSScriptRoot
	}
	$script:cfg["intf"]["pathfilelog"] 	= GetFilePath $script:cfg["intf"]["pathfilelog"]
	$script:cfg["intf"]["pathfileerr"]	= GetFilePath $script:cfg["intf"]["pathfileerr"]
	$script:cfg["intf"]["pathfilemod"]  = GetFilePath $script:cfg["intf"]["pathfilemod"]

	# Suppression des fichiers One_Shot
	if ((Test-Path $($script:cfg["intf"]["pathfilelog"]) -PathType Leaf)) { Remove-Item -Path $script:cfg["intf"]["pathfilelog"]}    

	# Création des fichiers innexistants
	$null = New-Item -type file $($script:cfg["intf"]["pathfilelog"]) -Force;
	if (-not(Test-Path $($script:cfg["intf"]["pathfileerr"]) -PathType Leaf)) { $null = New-Item -type file $($script:cfg["intf"]["pathfileerr"]) -Force; }
	if (-not(Test-Path $($script:cfg["intf"]["pathfilemod"]) -PathType Leaf)) { $null = New-Item -type file $($script:cfg["intf"]["pathfilemod"]) -Force; }

    $script:token = Encode $script:cfg["Cegedim"]["Token"]

	# Initialisation de la hashtable pour stocker les données CI
	$script:CI = @{}
	Trace-CI "LoadIni" "INIT"
}
function GetFilePath {
	param ( [string]$pattern, [switch]$Needed )

	# Remplacement de la chaîne $rootpath$ par le contenu de $script:cfg["intf"]["rootpath"]
	if ($pattern -match '\$rootpath\$') {
		$pattern = $pattern -replace '\$rootpath\$', $script:cfg["intf"]["rootpath"]
	}

	$folder = Split-Path $pattern -Parent
	$filter = Split-Path $pattern -Leaf

	# Créer le répertoire s'il n'existe pas
	if (-not (Test-Path -Path $folder -PathType Container)) {
		try {
			New-Item -Path $folder -ItemType Directory -Force | Out-Null
			DBG "GetFilePath" "Répertoire créé : $folder"
		}
		catch {
			QUITEX "GetFilePath" "Impossible de créer le répertoire '$folder' : $($_.Exception.Message)"
		}
	}

	$files = Get-ChildItem -Path $folder -Filter $filter -File

	if ($files.Count -eq 1) {
		$filepath = $files[0].FullName
	} elseif ($files.Count -eq 0) {
		if ($Needed) {
			QUITEX "GetFilePath" "Aucun fichier ne correspond au filtre '$filter' dans '$folder'" -ADDERR
		} else {
			WRN "GetFilePath" "Aucun fichier ne correspond au filtre '$filter' dans '$folder'"
			$filepath = $pattern
		}
	} else {
		QUITEX "GetFilePath" "Plusieurs fichiers correspondent au filtre '$filter' dans '$folder'" -ADDERR
	}
	return $filepath
}
function Query_type {
    param ($url, [switch]$Debug)

    $ignore    = @('A_CommandeLink','AssetLink','DefaultSLPLink','DeviceSubnet','DeviceLocationLink','DSMPlatformLink','InventorySettingsLink','LocationLink','ManufacturerLink','OrgUnitLink','pPage','ServiceProviderLink')
    $datefield = @('A_DateLivraison','A_DateInventaireManuel','LastModDateTime','BIOSDate','LastReportDateTime','LastAuditDateTime','CreatedDateTime','LastChangeDateTime','LastScanDateTime','NextReviewDate','LastComplianceCheck','PushNotificationRegDateTime','Last_login_date','token_end_date','token_start_date','A_DateFinGarantie','A_last_logon')
    $decimal_2 = @('TotalMemory','CPUSpeed','CIVersion')
    $decimal_4 = @('TargetAvailability')

    LOG "Query_type" "$url" -CRLF
    Trace-CI "Query_type-DEBUT" "CHECK"
    
    # Vérification que $script:CI existe et est bien une hashtable
    if ($null -eq $script:CI) {
        DBG "Query_type" "ERREUR: `$script:CI n'est pas initialisé !"
        $script:CI = @{}
        Trace-CI "Query_type-REINIT" "REINIT"
    }

    $handler = $null
    $client = $null
    
    try {
        $handler = New-Object System.Net.Http.HttpClientHandler
        $client = [System.Net.Http.HttpClient]::new($handler)

        $originalUrl = $url
        $pageCount = 0
        $pageSize = 25  # Taille de page fixe de l'API Ivanti/HEAT
        $skip = 0
        $totalCount = $null
        $recordsProcessed = 0
        $maxPages = 1000  # Limite de sécurité pour éviter les boucles infinies

        while ($true) {
            $pageCount++
            
            # Vérification de la limite de sécurité
            if ($pageCount -gt $maxPages) {
                WRN "Query_type" "ALERTE: Limite de sécurité atteinte ($maxPages pages) - Arrêt de la pagination" -ForegroundColor Red
                break
            }
            
            # Diagnostic : vérifier la taille de $script:CI au début de chaque page
            Trace-CI "Page-$pageCount-DEBUT" "CHECK"
            
            # Construction de l'URL avec pagination manuelle et tri explicite
            $separator = if ($url.Contains("?")) { "&" } else { "?" }
            $currentUrl = "$url${separator}`$top=$pageSize&`$skip=$skip&`$orderby=RecId"
            
            DBG "Query_type" "  URL appelée: $currentUrl" -ForegroundColor Gray
            
            $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $currentUrl)
            $request.Headers.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new("application/json"))

            $authValue = "rest_api_key=$script:token"
            $request.Headers.TryAddWithoutValidation("Authorization", $authValue) | Out-Null

            $response = $client.SendAsync($request).Result
            
            if (-not $response.IsSuccessStatusCode) {
                throw "HTTP Error: $($response.StatusCode) - $($response.ReasonPhrase)"
            }
            
            $content = $response.Content.ReadAsStringAsync().Result
            $data = $content | ConvertFrom-Json

            # Récupération du total à la première page
            if ($pageCount -eq 1 -and $data.'@odata.count') {
                $totalCount = $data.'@odata.count'
                DBG "Query_type" "Total disponible selon l'API: $totalCount"
            }

            # Forcer $data.value à être un tableau pour éviter les problèmes de comptage
            $currentPageData = @($data.value)
            $currentPageCount = $currentPageData.Count

            # Si aucune donnée reçue, on arrête
            if ($currentPageCount -eq 0) {
                DBG "Query_type" "Aucune donnée reçue - Arrêt de la pagination"
                break
            }

            # Vérifier si on va dépasser le total avant de traiter les données
            if ($totalCount -and ($recordsProcessed + $currentPageCount) -gt $totalCount) {
                # Ne traiter que le nombre d'éléments nécessaires pour atteindre le total
                $recordsToProcess = $totalCount - $recordsProcessed
                if ($recordsToProcess -le 0) {
                    DBG "Query_type" "Total déjà atteint - Arrêt de la pagination"
                    break
                }
                $currentPageData = $currentPageData[0..($recordsToProcess - 1)]
                DBG "Query_type" "Limitation à $recordsToProcess enregistrements pour ne pas dépasser le total de $totalCount"
            }

            # Traitement des données : stockage dans $script:CI
            #Trace-CI "Page-$pageCount-AVANT-TRAITEMENT" "CHECK"
            $recordsInThisPage = 0
            $ciSizeBefore = $script:CI.Count
            
            # Diagnostic : afficher les premiers RecId de cette page
            $pageRecIds = @()
            foreach ($record in $currentPageData) {
                if ($record.RecId) {
                    $pageRecIds += $record.RecId
                }
            }
            DBG "Query_type" "  RecIds de cette page (premiers 5): $($pageRecIds[0..4] -join ', ')"
            
            # Diagnostic : vérifier si ces RecIds existent déjà dans $script:CI
            $existingCount = 0
            $newCount = 0
            $duplicateRecIds = @()
            foreach ($recId in $pageRecIds) {
                if ($script:CI.ContainsKey($recId)) {
                    $existingCount++
                    $duplicateRecIds += $recId
                } else {
                    $newCount++
                }
            }
            DBG "Query_type" "  RecIds existants: $existingCount, Nouveaux RecIds: $newCount"
            
            # Alerte si tous les RecIds sont des doublons (problème de pagination)
            if ($existingCount -gt 0 -and $newCount -eq 0 -and $pageCount -gt 1) {
                DBG  "Query_type" "  ALERTE: Tous les RecIds de cette page sont des doublons!" -ForegroundColor Red
                DBG "Query_type" "  RecIds dupliqués: $($duplicateRecIds[0..4] -join ', ')" -ForegroundColor Red
                DBG "Query_type" "  Cela indique un problème de pagination - l'API retourne les mêmes données" -ForegroundColor Red
                DBG "Query_type" "  Arrêt de la pagination pour éviter une boucle infinie" -ForegroundColor Red
                #break
            }
            
            foreach ($record in $currentPageData) {
                if ($record.RecId) {
                    # Diagnostic : vérifier si ce RecId existe déjà
                    $existsBefore = $script:CI.ContainsKey($record.RecId)
                    
                    # Créer une hashtable pour ce RecID s'il n'existe pas déjà
                    if (-not $script:CI.ContainsKey($record.RecId)) {
                        $script:CI[$record.RecId] = @{}
                        if ($recordsInThisPage -lt 3) {  # Limiter les messages pour les 3 premiers
                            DBG "Query_type" "  Nouveau RecID ajouté: $($record.RecId)"
                            Trace-CI "AJOUT-$($record.RecId)" "ADD"
                        }
                    } else {
                        if ($recordsInThisPage -lt 3) {  # Limiter les messages pour les 3 premiers
                            WRN "Query_type" "  RecID existant mis à jour: $($record.RecId)"
                        }
                    }
                    
                    # Stocker tous les champs de l'enregistrement (sauf ceux de la liste $ignore)
                    foreach ($property in $record.PSObject.Properties) {
                        if ($property.Name -notin $ignore) {
                            $value = $property.Value
                            
                            # Conversion des champs datetime selon la liste $datefield
                            if ($property.Name -in $datefield -and $value -ne $null -and $value -ne "") {
                                try {
                                    # Utiliser la fonction ConvertDateToString pour convertir au format spécifié
                                    $convertedValue = ConvertDateToString -value $value -formatOut $script:cfg["SQL_Server"]["frmtdateOUT"]
                                    if ($convertedValue -ne $value) {
                                        $value = $convertedValue
                                        if ($recordsInThisPage -le 3) {  # Limiter les messages pour les 3 premiers
                                            DBG "Query_type" "    Champ date converti: $($property.Name) = $value"
                                        }
                                    }
                                } catch {
                                    # En cas d'erreur de conversion, garder la valeur originale et logger
                                    if ($recordsInThisPage -le 3) {
                                        WRN "Query_type" "    Erreur conversion date pour $($property.Name): $($_.Exception.Message)"
                                    }
                                }
                            }
                            
                            # Conversion des champs numériques selon la liste $decimal_2
                            if ($property.Name -in $decimal_2 -and $value -ne $null -and $value -ne "") {
                                try {
                                    # Convertir en nombre décimal avec 2 chiffres après la virgule
                                    $numericValue = [decimal]$value
                                    $value = "{0:F2}" -f $numericValue
                                } catch {}
                            }
                            # Conversion des champs numériques selon la liste $decimal_4
                            if ($property.Name -in $decimal_4 -and $value -ne $null ) {
                                $vrec = $record.RecId
                                try {
                                    # Convertir en nombre décimal avec 2 chiffres après la virgule
                                    $numericValue = [decimal]$value
                                    $value = "{0:F4}" -f $numericValue
                                } catch {}
                            }
                            
                            $script:CI[$record.RecId][$property.Name] = $value
                        } else {
                            # Log optionnel pour debug : champ ignoré
                            if ($recordsInThisPage -le 3) {  # Limiter les messages pour les 3 premiers
                                DBG "Query_type" "    Champ ignoré: $($property.Name)"
                            }
                        }
                    }
                    
                    $recordsInThisPage++
                    
                    # Diagnostic : vérifier la taille après chaque ajout (pour les 3 premiers)
                    if ($recordsInThisPage -le 3) {
                        DBG "Query_type" "    Après ajout RecId $($record.RecId): `$script:CI.Count = $($script:CI.Count)"
                    }
                }
            }
            
            $ciSizeAfter = $script:CI.Count
            DBG "Query_type" "Page-$pageCount-APRES-TRAITEMENT" "CHECK"
            DBG "Query_type" "  Taille `$script:CI avant traitement: $ciSizeBefore"
            DBG "Query_type" "  Taille `$script:CI après traitement: $ciSizeAfter"
            DBG "Query_type" "  Différence: $($ciSizeAfter - $ciSizeBefore)"
            
            $recordsProcessed += $recordsInThisPage

            # Debug activé par défaut pour diagnostiquer le problème de pagination
            DBG "Query_type" "Page $pageCount - URL: $currentUrl"
            DBG "Query_type" "Données reçues: $currentPageCount enregistrements"
            DBG "Query_type" "Enregistrements avec RecId traités: $recordsInThisPage"
            DBG "Query_type" "Total récupéré jusqu'à présent: $recordsProcessed"
            DBG "Query_type" "Total dans `$script:CI: $($script:CI.Count)"
            if ($totalCount) {
                DBG "Query_type" "Total disponible: $totalCount"
            }
            DBG "Query_type" "Skip actuel: $skip, Prochain skip: $($skip + $pageSize)"
            DBG "Query_type" "---"
            
            # Arrêt si on a atteint le total exact
            if ($totalCount -and $recordsProcessed -ge $totalCount) {
                DBG "Query_type" "Total atteint ($recordsProcessed >= $totalCount) - Arrêt de la pagination"
                break
            }
            
            # Arrêt si on a reçu moins d'éléments que la taille de page (dernière page)
            if ($currentPageCount -lt $pageSize) {
                DBG "Query_type" "Dernière page détectée ($currentPageCount < $pageSize) - Arrêt de la pagination"
                break
            }
            
            # Incrémenter skip avec la taille de page fixe, pas le nombre d'éléments reçus
            $skip += $pageSize
            
            # Diagnostic : vérifier si $script:CI est stable entre les pages
            Trace-CI "Page-$pageCount-FIN" "CHECK"
        }

        LOG "Query_type"  "$recordsProcessed enregistrements ($pageCount pages)"
        LOG "Query_type"  "Taille finale de `$script:CI: $($script:CI.Count) enregistrements"
        
        # Diagnostic : afficher quelques exemples de RecId stockés
        $sampleRecIds = $script:CI.Keys | Select-Object -First 5
        DBG "Query_type"  "Exemples de RecId dans `$script:CI:"
        foreach ($recId in $sampleRecIds) {
            $name = if ($script:CI[$recId]['Name']) { $script:CI[$recId]['Name'] } else { "N/A" }
            DBG "Query_type"  "  RecId: $recId - Name: $name"
        }
        
        Trace-CI "Query_type-FIN" "CHECK"
    }
    catch {
        ERR "Query_type" "Erreur lors de l'appel à $originalUrl : $($_.Exception.Message)"
    }
    finally {
        if ($client) { $client.Dispose() }
        if ($handler) { $handler.Dispose() }
    }
}

function Query_BDD_CI {
	$script:BDDCI = @{}
	Query_BDDTable -tableName $script:cfg["SQL_Server"]["table"] -functionName "Query_BDD_CI" -keyColumns @("RecID") -targetVariable $script:BDDCI -UseFrmtDateOUT
}

function Update_BDD_CI {
	Update_BDDTable  $script:CI $script:BDDCI  @("RecID") $script:cfg["SQL_Server"]["table"] "Update_BDD_CI" { Query_BDD_CI }
}


function Get-BDDConnectionParams {
    return @{
        server      = $script:cfg["SQL_Server"]["server"]
        database    = $script:cfg["SQL_Server"]["database"]
        login       = $script:cfg["SQL_Server"]["login"]
        password    = Encode $script:cfg["SQL_Server"]["password"]
        datefrmtout = $script:cfg["SQL_Server"]["frmtdateOUT"]
    }
}

# Fonction utilitaire pour effectuer une requête BDD standard
function Query_BDDTable {
    param(
        [string]$tableName,
        [string]$functionName,
        [array]$keyColumns,
        [hashtable]$targetVariable,
        [switch]$UseFrmtDateOUT
    )
    
    $params = Get-BDDConnectionParams
    
    LOG $functionName "Chargement de la table [$tableName] en memoire" -CRLF
    
    # Vider la hashtable cible
    $targetVariable.Clear()
    
    # Paramètres pour QueryTable
    $queryParams = @{
        server = $params.server
        database = $params.database
        table = $tableName
        login = $params.login
        password = $params.password
        keycolumns = $keyColumns
    }
    
    # Ajouter le format de date si demandé
    if ($UseFrmtDateOUT) {
        $queryParams.frmtdateOUT = $script:cfg["SQL_Server"]["frmtdateOUT"]
    }
    
    # Exécuter la requête et affecter le résultat
    $result = QueryTable @queryParams
    
    # Copier le résultat dans la variable cible
    foreach ($key in $result.Keys) {
        $targetVariable[$key] = $result[$key]
    }
}
# Fonction utilitaire pour effectuer une mise à jour BDD standard
function Update_BDDTable {
    param(
        [hashtable]$sourceData,
        [hashtable]$targetData,
        [array]$keyColumns,
        [string]$tableName,
        [string]$functionName,
        [scriptblock]$reloadFunction
    )
    
    $params = Get-BDDConnectionParams
    
    LOG $functionName "Update de la table $tableName" -CRLF
    
    UpdateTable $sourceData $targetData $keyColumns $params.server $params.database $tableName $params.login $params.password $script:cfg["start"]["ApplyUpdate"]
    
    # Recharger les modifs en memoire
    if ($reloadFunction) {
        & $reloadFunction
    }
}
# --------------------------------------------------------
#               Main
# --------------------------------------------------------

# Initialisation Culture pour encodage UTF8 et separator numerique "."
chcp 65001 > $null # Encodage avec accent
# Cloner la culture actuelle
$culture = [System.Globalization.CultureInfo]::CurrentCulture.Clone()
# Modifier uniquement le séparateur décimal (de ',' à '.')
$culture.NumberFormat.NumberDecimalSeparator = '.'
# Appliquer cette culture modifiée à la session en cours
[System.Threading.Thread]::CurrentThread.CurrentCulture = $culture

$script:cfgFile = "$PSScriptRoot\CegedimToSQL.ini"
. "$PSScriptRoot\Modules\Ini.ps1" > $null 

LoadIni

Add-Type -AssemblyName System.Web

# Chargement des modules
. "$PSScriptRoot\Modules\Encode.ps1"     > $null 
. "$PSScriptRoot\Modules\SendEmail.ps1"  > $null 
. "$PSScriptRoot\Modules\StrConvert.ps1" > $null 
if ($script:cfg["start"]["TransacSQL"] -eq "AllInOne" ) {
	. "$PSScriptRoot\Modules\SQLServer - TransactionAllInOne.ps1" > $null
} else {
	. "$PSScriptRoot\Modules\SQLServer - TransactionOneByOne.ps1" > $null
}


Query_BDD_CI

LOG "MAIN" "Récupération des équipements Ivanti/HEAT..." -CRLF

# Boucle sur toutes les URLs définies dans la section [URL] du fichier .ini
foreach ($urlKey in $script:cfg["URL"].Keys) {
    $url = $script:cfg["URL"][$urlKey]
    Query_type $url
}


Update_BDD_CI

QUIT "MAIN" "Process terminé"


