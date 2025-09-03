function QueryTable {
    param (
        [string]$server,
        [string]$database,
        [string]$table,
        [string]$login,
        [string]$password,
        [string[]]$keycolumns,
        [string]$frmtdateOUT = $null
    )

    DBG "QueryTable" "Chargement table [$table] avec clé(s): $($keycolumns -join ', ')"
    if ($frmtdateOUT) {
        DBG "QueryTable" "Conversion des dates au format: $frmtdateOUT"
    }

    $connectionString = "Server=$server;Database=$database;User Id=$login;Password=$password;TrustServerCertificate=True;Encrypt=True"
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)

    try {
        $connection.Open()
    } catch {
        QUITEX "QueryTable" "Connexion à la base '$database' sur $server impossible" -ADDERR
    }

    $command = $connection.CreateCommand()
    $command.CommandText = "SELECT * FROM [$table]"

    try {
        $reader = $command.ExecuteReader()
    } catch {
        QUITEX "QueryTable" "Lecture de la table '$table' échouée sur $server" -ADDERR
    }

    $hash = @{}
    $cpt = 0

    while ($reader.Read()) {
        $row = @{}
        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
            $field = $reader.GetName($i)
            $value = $reader.GetValue($i)
            
            # Conversion des dates au format spécifié si nécessaire
            if ($frmtdateOUT -and $value -is [DateTime]) {
                $value = $value.ToString($frmtdateOUT)
            }
            
            $row[$field] = $value
        }

        # Construction de la clé hiérarchique
        $current = $hash
        for ($k = 0; $k -lt ($keycolumns.Count - 1); $k++) {
            $keyname = $keycolumns[$k]
            if (-not $row.ContainsKey($keyname)) {
                QUITEX "QueryTable" "Clé manquante '$keyname' dans la ligne (colonnes: $($row.Keys -join ', '))"
            }
            
            # Gestion spéciale pour les dates
            $kval = $row[$keyname]
            if ($kval -is [DateTime]) {
                if ($frmtdateOUT) {
                    $kval = $kval.ToString($frmtdateOUT)
                } else {
                    $kval = $kval.ToString("dd/MM/yyyy")
                }
            } else {
                $kval = [string]$kval
            }
            
            if (-not $current.ContainsKey($kval)) {
                $current[$kval] = @{}
            }
            $current = $current[$kval]
        }

        $lastkey = $keycolumns[-1]
        if (-not $row.ContainsKey($lastkey)) {
            QUITEX "QueryTable" "Clé manquante '$lastkey' dans la ligne (colonnes: $($row.Keys -join ', '))"
        }

        # Gestion spéciale pour les dates
        $lastval = $row[$lastkey]
        if ($lastval -is [DateTime]) {
            if ($frmtdateOUT) {
                $lastval = $lastval.ToString($frmtdateOUT)
            } else {
                $lastval = $lastval.ToString("dd/MM/yyyy")
            }
        } else {
            $lastval = [string]$lastval
        }
        
        $current[$lastval] = $row
        $cpt++
    }

    $reader.Close()
    $connection.Close()

    LOG "QueryTable" "$cpt enregistrements chargés depuis la table '$table' ( $($hash.Count) clés [$($keycolumns[0])] )"
    return $hash
}

function Normalize {
    param ([string]$str)

	#$shortfrmtin = $frmtin.Substring(1, [Math]::Min(10, $frmtin.Length - 1))
	#$shortfrmtout = $frmtout.Substring(1, [Math]::Min(10, $frmtout.Length - 1))

	if ([string]::IsNullOrWhiteSpace($str)) {
		return "NULL"
	} elseif ($str -eq "True") {
		return "1"
	} elseif ($str -eq "False") {
		return "0"
    }
	$str = $str.Trim()
	return "N'$($str -replace "'", "''")'" 
}

function IsDifferent {
    param (
        [string]$src,
        [string]$dst
    )

    $dtSrc = $null
    $dtDst = $null
    $isSrcDate = $false
    $isDstDate = $false

    # Sinon, comparer les chaînes brutes
    if ($src -ne $dst) {
        # DBG "IsDifferent" "Valeurs différentes : $src <> $dst"
        return $true
    } 
    return $false
}

function ExecuteQuery {
    param (
        [string]$query,
        [string]$server,
        [string]$database,
        [string]$login,
        [string]$password,
        [string]$operationType
    )

    $connectionString = "Server=$server;Database=$database;User Id=$login;Password=$password;TrustServerCertificate=True;Encrypt=True"
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)

    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteNonQuery() | Out-Null
        $connection.Close()
        return $true
    } catch {
        if ($connection.State -eq 'Open') {
            $connection.Close()
        }
        ERR "ExecuteQuery" "Échec de l'exécution de la requête $operationType : $($_.Exception.Message)"
        ERR "ExecuteQuery" "Requête : $query"
        return $false
    }
}

function UpdateTable {
    param (
        $hsrc, $hdst, [string[]]$keycolumns,
        $server, $database, $table,
        $login, $password, $apply,
        [bool]$allowDelete = $false
    )

    if ($allowDelete) {
        LOG "UpdateTable" "Synchronisation de la table '$table' (avec suppression des enregistrements orphelins)"
    } else {
        LOG "UpdateTable" "Synchronisation de la table '$table' (sans suppression)"
    }

    # Compteurs pour les statistiques
    $updateCount = 0
    $insertCount = 0
    $deleteCount = 0
    $updateSuccess = 0
    $insertSuccess = 0
    $deleteSuccess = 0

    function TraverseHash {
        param ($nodeSrc, $nodeDst, $depth, $keyPath)

        $datefrmtBDD = $cfg["SQL_Server"]["frmtdateOUT"]
        
        foreach ($k in $nodeSrc.Keys) {
            $newKeyPath = $keyPath + @($k)

            if ($depth -lt ($keycolumns.Count - 1)) {
                $subDst = if ($nodeDst.ContainsKey("$($k)")) { $nodeDst["$($k)"] } else { @{} }
                TraverseHash $nodeSrc[$k] $subDst ($depth + 1) $newKeyPath
            }
            else {
                $srcRow = $nodeSrc[$k]
                $dstRow = if ($nodeDst.ContainsKey("$($k)")) { $nodeDst["$($k)"] } else { $null }

                if ($dstRow) {
                    # Mise à jour
                    $diff = @{}
                    foreach ($col in $srcRow.Keys) {
                        $srcVal = if ($srcRow["$($col)"] -ne $null) { $srcRow["$($col)"].ToString().Trim() } else { "" }
                        $dstVal = if ($dstRow["$($col)"] -ne $null) { $dstRow["$($col)"].ToString().Trim() } else { "" }

                        if (IsDifferent $srcVal $dstVal) {
                            $diff["$($col)"] = $srcRow["$($col)"]
                        }
                    }

                    if ($diff.Count -gt 0) {
                        $script:updateCount++
                        
                        $setClause = ($diff.GetEnumerator() | ForEach-Object {
                            $val = Normalize $_.Value
                            "[$($_.Key)] = $val"
                        }) -join ", "

                        $whereClause = ""
                        for ($i = 0; $i -lt $keycolumns.Count; $i++) {
                            $col = $keycolumns[$i]
                            $val = $newKeyPath[$i] -replace "'", "''"
                            $whereClause += "[$col] = N'$val'"
                            if ($i -lt $keycolumns.Count - 1) {
                                $whereClause += " AND "
                            }
                        }

                        $query = "UPDATE [$table] SET $setClause WHERE $whereClause;"
                        
                        if ($apply -eq "yes") {
                            if (ExecuteQuery $query $server $database $login $password "UPDATE") {
                                $script:updateSuccess++
                                MOD "UpdateTable" "SQL UPDATE : $query"
                            }
                        } else {
                            MOD "UpdateTable" "SQL UPDATE (non exécutée) : $query"
                        }
                    }

                } else {
                    # Insertion
                    $script:insertCount++
                    
                    $columns = $srcRow.Keys
                    $values = foreach ($col in $columns) {
                        Normalize $srcRow[$col]
                    }

                    $colList = ($columns | ForEach-Object { "[$_]" }) -join ", "
                    $valList = $values -join ", "
                    $query = "INSERT INTO [$table] ($colList) VALUES ($valList);"
                    
                    if ($apply -eq "yes") {
                        if (ExecuteQuery $query $server $database $login $password "INSERT") {
                            $script:insertSuccess++
                            MOD "UpdateTable" "SQL INSERT : $query"
                        }
                    } else {
                        MOD "UpdateTable" "SQL INSERT (non exécutée) : $query"
                    }
                }
            }
        }
    }

    function TraverseHashForDeletion {
        param ($nodeSrc, $nodeDst, $depth, $keyPath)
        
        foreach ($k in $nodeDst.Keys) {
            $newKeyPath = $keyPath + @($k)

            if ($depth -lt ($keycolumns.Count - 1)) {
                $subSrc = if ($nodeSrc.ContainsKey("$($k)")) { $nodeSrc["$($k)"] } else { @{} }
                TraverseHashForDeletion $subSrc $nodeDst[$k] ($depth + 1) $newKeyPath
            }
            else {
                # Vérifier si cet enregistrement existe dans la source
                if (-not $nodeSrc.ContainsKey("$($k)")) {
                    # Cet enregistrement existe dans la destination mais pas dans la source -> à supprimer
                    $script:deleteCount++
                    
                    $whereClause = ""
                    for ($i = 0; $i -lt $keycolumns.Count; $i++) {
                        $col = $keycolumns[$i]
                        $val = $newKeyPath[$i] -replace "'", "''"
                        $whereClause += "[$col] = N'$val'"
                        if ($i -lt $keycolumns.Count - 1) {
                            $whereClause += " AND "
                        }
                    }

                    $query = "DELETE FROM [$table] WHERE $whereClause;"
                    
                    if ($apply -eq "yes") {
                        if (ExecuteQuery $query $server $database $login $password "DELETE") {
                            $script:deleteSuccess++
                            MOD "UpdateTable" "SQL DELETE : $query"
                        }
                    } else {
                        MOD "UpdateTable" "SQL DELETE (non exécutée) : $query"
                    }
                }
            }
        }
    }

    # Traitement des mises à jour et insertions
    TraverseHash $hsrc $hdst 0 @()
    
    # Traitement des suppressions uniquement si autorisé
    if ($allowDelete) {
        TraverseHashForDeletion $hsrc $hdst 0 @()
    }

    $totalOperations = $updateCount + $insertCount + $deleteCount

    if ($totalOperations -gt 0) {
        if ($apply -eq "yes") {
            if ($allowDelete) {
                LOG "UpdateTable" "Opérations effectuées: $updateSuccess/$updateCount mises à jour, $insertSuccess/$insertCount insertions, $deleteSuccess/$deleteCount suppressions"
            } else {
                LOG "UpdateTable" "Opérations effectuées: $updateSuccess/$updateCount mises à jour, $insertSuccess/$insertCount insertions (suppressions désactivées)"
            }
            
            $totalSuccess = $updateSuccess + $insertSuccess + $deleteSuccess
            if ($totalSuccess -eq $totalOperations) {
                LOG "UpdateTable" "Toutes les opérations ont été exécutées avec succès"
            } else {
                $totalErrors = $totalOperations - $totalSuccess
                WRN "UpdateTable" "$totalErrors opération(s) ont échoué sur $totalOperations"
            }
        } else {
            if ($allowDelete) {
                LOG "UpdateTable" "Opérations simulées: $updateCount mises à jour, $insertCount insertions, $deleteCount suppressions"
            } else {
                LOG "UpdateTable" "Opérations simulées: $updateCount mises à jour, $insertCount insertions (suppressions désactivées)"
            }
        }
    } else {
        LOG "UpdateTable" "Aucune modification à effectuer"
    }
}