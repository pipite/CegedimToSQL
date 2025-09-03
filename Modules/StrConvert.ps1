# --------------------------------------------------------
#               Utilitaires de conversion
# --------------------------------------------------------

# Cache global pour les formats de date détectés
$script:DateFormatCache = @{}

# Fonction optimisée pour convertir les dates
function ConvertDateToString {
    param([object]$value, [string]$formatOut, [string]$locale = "FR")
    
    if ([string]::IsNullOrWhiteSpace($formatOut) -or $value -eq $null) {
        return $value
    }
    
    # Gérer les valeurs vides ou "-"
    $stringValue = $value.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($stringValue) -or $stringValue -eq "-") {
        return $value
    }
    
    # Vérifier si la valeur est déjà un DateTime
    if ($value -is [DateTime]) {
        return $value.ToString($formatOut)
    }
    
    # Utiliser le cache si disponible
    $cacheKey = "$stringValue|$locale"
    if ($script:DateFormatCache.ContainsKey($cacheKey)) {
        $cachedFormat = $script:DateFormatCache[$cacheKey]
        if ($cachedFormat -ne $null) {
            try {
                $dateObj = [DateTime]::ParseExact($stringValue, $cachedFormat, [System.Globalization.CultureInfo]::InvariantCulture)
                return $dateObj.ToString($formatOut)
            } catch {
                # Le format en cache ne fonctionne plus, le supprimer
                $script:DateFormatCache.Remove($cacheKey)
            }
        } else {
            # Valeur en cache comme non-date
            return $value
        }
    }
    
    try {
        # Formats optimisés (les plus courants en premier)
        $dateFormats = if ($locale -eq "US") {
            @(
                "yyyy-MM-dd HH:mm:ss.f",      # Excel export format
                "MM/dd/yyyy",                 # US format - PRIORITÉ
                "M/d/yyyy",                   # US format simple
                "yyyy-MM-dd",                 # ISO format
                "dd/MM/yyyy",                 # FR format - SECONDAIRE
                "d/M/yyyy",                   # FR format simple
                "MM/dd/yyyy HH:mm:ss",        # US avec heure
                "dd/MM/yyyy HH:mm:ss",        # FR avec heure
                "yyyy-MM-ddTHH:mm:ss"         # ISO avec T
                "MM/yyyy"                     # Mois/année uniquement
            )
        } else {
            @(
                "yyyy-MM-dd HH:mm:ss.f",      # Excel export format
                "dd/MM/yyyy",                 # FR format - PRIORITÉ
                "d/M/yyyy",                   # FR format simple
                "yyyy-MM-dd",                 # ISO format
                "MM/dd/yyyy",                 # US format - SECONDAIRE
                "M/d/yyyy",                   # US format simple
                "dd/MM/yyyy HH:mm:ss",        # FR avec heure
                "MM/dd/yyyy HH:mm:ss",        # US avec heure
                "yyyy-MM-ddTHH:mm:ss"         # ISO avec T
                "MM/yyyy"                     # Mois/année uniquement
            )
        }
        
        # Essayer de parser avec chaque format
        foreach ($format in $dateFormats) {
            try {
                $dateObj = [DateTime]::ParseExact($stringValue, $format, [System.Globalization.CultureInfo]::InvariantCulture)
                # Mettre en cache le format qui a fonctionné
                $script:DateFormatCache[$cacheKey] = $format
                return $dateObj.ToString($formatOut)
            } catch {
                # Continuer avec le format suivant
            }
        }
        
        # Si aucun format spécifique ne fonctionne, essayer la conversion générique
        $dateObj = [DateTime]$stringValue
        $script:DateFormatCache[$cacheKey] = "generic"
        return $dateObj.ToString($formatOut)
        
    } catch {
        # Mettre en cache comme non-date pour éviter les futurs essais
        $script:DateFormatCache[$cacheKey] = $null
        return $value
    }
}

# Fonction pour vider le cache des formats de date
function Clear-DateFormatCache {
    $script:DateFormatCache.Clear()
    DBG "Clear-DateFormatCache" "Cache des formats de date vidé"
}

# Fonction utilitaire pour parser les dates avec gestion d'erreurs
function ConvertTo-SafeDate {
    param(
        [string]$dateString,
        [string]$format,
        [string]$matricule,
        [string]$fieldName,
        [string]$functionName
    )

    $culture = [System.Globalization.CultureInfo]::InvariantCulture

    if ([string]::IsNullOrWhiteSpace($dateString)) {
        return @{ Success = $true; Date = $null; Error = $null }
    }
    
    try {
        $parsedDate = [datetime]::ParseExact($dateString, $format, $culture)
        return @{ Success = $true; Date = $parsedDate; Error = $null }
    } catch {
        $errorMsg = "Matricule [$matricule] : Erreur parsing [$fieldName]  : $dateString"
        ERR $functionName $errorMsg
        return @{ Success = $false; Date = $null; Error = $errorMsg }
    }
}

# Fonction utilitaire pour parser les montants avec gestion d'erreurs
function ConvertTo-SafeAmount {
    param(
        [string]$amountString,
        [string]$matricule,
        [string]$annee,
        [string]$cat,
        [string]$functionName,
        [double]$defaultValue = 0
    )
    
    if ([string]::IsNullOrWhiteSpace($amountString)) {
        return @{ Success = $true; Amount = $defaultValue; Error = $null }
    }
    
    # Nettoyage du montant : supprime les espaces insécables et remplace , par .
    $cleanAmount = $amountString -replace '\s', '' -replace ',', '.'
    
    try {
        $parsedAmount = [double]::Parse($cleanAmount, [System.Globalization.CultureInfo]::InvariantCulture)
        return @{ Success = $true; Amount = $parsedAmount; Error = $null }
    } catch {
        $errorMsg = "Matricule [$matricule] [$annee] [$cat] : Montant invalide : $amountString"
        ERR $functionName $errorMsg
        return @{ Success = $false; Amount = $defaultValue; Error = $errorMsg }
    }
}
# Fonction utilitaire pour parser les dates avec gestion d'erreurs
function Convert-ToSafeDate2 {
    param (
        [string]$DateStr
    )

    # Liste de formats courants (FR et US)
    $formats = @(
        'dd/MM/yyyy',
        'MM/dd/yyyy',
        'yyyy-MM-dd',
        'dd-MM-yyyy',
        'dd/MM/yyyy HH:mm:ss',
        'MM/dd/yyyy HH:mm:ss',
        'yyyy-MM-dd HH:mm:ss',
        'dd-MM-yyyy HH:mm:ss'
    )

    foreach ($format in $formats) {
        if ([datetime]::TryParseExact($DateStr, $format, $null, 'None', [ref]$dt)) {
            return $dt.ToString("yyyy-MM-dd HH:mm:ss")  # Format SQL sûr
        }
    }

    # Dernier recours : tentative auto
    if ([datetime]::TryParse($DateStr, [ref]$dt2)) {
        return $dt2.ToString("yyyy-MM-dd HH:mm:ss")
    }

    # Si rien ne marche
    return $null
}
function StrConvert {
    param([string]$s)

    if (-not $s) { return $s }

    $replacements = @{
        'é' = 'e'; 'è' = 'e'; 'ë' = 'e'
        'à' = 'a'; 'â' = 'a'
        'ê' = 'e'; 'ô' = 'o'
        'ç' = 'c'; 'ù' = 'u'
        "'" = ' '; 'ï' = 'i'
        '&' = 'et'; '/' = '.'
    }

    foreach ($key in $replacements.Keys) {
        $s = $s -replace [regex]::Escape($key), $replacements[$key]
    }

    return $s
}