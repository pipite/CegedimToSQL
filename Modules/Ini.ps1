function Get-Inifile {
    param ( [string]$Path )

    try {
        #DBG "Get-Inifile" "Hachage du fichier $Path"
        if (Test-Path $Path -PathType Leaf) {
            $ini = @{}
            $section = ""

            $lines = Get-Content -Path $Path -Encoding UTF8
            foreach ($ln in $lines) {
                $line = $ln.Trim()
                if (-not $line.startsWith("#")) {
                    if ($line -match "^\[([^\]]+)\]") {
                        $section = $matches[1].ToLower().Trim()
                        $ini[$section] = @{}
                    } elseif ($line -match "^([^=]+)=(.*)") {
                        $name = $matches[1].ToLower().Trim()
                        $value = $matches[2].Trim()
                        $ini[$section][$name] = $value.Trim()
                    }
                }
            }
            return $ini
        } else {
            ERR "Get-Inifile" "Le fichier $Path n'existe pas."
            return $null
        }
    } catch { QUITEX "Get-Inifile" "$($_.Exception.Message)" }
} # code   param -path                                                        Return : [string[]]  -> Cree une hashtable d'un fichier (path) .ini
function Add-IniFiles {
    param ( $ini, [string]$inifile )

    try {
        $mergedIni = @{}

	    if ($inifile -ne $null -and (Test-Path $inifile -PathType Leaf)) {
		    $newini = Get-Inifile -Path $inifile
            foreach ($section in $ini.Keys + $newini.Keys | Select-Object -Unique) {
                $mergedIni[$section] = @{}
                if ($ini.ContainsKey($section))   { $mergedIni[$section] += $ini[$section] }
                if ($newini.ContainsKey($section)) { $mergedIni[$section] += $newini[$section] }
            }
            return $mergedIni
	    }
        Write-Host "Add-IniFiles : Le fichier $inifile n'existe pas"
        return $ini
    } catch { Write-Host "Add-IniFiles - $($_.Exception.Message)" }
} # code   param -ini -inifile                                                Return : [string[]]  -> Charge les cle (ini) d'un fichier (inifile)
