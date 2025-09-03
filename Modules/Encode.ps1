function Encode {
    param ( [string]$pw )

    $epw = New-Object System.Text.StringBuilder
    foreach ($char in $pw.ToCharArray()) {
        $ascii = [int][char]$char
        
        if ($ascii -ge 97 -and $ascii -le 122) { # a-z
            [void]$epw.Append([char](219 - $ascii))
        }
        elseif ($ascii -ge 65 -and $ascii -le 90) { # A-Z
            [void]$epw.Append([char](155 - $ascii))
        }
        elseif ($ascii -ge 48 -and $ascii -le 57) { # 0-9
            [void]$epw.Append([char](105 - $ascii))
        }
        else {
            [void]$epw.Append($char)
        }
    }
    return $epw.ToString()
} # Hook   param -password                                                    Return : [string]    >> Decode un Password crypte

function New-SecurePassword {
    param( [int]$length = 18 )

    $lower = 'abcdefghijklmnopqrstuvwxyz'.ToCharArray()
    $upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.ToCharArray()
    $digits = '0123456789'.ToCharArray()
    $special = '!@#$%^&*()-_=+[]{}|;:,.<>?'.ToCharArray()

    # Assurer au moins un de chaque catégorie
    $passwordChars = @()
    $passwordChars += Get-Random -InputObject $lower -Count 1
    $passwordChars += Get-Random -InputObject $upper -Count 1
    $passwordChars += Get-Random -InputObject $digits -Count 1
    $passwordChars += Get-Random -InputObject $special -Count 1

    # Reste des caractères tirés de l'ensemble complet
    $allChars = $lower + $upper + $digits + $special
    $remainingCount = $length - $passwordChars.Count
    $passwordChars += Get-Random -InputObject $allChars -Count $remainingCount

    # Mélanger aléatoirement
    $password = ($passwordChars | Get-Random -Count $length) -join ''

    return $password
} # code   param length                                                       Return : (string]    -> genere un password complexe aleatoire
