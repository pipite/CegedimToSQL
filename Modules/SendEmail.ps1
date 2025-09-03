<# 
# Parametre pour l'envoi de mails (Protocoles possible : Microsoft.Graph ou SMTP)
# Le parametre "emailmode" permet de choisir le mode d'emission d'un mail (GRAPH ou SMTP)
# Envoi de mail si sendemail = "yes" / "no"
[email]
# AppId , AppSecret, TenantId, FromGUID sont encodés
sendemail    = yes
destinataire = adremaildetest@gmail.com
Subject      = Synchro AD BERTIN
emailmode    = SMTP

# Info pour SMTP
expediteur   = adremaildetest@gmail.com
server       = xxxxx.xxxxx.xxxx  (serveur encode)
port         = 587
password     = atma zfir ndpf cplx (encode)

# Info pour GRAPH
#AppId        = 6vu5u86u-7z7y-5x9v-y820-5054826z3u00 (encode)
#TenantId     = 05v49vxu-z10x-5uxy-068z-yz2u914x53z4 (encode)
#clientSecret = D6A1J~jeekZ2Rb.mNIGnoiVJ6IsD5.KWcexGDwmT (encode)
#senderId     = 481u7w13-74xw-5556-yw41-yuz95w55241u (encode)
#>

# --------------------------------------------------------
#               Gestion Emails
# --------------------------------------------------------
function SendEmail {
	param ($subject, $message)

	if ( $script:cfg["email"]["sendemail"] -eq "no") { return }

	if ( $script:cfg["email"]["emailmode"] -eq "GRAPH") {
		SendEmailGraph $subject $message
	} elseif ( $script:cfg["email"]["emailmode"] -eq "SMTP" ) {
		SendEmailSMTP $subject $message
	}
}

function IsEmail {
	param ( [string]$str, $matricule )
	$regex = '^[\w\.-]+@[\w\.-]+\.\w{2,}$'

	if ($str -match $regex) { 
		return $true 
	} else {
		if ($str.Trim() -match $regex) { 
			WRN "IsEmail" "Matricule : [$matricule] UPN O365 [$str] valide mais avec des espaces en debut ou fin de chaine"
			return $true 
		} 
	}
	return $false
}

function SendEmailSMTP {
	param ($subject, $message)
    try {
        if ( $script:ERREUR -gt 0 -or $script:WARNING -gt 0 ) {
            $corpsTexte = $message -join "`r`n"
			[string[]]$To = $($script:cfg["email"]["destinataire"]).Split(',')
			$params = @{
				From       = $script:cfg["email"]["expediteur"]
				To         = $To
				Subject    = $subject
				Body       = $corpsTexte
			}

			if ( -not [string]::IsNullOrEmpty($script:cfg["email"]["password"]) ) {
				$Password = Encode $script:cfg["email"]["password"]
				$Creds = New-Object System.Management.Automation.PSCredential ($script:cfg["email"]["expediteur"], (ConvertTo-SecureString $Password -AsPlainText -Force))
				$params['Credential'] = $Creds
			}
			if ( $script:cfg["email"]["UseSSL"] -eq "true" ) {
				$params['UseSSL'] = $true
			}
			
			if ( -not [string]::IsNullOrEmpty($script:cfg["email"]["server"]) ) {
				$server = Encode $script:cfg["email"]["server"]
				$params['SmtpServer'] = $server
			}
			if ( -not [string]::IsNullOrEmpty($script:cfg["email"]["port"]) ) {
				$params['port']       = $script:cfg["email"]["port"]
			}
			
            DBG "SendEmailSMTP" "Send-MailMessage est obsolete mais n'a pas encore d'alternative native dans Powershell"
            Send-MailMessage @params -ErrorAction Stop -WarningAction Ignore
            LOG "SendEmailSMTP" "Email envoye a $($script:cfg["email"]["destinataire"])"
        }
    } catch {
        $script:MailErr = $true
        QUITEX "SendEmailSMTP" "$($_.Exception.Message)"
    }
} # code   param N/A                                                          Return : N/A         >> Envoi un mail en cas d'erreur                               

function SendEmailGraph {
	param ($subject, $message)

	# Chargement de la configuration
	$AppId        = Encode $script:cfg["email"]["AppId"]
	$TenantId     = Encode $script:cfg["email"]["TenantId"]
	$clientSecret = Encode $script:cfg["email"]["clientSecret"]
	$senderId     = Encode $script:cfg["email"]["senderId"]
	$toField      = $script:cfg["email"]["Destinataire"]

	# Contenu texte de l'email (liste de lignes -> texte)
	$corpsTexte = $message -join "`n"

	# Construction de la liste des destinataires
	$recipientList = @()
	if ($toField -is [string]) {
		$recipientList = $toField -split "," | ForEach-Object { $_.Trim() }
	} elseif ($toField -is [array]) {
		$recipientList = $toField
	}

	# Création de la structure correcte pour Graph
	$toRecipientsArray = @()
	foreach ($email in $recipientList) {
		$toRecipientsArray += @{
			emailAddress = @{
				address = $email
			}
		}
	}

	# Authentification Graph API
	$tokenBody = @{
		grant_type    = "client_credentials"
		client_id     = $AppId
		client_secret = $clientSecret
		scope         = "https://graph.microsoft.com/.default"
	}

	$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
									   -Method POST `
									   -Body $tokenBody `
									   -ContentType "application/x-www-form-urlencoded"

	$accessToken = $tokenResponse.access_token

	# Construction du message
	$mail = @{
		message = @{
			subject = $subject
			body = @{
				contentType = "Text"
				content     = $corpsTexte
			}
			toRecipients = $toRecipientsArray
		}
		saveToSentItems = $true
	} | ConvertTo-Json -Depth 10

	# Envoi du message
	try {
		Invoke-RestMethod -Method POST `
			-Uri "https://graph.microsoft.com/v1.0/users/$senderId/sendMail" `
			-Headers @{ Authorization = "Bearer $accessToken" } `
			-Body $mail `
			-ContentType "application/json"

		LOG "SendEmailGraph" "Email envoyé à $($recipientList -join ', ')"
	}
	catch {
		QUITEX "SendEmailGraph" "Erreur lors de l'envoi : $($_.Exception.Message)"
	}
}
