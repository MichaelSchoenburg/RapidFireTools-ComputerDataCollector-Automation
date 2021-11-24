$MailPwFile = "MailPw.txt"
$MailKeyFile = "MailAes.key"
$MailUserFile = "MailUser.txt"

Write-Host "Im folgenden bitte die Zugangsdaten für das Gmail-Postfach ein. Sollte MFA aktiv sein, geben Sie im folgenden als Passwort ein AppPasswort ein. Benutzername ist Ihre E-Mail-Adresse."
$MailCred = Get-Credential

$MailKey = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($MailKey)
$MailKey | out-file $MailKeyFile

$MailCred.Password | ConvertFrom-SecureString -key $MailKey | Out-File $MailPwFile

$MailCred.GetNetworkCredential().UserName | Out-File $MailUserFile

Write-Host "Fertig. Konsole kann geschlossen werden." -ForegroundColor Green
Read-Host
