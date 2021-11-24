$DomAdminPwFile = "DomAdminPw.txt"
$DomAdminKeyFile = "DomAdminAes.key"
$DomAdminUserFile = "DomAdminName.txt"

Write-Host "Im folgenden bitte die Zugangsdaten des Domänen-Administrators mit vorangeschriebener Domäne im Benutzernamen angeben. Beispiel: Domain\Administrator"
$DomAdminCred = Get-Credential

$DomAdminKey = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($DomAdminKey)
$DomAdminKey | out-file $DomAdminKeyFile

$DomAdminCred.Password | ConvertFrom-SecureString -key $DomAdminKey | Out-File $DomAdminPwFile

$DomAdminCred.GetNetworkCredential().Domain | Out-File $DomAdminUserFile
$DomAdminCred.GetNetworkCredential().UserName | Out-File $DomAdminUserFile -Append

Write-Host "Fertig. Konsole kann geschlossen werden." -ForegroundColor Green
Read-Host
