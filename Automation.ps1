<# 
    Variable declaration
 #>

# Editable
$LocBase = "C:\TSD.CenterVision\Software\_Scripts\RapidFireTools-ComputerDataCollector-Automation"
$Destination = "C:\TSD.CenterVision\Software\_Scripts\"
$MailRecipientFile = "MailRecipient.txt"

# Dependent
$LocNdcdc = "$( $LocBase )\NetworkDetectiveComputerDataCollector" # Path/folder depending on Downloader.ps1

# Other
$MailRecipient = Get-Content -Path $MailRecipientFile

<# 
    Move data from usb stick to C:
 #>

$Loc = [System.IO.FileInfo]((Get-Location).Path)
Copy-Item -Path $Loc -Destination $Destination -Recurse -Force

<# 
    Decode credentials
 #>

$DomAdminPwFile = "$( $LocBase )\DomAdminPw.txt"
$DomAdminKeyFile = "$( $LocBase )\DomAdminAes.key"
$DomAdminUserFile = "$( $LocBase )\DomAdminName.txt"

$DomAdminKey = Get-Content -Path $DomAdminKeyFile
$DomAdminPw = Get-Content -Path $DomAdminPwFile
$DomAdminUserContents = Get-Content -Path $DomAdminUserFile
$DomAdminDomain = $DomAdminUserContents[0]
$DomAdminUserName = $DomAdminUserContents[1]

$DomAdminCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$( $DomAdminDomain )\$( $DomAdminUserName )", ($DomAdminPw | ConvertTo-SecureString -Key $DomAdminKey)

<# 
    Decode Gmail mailbox credential
 #>

$MailPwFile = "MailPw.txt"
$MailKeyFile = "MailAes.key"
$MailUserFile = "MailUser.txt"

$MailKey = Get-Content -Path $MailKeyFile
$MailPw = Get-Content -Path $MailPwFile
$MailUser = Get-Content -Path $MailUserFile

$MailCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $MailUser, ($MailPw | ConvertTo-SecureString -Key $MailKey)

<# 
    Ask for InvId
 #>

$InvId = Read-Host -Prompt "Inventar-ID"

<# 
    Start RapidFire Tools
 #>

$Proc = Start-Process -FilePath "$( $LocNdcdc )\runLocal.bat" -Credential $DomAdminCred -PassThru
$Proc.WaitForExit()

<# 
    Rename file to contain InvId
 #>

$CdfOrig = Get-ChildItem -Path $LocNdcdc -Filter "$( $env:COMPUTERNAME )*.cdf"
$CdfNewName = "$( $env:COMPUTERNAME )_$( $InvId ).cdf"
Rename-Item -Path $CdfOrig.FullName -NewName $CdfNewName
$Cdf = Get-ChildItem "$( $LocNdcdc )\$( $CdfNewName )"

<# 
    Send mail
 #>

$mailParams = @{
    SmtpServer                 = 'smtp.gmail.com'
    Port                       = '587'
    UseSSL                     = $true
    Credential                 = $MailCred
    From                       = $MailUser
    To                         = $MailRecipient
    Subject                    = "RapidFire-Tools | Computername: $( $env:COMPUTERNAME ) | Inventar-ID: $( $InvId )"
    # Body                       = 'Siehe Anhang.'
    DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
    Attachment                  = $Cdf.FullName
}

Send-MailMessage @mailParams

<# 
    End
 #>

Write-Host "Fertig. Die Konsole kann geschlossen werden." -ForegroundColor Green
Read-Host
