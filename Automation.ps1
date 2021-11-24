<#
.SYNOPSIS
    Automation for RapidFire Tools Computer Data Collector. 
.DESCRIPTION
    This Script works through following steps:
    1. Decode credentials for local admin and mail account.
    2. Copy Computer Data Collector to the local machine.
    3. Executes the Computer Data Collector as administrator.
    4. Sends the results to a mail address.
.INPUTS
    None
.OUTPUTS
    None
.NOTES
    Author: Michael Schönburg
    GitHub Repository: https://github.com/MichaelSchoenburg/RapidFireTools-ComputerDataCollector-Automation
#>

#region INITIALIZATION

# Editable
$LocBase = "C:\TSD.CenterVision\Software"
$MailRecipientFile = "MailRecipient.txt"

# Dependent
$LocNdcdc = "$( $LocBase )\NetworkDetectiveComputerDataCollector" # Path/folder depending on Downloader.ps1
$LocNdcdcSource = "NetworkDetectiveComputerDataCollector" # Path/folder depending on Downloader.ps1

# Other
$MailRecipient = Get-Content -Path $MailRecipientFile

#endregion INITIALIZATION
#---------------------------------------------------------------------------------------------------------
#region FUNCTIONS

function Write-ConsoleLog {
    <#
    .SYNOPSIS
    Logs an event to the console.
    
    .DESCRIPTION
    Writes text to the console with the current date (US format) in front of it.
    
    .PARAMETER Text
    Event/text to be outputted to the console.
    
    .EXAMPLE
    Write-ConsoleLog -Text 'Subscript XYZ called.'
    
    Long form
    .EXAMPLE
    Log 'Subscript XYZ called.
    
    Short form
    #>
    [alias('Log')]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
        Position = 0)]
        [string]
        $Text
    )

    # Save current VerbosePreference
    $VerbosePreferenceBefore = $VerbosePreference

    # Enable verbose output
    $VerbosePreference = 'Continue'

    # Write verbose output
    Write-Verbose "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - $( $Text )"

    # Restore current VerbosePreference
    $VerbosePreference = $VerbosePreferenceBefore
}

#endregion FUNCTIONS
#---------------------------------------------------------------------------------------------------------
#region EXECUTION

<# 
    Copy the Network Detective Computer Data Collector from usb stick to target folder
 #>

Log "Copying NDCDC from $( $LocNdcdcSource ) to $( $LocNdcdc )"
Copy-Item -Path $LocNdcdcSource -Destination $LocBase -Recurse -Force
Log "Finished copy."

<# 
    Decode credentials
 #>

Log "Decoding domain admin credentials."

$DomAdminPwFile = "DomAdminPw.txt"
$DomAdminKeyFile = "DomAdminAes.key"
$DomAdminUserFile = "DomAdminName.txt"

$DomAdminKey = Get-Content -Path $DomAdminKeyFile
$DomAdminPw = Get-Content -Path $DomAdminPwFile
$DomAdminUserContents = Get-Content -Path $DomAdminUserFile
$DomAdminDomain = $DomAdminUserContents[0]
$DomAdminUserName = $DomAdminUserContents[1]

Log "Creating PSCredential object with domain admin credentials."

$DomAdminCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$( $DomAdminDomain )\$( $DomAdminUserName )", ($DomAdminPw | ConvertTo-SecureString -Key $DomAdminKey)

Log "Domain admin credential section finished."

<# 
    Decode Gmail mailbox credential
 #>

Log "Decoding gmail mailbox credentials."

$MailPwFile = "MailPw.txt"
$MailKeyFile = "MailAes.key"
$MailUserFile = "MailUser.txt"

$MailKey = Get-Content -Path $MailKeyFile
$MailPw = Get-Content -Path $MailPwFile
$MailUser = Get-Content -Path $MailUserFile

Log "Creating PSCredential object from gmail mailbox credentials."

$MailCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $MailUser, ($MailPw | ConvertTo-SecureString -Key $MailKey)

Log "Gmail mailbox credential section finished."

<# 
    Ask for InvId
 #>

Log "Asking for InvId."

$InvId = Read-Host -Prompt "Inventar-ID"

<# 
    Start RapidFire Tools
 #>

Log "Starting NDCDC at $( $LocNdcdc )\runLocal.bat"

$Proc = Start-Process -FilePath "$( $LocNdcdc )\runLocal.bat" -Credential $DomAdminCred -PassThru

Log "Started. Now waiting for Process."

$Proc.WaitForExit()

Log "Done waiting."

<# 
    Rename file to contain InvId
 #>

Log "Renaming CDF file."

$CdfOrig = Get-ChildItem -Path $LocNdcdc -Filter "$( $env:COMPUTERNAME )*.cdf"
$CdfNewName = "$( $env:COMPUTERNAME )_$( $InvId ).cdf"
Rename-Item -Path $CdfOrig.FullName -NewName $CdfNewName

Log "Retrieving new CDF file."

$Cdf = Get-ChildItem "$( $LocNdcdc )\$( $CdfNewName )"

<# 
    Send mail
 #>

 Log "Building mail."

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

Log "Sending mail."

Send-MailMessage @mailParams

<# 
    End
 #>

Log "Reached the end of the script."

Write-Host "Fertig. Die Konsole kann geschlossen werden." -ForegroundColor Green
Read-Host

#endregion EXECUTION
