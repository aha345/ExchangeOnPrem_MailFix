#####################################################################
#                                                                   #
#                                                                   #
#                       Script created by                           #
#                                                                   #
#                              Kenkot                               #
#                                &                                  #
#                              Anderi                               #
#                                                                   #
#                               #AMS                                #
#                                                                   #
#                                                                   #
#####################################################################


function Import-CsvFile {
    param(
        [string] $DefaultPath = "$($env:USERPROFILE)\Desktop",
        [string] $DefaultFileName = "$(Get-Date -Format 'yyyyMMdd')-$($ADUser.Name).csv"
    )

    do {
        $filename = Read-Host -Prompt "Please enter the filename (default: $DefaultFileName):"
        if (-not $filename) { $filename = $DefaultFileName }
        $filepath = Join-Path -Path $DefaultPath -ChildPath $filename
        if (-not (Test-Path $filepath)) {
            Write-Warning "File not found. Please try again."
        }
    } until (Test-Path $filepath)

    return Import-Csv -Path $filepath
}

# Define the date format
$dateFormat = "dd.MM.yyyy HH.mm.ss"

# Get the current date and time
$currentDate = Get-Date

# Clear the console window
Clear-Host

# Add some variables we'll use later
$CurrentHostname = hostname

# Check if the user running the script has access to the on-premises PowerShell environment
try {
    Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
}
catch {
    Write-Warning "You don't have access to the on-premises PowerShell environment."
    Start-Sleep -Seconds 10
    exit
}

# Import the CSV file
$importedData = Import-CsvFile

# Access the data from the CSV file
foreach ($item in $importedData) {
    $Name = $item.Name
    $UserUPN = $item.UPN
    $SAMAccountName = $item.SAMAccountName
    $Mail = $item.Mail
    $TargetAddress = $item.TargetAddress
    $msexchRecipientDisplayType = $item.msexchRecipientDisplayType
    $msExchRecipientTypeDetails = $item.msExchRecipientTypeDetails
    $ExchangeGuid = $item.ExchangeGuid
    $AllSortedProxyAddresses = $item.AllSortedProxyAddresses -split ', '
    $ExistingMailOnMicrosoftAddress = $item.ExistingMailOnMicrosoftAddress
}

# Check if $item.mail is blank
if ($Mail -notlike '*@*'){
    $Mail = $UserUPN
}

# Check if $targetaddress is blank
if ($TargetAddress -notlike '*@*'){
    $TargetAddress = $ExistingMailOnMicrosoftAddress.Insert(0,"SMTP:")
}

# Prompt the user to select mailbox type
$MailboxType = $(Write-Host "What type is the mailbox?" -ForegroundColor Magenta -NoNewLine) + $(Write-Host "(Type 'User', 'Shared', 'Room' or 'Equipment'.)" -ForegroundColor Yellow -NoNewLine) + $(Write-Host ": " -ForegroundColor Magenta -NoNewLine; Read-Host)

# Check if a recipient object (such as a mailbox) exists for the given UPN, Email, or Alias
$Mailbox = $null
try {
    if (-not $Mailbox) {
        $Mailbox = Get-Recipient -ResultSize 1 -Filter "UserPrincipalName -eq '$UserUPN' -or EmailAddresses -like 'SMTP:$Mail' -or Alias -eq '$SAMAccountName'" -ErrorAction Stop
        Start-Sleep -Seconds 5
    }
} catch {
    Write-Warning "Get-Recipient can't find mailbox"
}

# Check if the mailbox exists
if ($Mailbox) {
    $CurrentRecipientType = $Mailbox.RecipientType
    $(Write-Host "Mailbox for UPN '$UserUPN' exists with RecipientType: " -ForegroundColor Gray -NoNewLine) + $(Write-Host "$($Mailbox.RecipientTypeDetails)" -ForegroundColor Green)
} else {
    Write-Warning "Mailbox for UPN '$UserUPN' does not exist"
}

# Get mailbox size
$MailboxStatistics = Get-MailboxStatistics -Identity $UserUPN -ErrorAction SilentlyContinue
$MailboxSize = $MailboxStatistics.TotalItemSize

if ($MailboxStatistics) {
    # Output mailbox size
    Write-Host "Mailbox size: $($MailboxSize.ToString())" -ForegroundColor Red
    $(Write-Host "$($UserUPN) has an " -ForegroundColor Cyan -NoNewLine) + $(Write-Host "On-Prem Mailbox" -ForegroundColor Yellow)
} else {
    $(Write-Host "$($UserUPN) does not have an " -ForegroundColor Green -NoNewLine) + $(Write-Host "On-Prem Mailbox" -ForegroundColor Yellow)
}

# Check if the mailbox exists
if ($MailboxStatistics) {
    # Prompt user to confirm mailbox deletion
    $confirmation = $(Write-Host "Do you want to delete the on-prem mailbox for $($UserUPN)?" -ForegroundColor Magenta -NoNewLine) + $(Write-Host "(Type 'Yes' to confirm, or any other key to skip)" -ForegroundColor Yellow -NoNewLine) + $(Write-Host ": " -ForegroundColor Magenta -NoNewLine; Read-Host)

    if ($confirmation -eq "Yes") {
        # Get the current user's distinguished name (DN) 
        $currentuser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $username = $currentuser.split('\')[-1]

        # Check if the user has the "Mailbox Import Export" role
        $roleName = "Mailbox Import Export"
        $hasRole = (Get-ManagementRoleAssignment -Role $roleName -RoleAssignee $userName -RoleAssigneeType User -ErrorAction SilentlyContinue) -ne $null

        $null = Add-MailboxPermission -Identity $UserUPN -User $username -AccessRights FullAccess -InheritanceType All -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

        Start-Sleep -Seconds 5

        if ($hasRole) {
        } else {
            # Assign the "Mailbox Import Export" role to the user
            New-ManagementRoleAssignment -Role $roleName -User $userName
            Write-Host "The '$roleName' role has been assigned to the user '$userName'." -ForegroundColor Yellow
            Start-Sleep -Seconds 5
        }

        # Prompt user to confirm mailbox export
        $ConfirmExport = $(Write-Host "Do you want to export the on-prem mailbox for $($UserUPN)? " -ForegroundColor Magenta -NoNewLine) + $(Write-Host "(Type 'Yes' to confirm, or any other key to skip)" -ForegroundColor Yellow -NoNewLine) + $(Write-Host ": " -ForegroundColor Magenta -NoNewLine; Read-Host)

        if ($ConfirmExport -eq "Yes") {
            $(Write-Host "") + $(Write-Host "Open browser and go to the following address to export the mailbox: localhost/ecp" -ForegroundColor Yellow) + $(Write-Host "")
            Start-Sleep -Seconds 5

            # Path for export mailbox content to a PST file
            $PSTPath = "\\$CurrentHostname\$env:UserName\Desktop\$(Get-Date -Format 'yyyyMMdd')-$UserUPN.pst"
            $(Write-Host "The following path is needed for export of the .pst file:" -ForegroundColor Yellow) + $(Write-Host "") + $(Write-Host "$($PSTPath)" -ForegroundColor Cyan)
            
            # Promt user to add path to clipboard
            $ConfirmClip = $(Write-Host "Do you want to copy the path to clipboard? " -ForegroundColor Magenta -NoNewLine) + $(Write-Host "(Type 'Yes' to confirm, or any other key to skip)" -ForegroundColor Yellow -NoNewLine) + $(Write-Host ": " -ForegroundColor Magenta -NoNewLine; Read-Host)
            if ($ConfirmClip -eq "Yes") {
                $PSTPath | clip.exe
            }

            # Check if mailbox is exported
            $Exported = $false
            while (-not $Exported) {
                try {
                    $lastDay = $currentDate.AddDays(-7)
                    $ExportStatus = Get-MailboxExportRequest -Mailbox $SamAccountName | Where-Object {$_.WhenCreated -gt $lastDay}

                    do {
                        $stats = Get-MailboxExportRequestStatistics -Identity $SamAccountName -WarningAction SilentlyContinue
                        Write-Warning "Export status: $($stats.Status), Percent complete: $($stats.PercentComplete)!"
                        Start-Sleep -Seconds 5
                    } while ($stats.Status -eq "Queued" -or $stats.Status -eq "InProgress")

                    if ($ExportStatus.status -eq "Completed"){
                        $Exported = $true
                        Write-Host ""
                        Write-Host "Exchange has completed the export of the mailbox" -ForegroundColor Green
                    }
                }
                catch {
                    Start-Sleep -Seconds 5
                    
                }
            }

            if ($Exported = $true) {
                # Prompt user to confirm disabling of mailbox
                $ConfirmDisable = $(Write-Host "Do you want to disable the on-prem mailbox for $($UserUPN)? " -ForegroundColor Magenta -NoNewLine) + $(Write-Host "(Type 'Yes' to confirm, or any other key to skip)" -ForegroundColor Yellow -NoNewLine) + $(Write-Host ": " -ForegroundColor Magenta -NoNewLine; Read-Host)

                if ($ConfirmDisable = "Yes") {
                    # Disable the on-prem mailbox
                    Disable-Mailbox -Identity $Mailbox.PrimarySmtpAddress -Confirm:$false
                    Start-Sleep -Seconds 5
                    Write-Host "The on-prem mailbox for $($UserUPN) has been disabled." -ForegroundColor Yellow
                    Start-Sleep -Seconds 5
                }
            }           
        }
    } else {
        Write-Warning "Skipping on-prem mailbox deletion for $($UserUPN)."
    }
} else {
    Write-Warning "No on-prem mailbox found for $($UserUPN)."
}

# Get the current values of the AD user fields
$ADUser = Get-ADUser -Identity $SAMAccountName -Properties mail, TargetAddress, ProxyAddresses, msExchRecipientDisplayType, msExchRecipientTypeDetails

# Prompt user to confirm blanking the fields
$confirmationBlank = $(Write-Host "Do you want to blank Exchange properties for $($UserUPN)?" -ForegroundColor Magenta -NoNewLine) + $(Write-Host "(Type 'Yes' to confirm, or any other key to skip)" -ForegroundColor Yellow -NoNewLine) + $(Write-Host ": " -ForegroundColor Magenta -NoNewLine; Read-Host)

if ($confirmationBlank -eq "Yes") {
    # Blank the fields using Set-ADUser
    Set-ADUser -Identity $SAMAccountName -Clear mail, TargetAddress, ProxyAddresses, msExchRecipientDisplayType, msExchRecipientTypeDetails
    Write-Host "Fields have been blanked for $($UserUPN)." -ForegroundColor Gray
} else {
    Write-Host "Skipping blanking fields for $($UserUPN)." -ForegroundColor Cyan
}

# Get the updated values of the AD user fields
$UpdatedADUser = Get-ADUser -Identity $SAMAccountName -Properties mail, TargetAddress, ProxyAddresses, msExchRecipientDisplayType, msExchRecipientTypeDetails

# Initialize variables
$ADmsExchRecipientDisplayType = 0
$ADmsExchRecipientTypeDetails = 0

# Set variables based on mailbox type
if ($MailboxType -eq "User") {
    $ADmsExchRecipientDisplayType = 1073741824
    $ADmsExchRecipientTypeDetails = 1
} elseif ($MailboxType -eq "Shared") {
    $ADmsExchRecipientDisplayType = 0
    $ADmsExchRecipientTypeDetails = 4
}elseif ($MailboxType -eq "Room") {
    $ADmsExchRecipientDisplayType = 7
    $ADmsExchRecipientTypeDetails = 16
}elseif ($MailboxType -eq "Equipment") {
    $ADmsExchRecipientDisplayType = 8
    $ADmsExchRecipientTypeDetails = 32
} else {
    $(Write-Host "Invalid mailbox type." -ForegroundColor Red -NoNewLine) + $(Write-host "Please type 'User', 'Shared', 'Room' or 'Equipment'." -ForegroundColor Yellow)
    exit
}

# Set the mail, ProxyAddresses, msExchRecipientDisplayType, and msExchRecipientTypeDetails fields
Set-ADUser -Identity $SAMAccountName -Replace @{
        mail = $Mail;
        ProxyAddresses = $AllSortedProxyAddresses;
        msExchRecipientDisplayType = $ADmsExchRecipientDisplayType;
        msExchRecipientTypeDetails = $ADmsExchRecipientTypeDetails
    }

# Output a message to indicate the fields have been updated
Write-Host "The mail, ProxyAddresses, msExchRecipientDisplayType, and msExchRecipientTypeDetails fields have been updated for $($UserUPN)." -ForegroundColor Gray

# Enable the remote mailbox
try {
    $EnableRemoteMailBox = Enable-RemoteMailbox -Identity "$Mail" -RemoteRoutingAddress "$ExistingMailOnMicrosoftAddress" -ErrorAction Stop
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
}

if ($ExchangeGuid -eq $null) {
    Write-Host ""
    Write-Warning "ExchangeGuid is missing, wait up to 1 hour and rerun script from part 1"
    Start-Sleep -Seconds 60
    Exit
}

$TryCount = 1

# Set the Exchange GUID for the remote mailbox
$RecipientTypeDetails = Get-RemoteMailbox -Identity $Mail -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

if ($RecipientTypeDetails.RecipientTypeDetails -ne "Remote*Mailbox") {
    $success = $false
    while (-not $success) {
        try {
            Set-RemoteMailbox -Identity $SamAccountName -ExchangeGuid $ExchangeGuid -ErrorAction Stop
            Write-Host "RemoteMailbox is set" -ForegroundColor Green
            $success = $true
        } catch {
            $(Write-Host "An error occurred while setting RemoteMailbox (Try nr $TryCount): " -ForegroundColor Red -NoNewLine) + $(Write-Host "Retrying in 5 seconds..." -ForegroundColor Yellow)
            $TryCount++
            if ($TryCount -gt 5) {
                Write-Warning "RemoteMailbox not set, wait 30 minutes and try script from part 1"
                Exit
            } else {
                Start-Sleep -Seconds 5
            }
        }
    }
} else {
    Write-Host "RemoteMailbox already set" -ForegroundColor Yellow
}

# Get the updated TargetAddress value of the AD user
$UpdatedADUser = Get-ADUser -Identity $SAMAccountName -Properties TargetAddress

# Check if the TargetAddress field is populated
if ([string]::IsNullOrEmpty($UpdatedADUser.TargetAddress)) {
    # Set the TargetAddress field using Set-ADUser
    Set-ADUser -Identity $SAMAccountName -Replace @{TargetAddress = $TargetAddress}
    Write-Host "The TargetAddress field for $($UserUPN) has been set to $($TargetAddress)." -ForegroundColor Gray
} else {
    Write-Host "The TargetAddress field for $($UserUPN) was unchanged." -ForegroundColor Yellow
}
