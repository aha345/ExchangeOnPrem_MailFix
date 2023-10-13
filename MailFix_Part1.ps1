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

# Check if AzureAD credentials are already stored
if (-not $Credential) {
    # Prompt the user for AzureAD credentials
    $Credential = Get-Credential -Message "Enter your AzureAD credentials"
}

# Connect to AzureAD with the provided credentials
Connect-AzureAD -Credential $Credential | Out-Null

# Connect to Exchange Online with the provided credentials
Import-Module ExchangeOnlineManagement
$VerbosePreference = "SilentlyContinue"
Connect-ExchangeOnline -Credential $Credential | Out-Null

Get-Variable |
    Where-Object { $_.Name -ne "Credential" -and $_.Options -ne "ReadOnly" -and $_.Options -ne "AllScope" -and $_.Options -ne "Constant" } |
    Remove-Variable -ErrorAction SilentlyContinue

function Prompt-UserForPrimaryAddress {
    param(
        [Parameter(Mandatory = $true)] [string[]] $Addresses,
        [Parameter(Mandatory = $true)] [string] $AddressType
    )

    $menuItems = $Addresses | ForEach-Object { New-Object PSObject -Property @{ Label = $_ } }
    $title = "Please select the primary $AddressType to keep:"
    $menu = $($menuItems | ForEach-Object { "{0}. {1}" -f ($_.ReadCount, $_.Label) }) -join "`n"

    do {
        $choice = Read-Host -Prompt "$title`n$menu`n`nPlease enter the number:"
        $selectedAddress = $menuItems[$choice - 1].Label
    } until ($selectedAddress)

    return $selectedAddress
}

# Clear the console window
Clear-Host

# What AzureAD Tenant is connected
$AzureADTenant = ((Get-AzureADTenantDetail).VerifiedDomains | Where-Object { $_.Name -like "*.onmicrosoft.com" -and $_.Name -notlike "*.mail.onmicrosoft.com" } | Select-Object -ExpandProperty Name) -replace '.onmicrosoft.com', ''
$(Write-Host "Connected to AzureAD Tenant: " -ForegroundColor Cyan -NoNewLine) + $(Write-Host "$AzureADTenant" -ForegroundColor Green)

# Prompt the user for the username of the AD user to modify
$Username = $(Write-Host "Enter the username of the AD user to modify " -ForegroundColor Magenta -NoNewLine) + $(Write-Host "(Can be Name, UserPrincipalName, Email, or SamAccountName)" -ForegroundColor Yellow -NoNewLine) + $(Write-Host ": " -ForegroundColor Magenta -NoNewLine; Read-Host)

# Find the AD user based on the provided username and retrieve all properties
$ADUser = Get-ADUser -Filter {
    (Name -eq $Username) -or
    (UserPrincipalName -eq $Username) -or
    (EmailAddress -eq $Username) -or
    (SamAccountName -eq $Username)
} -Properties *

# Get the UPN of the AD user
$UPN = $ADUser.UserPrincipalName

# Get the specified AD attributes for the user
$Name = $ADUser.Name
$SamAccountName = $ADUser.SamAccountName
$Mail = $ADUser.Mail
$TargetAddress = $ADUser.TargetAddress
$msexchRecipientDisplayType = $ADUser.msexchRecipientDisplayType
$TypeDetails = $ADUser.msExchRecipientTypeDetails
$ADProxyAddresses = $ADUser.ProxyAddresses

# Find the user in Azure AD based on the UPN and retrieve all properties
$AzureADUser = Get-AzureADUser -ObjectId $UPN -ErrorAction SilentlyContinue

# Find the user in Exchange Online based on the UPN and retrieve all properties
$EXOUser = Get-Mailbox -Identity $UPN -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# Get the proxy addresses of the AzureAD user
if ($AzureADUser) {
    $AzureADProxyAddresses = $AzureADUser.ProxyAddresses
}

# Get the proxy addresses of the Exchange Online user
if ($EXOUser) {
    $EXOProxyAddresses = $EXOUser.EmailAddresses
    $ExchangeGuid = $EXOUser.ExchangeGuid
}

# Print the current date and time
Write-Host "Current date and time: $(Get-Date)"

# Write user info
$(Write-Host "Name:" -ForegroundColor Cyan) + $(Write-Host "$Name" -ForegroundColor DarkGray)
$(Write-Host "Mail:" -ForegroundColor Cyan) + $(Write-Host "$Mail" -ForegroundColor DarkGray)
$(Write-Host "SamAccountName:" -ForegroundColor Cyan) + $(Write-Host "$SamAccountName" -ForegroundColor DarkGray)
$(Write-Host "TargetAddress:" -ForegroundColor Cyan) + $(Write-Host "$TargetAddress" -ForegroundColor DarkGray)

# Print all ProxyAddresses with duplicates removed
$AllProxyAddresses = @($ADProxyAddresses) + @($AzureADProxyAddresses) + @($EXOProxyAddresses)
$AllProxyAddresses = $AllProxyAddresses | Sort-Object | Select-Object -Unique
Write-Host "All ProxyAddresses:" -ForegroundColor Cyan
foreach ($Address in $AllProxyAddresses) {
    Write-Host $Address -ForegroundColor DarkGray
}

# Check for duplicate all caps SMTP or X500 values
$SMTPPrimaryAddresses = $AllProxyAddresses | Where-Object { $_ -clike "SMTP:*" }
if ($SMTPPrimaryAddresses.Count -gt 1) {
    Write-Warning "There are multiple primary SMTP values in the ProxyAddresses:"
    $SMTPPrimaryAddresses | ForEach-Object {
        Write-Warning $_
    }

    $primarySMTPToKeep = Prompt-UserForPrimaryAddress -Addresses $SMTPPrimaryAddresses -AddressType "SMTP"
    $AllProxyAddresses = $AllProxyAddresses | ForEach-Object {
        if ($_ -clike "SMTP:*" -and $_ -cne $primarySMTPToKeep) {
            "smtp:$($_.Substring(5))"
        } else {
            $_
        }
    }
}

$X500PrimaryAddresses = $AllProxyAddresses | Where-Object { $_ -clike "X500:*" }
if ($X500PrimaryAddresses.Count -gt 1) {
    Write-Warning "There are multiple primary X500 values in the ProxyAddresses:"
    $X500PrimaryAddresses | ForEach-Object {
        Write-Warning $_
    }

    $primaryX500ToKeep = Prompt-UserForPrimaryAddress -Addresses $X500PrimaryAddresses -AddressType "X500"
    $AllProxyAddresses = $AllProxyAddresses | ForEach-Object {
        if ($_ -clike "X500:*" -and $_ -cne $primaryX500ToKeep) {
            "x500:$($_.Substring(5))"
        } else {
            $_
        }
    }
}

# Add onmicrosoft addresses if they don't exist
if ($AzureADTenant) {
    $onmicrosoftAddress = "$SamAccountName@$AzureADTenant.onmicrosoft.com"
    $mailOnmicrosoftAddress = "$SamAccountName@$AzureADTenant.mail.onmicrosoft.com"
    $onmicrosoftExists = $AllProxyAddresses | Where-Object { $_ -like "*$AzureADTenant.onmicrosoft.com" }
    $mailOnmicrosoftExists = $AllProxyAddresses | Where-Object { $_ -like "*$AzureADTenant.mail.onmicrosoft.com" }
    
    if (-not $onmicrosoftExists) {
        $AllProxyAddresses += $onmicrosoftAddress
    }
    if (-not $mailOnmicrosoftExists) {
        
        # Create a variable for the existing mail.onmicrosoft.com address
        if ($MailOnMicrosoftAddress -notlike "smtp:*$AzureADTenant.mail.onmicrosoft.com") {
            $SMTPMailOnMicrosoftAddress = $MailOnMicrosoftAddress.Insert(0,"smtp:")
            $AllProxyAddresses += $SMTPMailOnmicrosoftAddress
        }
    
        else {
            $AllProxyAddresses += $mailOnmicrosoftAddress
        }
    }
}

# Add mail.onmicrosoft.com address to ExistingMailOnMicrosoftAddress
if ($MailOnMicrosoftAddress -like "smtp:*$AzureADTenant.mail.onmicrosoft.com") {

    $ExistingMailOnMicrosoftAddress = $MailOnMicrosoftAddress -Replace '[smtp:]',''

} else {

    $ExistingMailOnMicrosoftAddress = $MailOnMicrosoftAddress

}

# Remove duplicates
$AllProxyAddresses = $AllProxyAddresses | Select-Object -Unique

# Create new variable for all sorted ProxyAddresses with duplicates removed
$AllSortedProxyAddresses = $AllProxyAddresses | Sort-Object @{
    Expression = {
        if ($_ -like "SMTP:*" -or $_ -like "X500:*") {
            "0" + $_.ToUpper()
        } else {
            "1" + $_
        }
    }
} | Select-Object -Unique


# Create an object with the values to export
$ExportObject = [PSCustomObject]@{
    Name = $Name
    UPN = $UPN
    SAMAccountName = $SamAccountName
    Mail = $Mail
    TargetAddress = $TargetAddress
    msexchRecipientDisplayType = $msexchRecipientDisplayType
    msExchRecipientTypeDetails = $TypeDetails
    ExchangeGuid = $ExchangeGuid
    AllSortedProxyAddresses = $AllSortedProxyAddresses -join ', '
    ExistingMailOnMicrosoftAddress = $ExistingMailOnMicrosoftAddress
}

# Specify the path to the output CSV file
$OutputFileName = "$($env:USERPROFILE)\Desktop\$(Get-Date -Format 'yyyyMMdd')-$($ADUser.Name).csv"

# Export the object to a CSV file
$ExportObject | Export-Csv -Path $OutputFileName -Encoding UTF8 -NoTypeInformation

# Get the current user's distinguished name (DN) 
$currentuser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$username = $currentuser.split('\')[-1]

# Check if the user is a member of the "Organization Management" group 
$orgmanager = Get-ADGroupMember -Identity "Organization Management" -Recursive | where {$_.SamAccountName -eq $username} 
 
if ($orgmanager) { 
} else { 
    Write-Host "$username is not a member of the Organization Management group" -ForegroundColor Red
}
