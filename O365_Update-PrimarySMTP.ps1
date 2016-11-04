################################################################################################################################################################  
# Script accepts 4 optional parameters from the command line  
#
# Arguments 
# SourcePrimaryDomain - Optional - The original domain name targeted for change. By not specifying it means all users will be processed for change of the their Primary SMTP address.
# TargetPrimaryDomain - Optional - The new domain to use as the primary SMTP. 
#  
# To run the script
# .\O365_Update-PrimarySTMP.ps1 [-SourcePrimaryDomain domain.com] [-TargetPrimaryDomain tenant.onmicrosoft.com]

################################################################################################################################################################  
#Accept input parameters  
Param(  

    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]
    [string] $sourcePrimaryDomain,
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $targetPrimaryDomain
)

#Remove all existing Powershell sessions
Get-PSSession | Remove-PSSession  
  
#Did they provide creds?  If not, ask them for it. 

    #Build credentials object  
$o365Credentials  = Get-Credential

Write-Host "Initiating Online session, loading PowerShell commandlets" -ForegroundColor Yellow

#Create remote Powershell session  
$exoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $o365Credentials -Authentication Basic –AllowRedirection          
 
#Import the session  

Import-PSSession $exoSession -AllowClobber | Out-Null

Write-Host "Retrieving Tenant details..." -ForegroundColor Yellow

$MSOLAccountSKU = Get-MsolAccountSku

if ($MSOLAccountSKU -ne $null)
{
    # in a multi-tenant environement 
    $tenantId = $MSOLAccountSKU.AccountObjectId[0]
    $tenantName = $MSOLAccountSKU.AccountName[0] # if the actual targetDomainName is not specificied, this will be used as the new Primary domain
}
  
#$oldPrimaryDomain = "sommet-education.com"
#$newPrimaryDomain = "sommeteducation.onmicrosoft.com"

# Is the SourcePrimaryDomain provided ?
#If NOT = LATER recover current PrimaryAddress and use it to detect the actual source Domain FOR EACH Account
#If YES = Use it as FILTER to load only those accounts matching the source primary domain, avoiding therefore unwanted updates

    if ([string]::IsNullOrEmpty($sourcePrimaryDomain) -eq $true) 
    {
        # WARNING - this loops across the entire Tenant - better ask validation from user before proceeding
        $mailboxes = @(get-mailbox -ResultSize unlimited | where { $_.RecipientTypeDetails -eq "UserMailbox" })  
    }
    else {
    
        $mailboxes = @(get-mailbox -ResultSize unlimited | where { $_.PrimarySmtpAddress -like "*@$sourcePrimaryDomain" -and $_.RecipientTypeDetails -eq "UserMailbox" }) 
    }

    # Is the TargetPrimaryDomain provided ? 
    #If NOT = Rely on the current TENANT name (retrieved earlier) to build the new alias
    #If YES = Simpyl used Use it to update accounts

    if ([string]::IsNullOrEmpty($targetPrimaryDomain) -eq $true) 
    {
        $targetPrimaryDomain = $tenantName, "onmicrosoft.com" -join "."
    }

    Write-Host "Target domain set to $targetPrimaryDomain" -ForegroundColor Yellow

# Start processing in loop each MailBox matching filtering criteria
$mailboxes | ForEach {

    Write-Host "Aliases before change $($_.EmailAddresses)"

    # Will be added as another alias
    $primarySTMPAddress = $_.PrimarySmtpAddress
    
    # Remove current PrimarySTMP from the list of aliases
    $originalAliases = $_.EmailAddresses -replace "SMTP:$primarySTMPAddress",""

    Write-Host "Change 1 - Alias $originalAliases"
 
    $newPrimarySMTP = $($_.Alias, $targetPrimaryDomain -join "@")

    #$originalAliases = $originalAliases -replace "smtp:$newPrimarySMTP",""

    Write-Host "Processing $($_.Name) | Primary SMTP ($($_.PrimarySMTPAddress) > $($_.Alias.ToLower(), $targetPrimaryDomain -join '@') )"

    Set-Mailbox -Identity $($_.Name) -EmailAddresses @{add=$newPrimarySMTP,$primarySTMPAddress}

 #  Set-Mailbox -Identity $($_.Name) -EmailAddresses "SMTP:$newPrimarySMTP",$primarySTMPAddress

    Get-Mailbox -Identity $($_.Name)| select EmailAddresses
}

#Clean up session
Get-PSSession | Remove-PSSession