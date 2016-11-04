################################################################################################################################################################  
# Script accepts 4 optional parameters from the command line  
#
# Arguments 
# DomainFilter - Optional - If needed this is used to target only recipients part of the designated domain  
# OutputFile - Optional - The path to the CSV file used for exporting results  
#  
# To run the script
#  
# .\Get-ResourceMailboxes_AccessRights.ps1 -Username admin@xxxxxx.onmicrosoft.com -Password Password123 -OutputFile c:\reports\ResourceMailboxes_AccessRights.csv [-DomainFilter "domain.com"]

################################################################################################################################################################  

#Accept input parameters  
Param(  
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $DomainFilter,  
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $OutputFile
)

#Remove all existing Powershell sessions  
Get-PSSession | Remove-PSSession  
  
#Build credentials object  
$O365Credentials  = Get-Credential

Write-Host "Initiating Online session, loading PowerShell commandlets" -ForegroundColor Yellow

connect-msolservice -credential $o365Credentials
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
    $tenantName = $MSOLAccountSKU.AccountName[0]
}

# OutputFile value provided ?
if ([string]::IsNullOrEmpty($OutputFile) -eq $true) 
{ 
    # Determine script location for PowerShell
    $ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

#    $OutputFile = ("D:\PS-Scripts\Sommet-Education\O365_DistributionGroup_Membership", (Get-Date -Format d) -join "_"), "csv" -join "."
    $OutputFile = ( "$scriptDir\O365_DistributionGroup_Details", $tenantName, (Get-Date -Format ddMMyyyy-HHmmss) -join "_"), "csv" -join "."

    #Write-Host "Output file not specified, setting output to $OutputFile" -ForegroundColor Yellow
} 

    Write-Host "Setting output file on $OutputFile" -ForegroundColor Green
    #Prepare Output file with headers  
    Out-File -FilePath $OutputFile -InputObject "Name, Email, MemberEmail, MemberRecipientType" -Encoding UTF8

    if ([string]::IsNullOrEmpty($DomainFilter) -eq $true) 
    { 
        Get-Mailbox -ResultSize Unlimited | Where {$_.RecipientTypeDetails -eq "RoomMailBox" -or $_.RecipientTypeDetails -eq "EquipmentMailBox"} | Export-Csv -NoTypeInformation -Encoding Default -Delimiter ";" $OutputFile
    }
    else 
    {
        Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySMTPAddress -like "*$DomainFilter" -and ($_.RecipientTypeDetails -eq "RoomMailBox" -or $_.RecipientTypeDetails -eq "EquipmentMailBox") } | Export-Csv -NoTypeInformation -Encoding Default -Delimiter ";" $OutputFile
    }

#Same as above, while limiting properties
#Get-Mailbox -Filter '(RecipientTypeDetails -eq "RoomMailBox" -or RecipientTypeDetails -eq "EquipmentMailBox")' | Select Name, ResourceType, PrimarySMTPAddress, Alias, ResourceDelegates, GrantSendOnBehalfTo, RejectMessagesFrom, RejectMessagesFromDLMembers,RejectMessagesFromSendersOrMembers,ModerationEnabled,ModeratedBy, SendModerationNotifications, DeleteComments, DeleteSubject | Export-Csv -NoTypeInformation -Encoding Default -Delimiter ";" D:\PS-Scripts\Sommet-Education\O365-ResourceMailbox_Details.csv

# Declare an array to collect our result objects
$resPerms =@()

Get-Mailbox -ResultSize Unlimited | Where {$_.PrimarySMTPAddress -like "*$realm" -and ($_.RecipientTypeDetails -eq "RoomMailBox" -or $_.RecipientTypeDetails -eq "EquipmentMailBox") } | ForEach {

    $resName = $_.Name
    $resType = $_.ResourceType

    Get-MailBoxFolderPermission  "$($_.Name):\Calendar"| Where {$_.user.toString() -ne "Default" }  | Select FolderName,User,AccessRights | ForEach {
    
        # Create a new custom object to hold our result.
        $perm = new-object PSObject
        
        # Add our data to $contactObject as attributes using the add-member commandlet
        $perm | add-member -membertype NoteProperty -name "ResourceType" -Value $resType
        $perm | add-member -membertype NoteProperty -name "Resource" -Value $resName
        $perm | add-member -membertype NoteProperty -name "FolderName" -Value $_.FolderName
        $perm | add-member -membertype NoteProperty -name "User" -Value $_.User
        $perm | add-member -membertype NoteProperty -name "AccessRights" -Value $_.AccessRights

        # Save the current $contactObject by appending it to $resultsArray ( += means append a new element to ‘me’)
        $resPerms += $perm
    }

    Get-Mailbox | Get-MailboxPermission | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -NoTypeInformation d:\ps-scripts\mailboxpermissions.csv


}

$resPerms | Export-Csv -NoTypeInformation -Delimiter ";" -Path D:\PS-Scripts\Sommet-Education\O365-ResourceMailbox_FolderPermissions.csv

#Clean up session
Get-PSSession | Remove-PSSession