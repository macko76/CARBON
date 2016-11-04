################################################################################################################################################################  
# Script accepts 4 optional parameters from the command line  
#
# Arguments 
# Realm - Optional - If needed this is used to target only recipients part of the designated domain  
# Username - Optional - G-Administrator Username ID for the tenant we are querying
# Password - Optional - Administrator Username password for the tenant we are querying  
# OutputFile - Optional - The path to the CSV file used for exporting results  
#  
# To run the script
#  
# .\Get-DistributionGroupMembership.ps1 -Username admin@xxxxxx.onmicrosoft.com -Password Password123 -OutputFile c:\reports\DistributionGroupMembers.csv [-DomainFilter "domain.com"]

################################################################################################################################################################  
  
#Accept input parameters  
Param(  

    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $Username,  
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $Password,
    [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $DomainFilter,  
    [Parameter(Position=3, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $OutputFile
)  
  
#Constant Variables  
$arrDLMembers = @{}  
  
#Remove all existing Powershell sessions  
Get-PSSession | Remove-PSSession  
  
#Did they provide creds?  If not, ask them for it. 
if (([string]::IsNullOrEmpty($Username) -eq $false) -and ([string]::IsNullOrEmpty($Password) -eq $false)) 
{ 
    $SecurePassword = ConvertTo-SecureString -AsPlainText $Password -Force      
      
    #Build credentials object  
    $O365Credentials  = New-Object System.Management.Automation.PSCredential $Username, $SecurePassword
} 
else 
{ 
    #Build credentials object  
    $O365Credentials  = Get-Credential
}


Write-Host "Initiating Online session, loading PowerShell commandlets" -ForegroundColor Yellow

  $O365Credentials  = Get-Credential
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

#Get all Distribution Groups from Office 365 

if ([string]::IsNullOrEmpty($DomainFilter) -eq $true) 
{ 
    $objDistributionGroups = Get-DistributionGroup -ResultSize Unlimited 
    Write-Host "No domain filter specified. Located $($objectDistributionGroups.Count) Distribution Groups, processing..." -ForegroundColor Green
}
else {
    $objDistributionGroups = Get-DistributionGroup -ResultSize Unlimited | Where { $_.PrimarySTMPAddress -like "*$DomainFilter" }
    Write-Host "Applied domain filter $(*$DomainFilter). Located $($objDistributionGroups.Count) Distribution Groups, processing..." -ForegroundColor Green
}
  
#Iterate through all groups, one at a time
Foreach ($objDistributionGroup in $objDistributionGroups)  
{      
     
    #$managedBy = $objDistributionGroup.ManagedBy

    Write-host "Processing $($objDistributionGroup.DisplayName)..."
  
    #Get members of this group  
    $objDGMembers = Get-DistributionGroupMember -Identity $($objDistributionGroup.PrimarySmtpAddress) 
      
    Write-host "Found $($objDGMembers.Count) members..." 
      
    #Iterate through each member  
    Foreach ($objMember in $objDGMembers)
    {  
        Out-File -FilePath $OutputFile -InputObject "$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" -Encoding UTF8 -append  
        write-host "`t$($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)" 
    }  
}  
 
#Clean up session
Get-PSSession | Remove-PSSession

#$DL = Get-DistributionGroupMember "Assistants Group" | Select-Object -ExpandProperty Name 
#ForEach ($Member in $DL ) 
#{
#    Add-MailboxPermission -Identity "FL1 Room1"  -User $S -AccessRights FullAccess -InheritanceType All
#}