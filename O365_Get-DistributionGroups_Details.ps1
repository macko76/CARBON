################################################################################################################################################################  
# Arguments 
# Realm - Optional - If needed this is used to target only recipients part of the designated domain  
# Username - Mandatory - G-Administrator Username ID for the tenant we are querying
# Password - Mandatory - Administrator Username password for the tenant we are querying  
# OutputFile - Mandatory - The path to the CSV file used for exporting results  
#  
# To run the script  
#  
# .\Get-DistributionGroup_Details.ps1 -Username [admin@xxxxxx.onmicrosoft.com] [-Password Password123] [-OutputFile c:\reports\distributionGroupDetails.csv] [-DomainFilter "domain.com"]
#  
################################################################################################################################################################  
  
#Accept input parameters  
Param(  

    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $Username, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $Password,
    [Parameter(Position=2, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $OutputFile,
    [Parameter(Position=3, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $DomainFilter  
)
  
cls

#Remove all existing Powershell sessions  
Get-PSSession | Remove-PSSession  
  
#Credentials provided ? If not, ask via a dialog 
if (([string]::IsNullOrEmpty($Username) -eq $false) -and ([string]::IsNullOrEmpty($Password) -eq $false)) 
{ 
    $SecurePassword = ConvertTo-SecureString -AsPlainText $Password -Force      
      
    #Build credentials object  
    $o365Credentials  = New-Object System.Management.Automation.PSCredential $Username, $SecurePassword
} 
else 
{ 
    #Build credentials object  
    $o365Credentials  = Get-Credential
}

connect-msolservice -credential $o365Credentials

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

#Get Distribution Groups from Office 365 
if ([string]::IsNullOrEmpty($DomainFilter) -eq $true) 
{ 
    Get-DistributionGroup -ResultSize Unlimited | Export-Csv -NoTypeInformation -Path $OutputFile
}
else {
   Get-DistributionGroup -ResultSize Unlimited | Where { $_.PrimarySTMPAddress -like '*$($DomainFilter)' } | Export-Csv -NoTypeInformation -Path $OutputFile
}

#Clean up session
Get-PSSession | Remove-PSSession

