
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
# .\O365-Report_AssignedLicenses.ps1 [-DomainFilter "domain.com"] -OutputFolder "C:\reports"
################################################################################################################################################################  
  
#Accept input parameters  
Param(  

    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $DomainFilter,  
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]  
    [string] $OutputFolder
)  

    $licenses =@()

    #Ask for credentials
    $o365Credentials  = Get-Credential

connect-msolservice -credential $o365Credentials

Write-Host "Retrieving Tenant details..." -ForegroundColor Yellow

$MSOLAccountSKU = Get-MsolAccountSku

if ($MSOLAccountSKU -ne $null)
{
    # in a multi-tenant environement 
    $tenantId = $MSOLAccountSKU.AccountObjectId[0]
    $tenantName = $MSOLAccountSKU.AccountName[0]
}

# OutputFile value provided ?
if ([string]::IsNullOrEmpty($OutputFolder) -eq $true)
{ 
    # Determine script location for PowerShell
    $OutputFolder = Split-Path $script:MyInvocation.MyCommand.Path
}
   
   $OutputFile = ( "$OutputFolder\O365-Report_AssignedLicenses", $tenantName, (Get-Date -Format ddMMyyyy-HHmmss) -join "_"), "csv" -join "."

   Write-Host "Setting output file on $OutputFile" -ForegroundColor Green

Get-MsolUser -All | Where {$_.isLicensed -eq $true} | ForEach {

    $msolUserUPN = $_.UserPrincipalName

    ForEach ($License in $_.Licenses) 
    {

      $DisabledOptions = @()

      $License.ServiceStatus | ForEach {


            # Create a new custom object to hold our result.
            $lic = new-object PSObject
        
            $lic | add-member -membertype NoteProperty -name "UserPrincipalName" -Value $msolUserUPN
            $lic | add-member -membertype NoteProperty -name "AccountSkuId" -Value $license.AccountSkuId
            $lic | add-member -membertype NoteProperty -name "ServiceName" -Value $_.ServicePlan.ServiceName
            $lic | add-member -membertype NoteProperty -name "Status" -Value $_.ProvisioningStatus

            $licenses += $lic

            If ($_.ProvisioningStatus -eq "Disabled") { $DisabledOptions += "$($_.ServicePlan.ServiceName)" }

    }

}

}

$licenses | Export-Csv -NoTypeInformation -Delimiter ";" -Path $OutputFile

#Get-MsolUser -All | Where {$_.isLicensed -eq $true} | Select-Object -ExpandProperty Licenses | Select-Object -ExpandProperty ServiceStatus